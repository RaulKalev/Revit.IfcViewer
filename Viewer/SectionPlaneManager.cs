using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System.Collections.Generic;
using Media3D = System.Windows.Media.Media3D;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Manages a single arbitrary section plane applied to every mesh in the scene.
    ///
    /// The plane is defined by a world-space normal and a point that lies on the plane.
    /// Call <see cref="SetPlane"/> whenever the user picks a new face.
    ///
    /// Meshes are loaded as plain <see cref="MeshGeometryModel3D"/> (cheap Blinn-Phong
    /// shader). When the section plane is enabled, this manager swaps each registered mesh
    /// to a <see cref="CrossSectionMeshGeometryModel3D"/> in-place inside its parent group,
    /// then applies the clip plane. When disabled it swaps them back to plain meshes.
    ///
    /// Also owns a semi-transparent blue quad that visualises the plane position.
    /// Call <see cref="AttachVisual"/> once on the UI thread.
    /// </summary>
    public sealed class SectionPlaneManager
    {
        // ── Registered mesh entries ───────────────────────────────────────────
        private sealed class MeshEntry
        {
            public GroupModel3D        Parent;
            public MeshGeometryModel3D LiveMesh;
            public bool                IsCrossSection;
        }

        private readonly List<MeshEntry> _entries = new List<MeshEntry>();

        // ── Backing fields ────────────────────────────────────────────────────
        private bool    _enabled;
        private Vector3 _normal      = new Vector3(0, 0, 1);   // default: Z-up
        private Vector3 _pointOnPlane = Vector3.Zero;

        // ── Visual quad ──────────────────────────────────────────────────────
        private MeshGeometryModel3D _planeVisual;
        private GroupModel3D        _visualParent;
        private const float         QuadHalfSize = 500f;

        /// <summary>
        /// The semi-transparent plane quad mesh, or <c>null</c> before
        /// <see cref="AttachVisual"/> is called.  Exposed so callers can skip it
        /// when collecting scene meshes (e.g. wireframe / outline helpers).
        /// </summary>
        public MeshGeometryModel3D PlaneVisual => _planeVisual;

        // ── Public API ────────────────────────────────────────────────────────

        public bool Enabled
        {
            get => _enabled;
            set
            {
                if (_enabled == value) return;
                _enabled = value;
                if (_enabled)
                    UpgradeAll();   // plain → CrossSection + apply clip
                else
                    DowngradeAll(); // CrossSection → plain (removes expensive shader)
                UpdateVisual();
            }
        }

        /// <summary>
        /// Define the section plane from a world-space face normal and a point on that face.
        /// The plane clips everything on the side the normal points towards.
        /// </summary>
        public void SetPlane(Vector3 normal, Vector3 pointOnPlane)
        {
            _normal       = Vector3.Normalize(normal);
            _pointOnPlane = pointOnPlane;
            if (_enabled) ApplyAll();
            UpdateVisual();
        }

        // ── Visual attachment ─────────────────────────────────────────────────

        /// <summary>
        /// Create the blue transparent plane visual and add it to <paramref name="parent"/>.
        /// Must be called on the WPF UI thread.
        /// </summary>
        public void AttachVisual(GroupModel3D parent)
        {
            if (_planeVisual != null) return;
            _visualParent = parent;

            var mesh = new MeshGeometry3D();
            BuildQuad(mesh, QuadHalfSize);

            var mat = new PhongMaterial
            {
                DiffuseColor      = new Color4(0.20f, 0.55f, 1.00f, 0.30f),
                AmbientColor      = new Color4(0.10f, 0.30f, 0.60f, 0.15f),
                SpecularColor     = new Color4(0f, 0f, 0f, 0f),
                SpecularShininess = 1f,
            };

            _planeVisual = new MeshGeometryModel3D
            {
                Geometry      = mesh,
                Material      = mat,
                IsTransparent = true,
                DepthBias     = -100,
            };

            parent.Children.Add(_planeVisual);
            UpdateVisual();
        }

        /// <summary>Remove and discard the plane visual from the scene.</summary>
        public void DetachVisual()
        {
            if (_planeVisual == null) return;
            _visualParent?.Children.Remove(_planeVisual);
            _planeVisual  = null;
            _visualParent = null;
        }

        // ── Mesh registration ─────────────────────────────────────────────────

        /// <summary>
        /// Register a mesh so it participates in section plane operations.
        /// The <paramref name="parent"/> must be the direct parent group of the mesh.
        /// </summary>
        public void Register(MeshGeometryModel3D mesh, GroupModel3D parent)
        {
            if (mesh == null || parent == null) return;
            foreach (var e in _entries)
                if (ReferenceEquals(e.LiveMesh, mesh)) return; // already registered

            var entry = new MeshEntry
            {
                Parent         = parent,
                LiveMesh       = mesh,
                IsCrossSection = mesh is CrossSectionMeshGeometryModel3D
            };
            _entries.Add(entry);

            // If already enabled, upgrade this new mesh immediately
            if (_enabled) UpgradeEntry(entry);
        }

        /// <summary>Unregister a mesh (e.g. when a model is removed).</summary>
        public void Unregister(MeshGeometryModel3D mesh)
        {
            for (int i = _entries.Count - 1; i >= 0; i--)
            {
                if (ReferenceEquals(_entries[i].LiveMesh, mesh))
                {
                    DowngradeEntry(_entries[i]);
                    _entries.RemoveAt(i);
                    return;
                }
            }
        }

        /// <summary>
        /// Register all <see cref="MeshGeometryModel3D"/> children of a group (recursive).
        /// Call this after inserting a loaded model's SceneGroup into the live scene.
        /// </summary>
        public void RegisterGroup(GroupModel3D group)
        {
            if (group == null) return;
            RegisterGroupRecursive(group);
        }

        private void RegisterGroupRecursive(GroupModel3D group)
        {
            foreach (var child in group.Children)
            {
                if (child is MeshGeometryModel3D mesh)
                    Register(mesh, group);
                else if (child is GroupModel3D sub)
                    RegisterGroupRecursive(sub);
            }
        }

        /// <summary>Unregister all meshes in a group (recursive).</summary>
        public void UnregisterGroup(GroupModel3D group)
        {
            if (group == null) return;
            foreach (var child in group.Children)
            {
                if (child is MeshGeometryModel3D mesh)
                    Unregister(mesh);
                else if (child is GroupModel3D sub)
                    UnregisterGroup(sub);
            }
        }

        /// <summary>
        /// Removes all entries whose <c>LiveMesh</c> is no longer a child of its
        /// registered parent group. Call this after an incremental scene patch to clean
        /// up stale references for meshes that were silently removed by
        /// <c>RevitExporter.PatchScene</c>.
        /// </summary>
        public void PruneDetachedEntries()
        {
            for (int i = _entries.Count - 1; i >= 0; i--)
            {
                var e = _entries[i];
                if (!e.Parent.Children.Contains(e.LiveMesh))
                    _entries.RemoveAt(i);
            }
        }

        // ── Upgrade / Downgrade ───────────────────────────────────────────────

        private void UpgradeAll()
        {
            foreach (var entry in _entries)
                UpgradeEntry(entry);
        }

        private void DowngradeAll()
        {
            foreach (var entry in _entries)
                DowngradeEntry(entry);
        }

        private void UpgradeEntry(MeshEntry entry)
        {
            if (entry.IsCrossSection)
            {
                // Already cross-section, just re-apply the plane
                ApplyCross((CrossSectionMeshGeometryModel3D)entry.LiveMesh);
                return;
            }

            var plain = entry.LiveMesh;
            var cs = new CrossSectionMeshGeometryModel3D
            {
                Geometry      = plain.Geometry,
                Material      = plain.Material,
                IsTransparent = plain.IsTransparent,
                Transform     = plain.Transform,
            };
            ApplyCross(cs);

            int idx = entry.Parent.Children.IndexOf(plain);
            if (idx >= 0)
                entry.Parent.Children[idx] = cs;
            else
                entry.Parent.Children.Add(cs);

            entry.LiveMesh      = cs;
            entry.IsCrossSection = true;
        }

        private void DowngradeEntry(MeshEntry entry)
        {
            if (!entry.IsCrossSection) return;

            var cs = (CrossSectionMeshGeometryModel3D)entry.LiveMesh;
            cs.EnablePlane1 = false;

            var plain = new MeshGeometryModel3D
            {
                Geometry      = cs.Geometry,
                Material      = cs.Material,
                IsTransparent = cs.IsTransparent,
                Transform     = cs.Transform,
            };

            int idx = entry.Parent.Children.IndexOf(cs);
            if (idx >= 0)
                entry.Parent.Children[idx] = plain;
            else
                entry.Parent.Children.Add(plain);

            entry.LiveMesh      = plain;
            entry.IsCrossSection = false;
        }

        // ── Internal — clipping ───────────────────────────────────────────────

        private void ApplyAll()
        {
            foreach (var entry in _entries)
            {
                if (entry.IsCrossSection)
                    ApplyCross((CrossSectionMeshGeometryModel3D)entry.LiveMesh);
            }
        }

        private void ApplyCross(CrossSectionMeshGeometryModel3D mesh)
        {
            // Hessian-normal form: Plane(n, d) where d = -dot(n, pointOnPlane)
            float d = -Vector3.Dot(_normal, _pointOnPlane);
            mesh.Plane1           = new Plane(_normal, d);
            mesh.EnablePlane1     = true;
            mesh.CuttingOperation = CuttingOperation.Subtract;
        }

        // ── Internal — visual ─────────────────────────────────────────────────

        private void UpdateVisual()
        {
            if (_planeVisual == null) return;

            _planeVisual.Visibility = _enabled
                ? System.Windows.Visibility.Visible
                : System.Windows.Visibility.Collapsed;

            if (!_enabled) return;

            // Build a rotation that brings the quad's Y-up normal to align with _normal.
            // The base quad lies in the XZ plane with Y = up (see BuildQuad).
            var yUp   = new Vector3(0, 1, 0);
            var n     = _normal;
            var cross = Vector3.Cross(yUp, n);
            double angle;
            Media3D.Vector3D axis;

            if (cross.LengthSquared() < 1e-6f)
            {
                // Parallel or anti-parallel to Y
                axis  = new Media3D.Vector3D(1, 0, 0);
                angle = Vector3.Dot(yUp, n) > 0 ? 0 : 180;
            }
            else
            {
                axis  = new Media3D.Vector3D(cross.X, cross.Y, cross.Z);
                angle = System.Math.Acos(
                    System.Math.Max(-1.0, System.Math.Min(1.0,
                        Vector3.Dot(yUp, n)))) * (180.0 / System.Math.PI);
            }

            var xform = new Media3D.Transform3DGroup();
            xform.Children.Add(new Media3D.RotateTransform3D(
                new Media3D.AxisAngleRotation3D(axis, angle)));
            xform.Children.Add(new Media3D.TranslateTransform3D(
                _pointOnPlane.X, _pointOnPlane.Y, _pointOnPlane.Z));

            _planeVisual.Transform = xform;
        }

        // ── Geometry helpers ──────────────────────────────────────────────────

        private static void BuildQuad(MeshGeometry3D mesh, float half)
        {
            mesh.Positions = new Vector3Collection
            {
                new Vector3(-half, 0,  half),
                new Vector3( half, 0,  half),
                new Vector3( half, 0, -half),
                new Vector3(-half, 0, -half),
            };
            mesh.Normals = new Vector3Collection
            {
                new Vector3(0, 1, 0), new Vector3(0, 1, 0),
                new Vector3(0, 1, 0), new Vector3(0, 1, 0),
            };
            mesh.Indices = new IntCollection
            {
                0, 1, 2,  0, 2, 3,   // front
                0, 2, 1,  0, 3, 2,   // back
            };
        }
    }
}
