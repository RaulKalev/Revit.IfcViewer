using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System.Collections.Generic;
using Media3D = System.Windows.Media.Media3D;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Manages a single axis-aligned section plane applied to every mesh in the scene.
    ///
    /// Meshes are loaded as plain <see cref="MeshGeometryModel3D"/> (cheap Blinn-Phong
    /// shader). When the section plane is enabled for the first time, this manager
    /// swaps each registered mesh to a <see cref="CrossSectionMeshGeometryModel3D"/>
    /// in-place inside its parent group, then applies the clip plane. When disabled,
    /// it swaps them back to plain meshes — restoring the cheaper shader path.
    ///
    /// Also owns a semi-transparent blue quad that visualises the plane position.
    /// Call <see cref="AttachVisual"/> once on the UI thread, then update
    /// <see cref="Axis"/> / <see cref="Offset"/> / <see cref="Enabled"/> freely.
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
        private bool        _enabled;
        private SectionAxis _axis   = SectionAxis.Z;
        private float       _offset = 0f;

        // ── Axis range hints (updated by caller after loading geometry) ───────
        public float MinBound { get; set; } = -50f;
        public float MaxBound { get; set; } =  50f;

        // ── Visual quad ──────────────────────────────────────────────────────
        private MeshGeometryModel3D _planeVisual;
        private GroupModel3D        _visualParent;
        private const float QuadHalfSize = 500f;

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

        public SectionAxis Axis
        {
            get => _axis;
            set { _axis = value; if (_enabled) ApplyAll(); UpdateVisual(); }
        }

        /// <summary>Signed distance from origin along the chosen axis (metres).</summary>
        public float Offset
        {
            get => _offset;
            set { _offset = value; if (_enabled) ApplyAll(); UpdateVisual(); }
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
                Parent        = parent,
                LiveMesh      = mesh,
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
            Vector3 normal     = AxisNormal(_axis);
            Vector3 planePoint = normal * _offset;
            float d            = -Vector3.Dot(normal, planePoint);

            mesh.Plane1           = new Plane(normal, d);
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

            var xform = new Media3D.Transform3DGroup();

            switch (_axis)
            {
                case SectionAxis.X:
                    xform.Children.Add(new Media3D.RotateTransform3D(
                        new Media3D.AxisAngleRotation3D(new Media3D.Vector3D(0, 0, 1), 90)));
                    xform.Children.Add(new Media3D.TranslateTransform3D(_offset, 0, 0));
                    break;

                case SectionAxis.Y:
                    xform.Children.Add(new Media3D.TranslateTransform3D(0, _offset, 0));
                    break;

                default: // Z
                    xform.Children.Add(new Media3D.RotateTransform3D(
                        new Media3D.AxisAngleRotation3D(new Media3D.Vector3D(1, 0, 0), 90)));
                    xform.Children.Add(new Media3D.TranslateTransform3D(0, 0, _offset));
                    break;
            }

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

        private static Vector3 AxisNormal(SectionAxis axis)
        {
            switch (axis)
            {
                case SectionAxis.X:  return new Vector3(1, 0, 0);
                case SectionAxis.Y:  return new Vector3(0, 1, 0);
                default:             return new Vector3(0, 0, 1);
            }
        }
    }

    public enum SectionAxis { X, Y, Z }
}
