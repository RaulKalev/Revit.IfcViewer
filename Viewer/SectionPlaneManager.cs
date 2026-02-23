using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;
using System.Windows.Media;
using Media3D = System.Windows.Media.Media3D;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Manages a single axis-aligned section plane applied to every
    /// <see cref="CrossSectionMeshGeometryModel3D"/> in the scene.
    ///
    /// Also owns a semi-transparent blue quad mesh that visualises the plane
    /// position. Call <see cref="AttachVisual"/> once (on the UI thread) to
    /// insert the visual into the scene, then update <see cref="Axis"/> /
    /// <see cref="Offset"/> / <see cref="Enabled"/> freely.
    /// </summary>
    public sealed class SectionPlaneManager
    {
        // ── Registered meshes ─────────────────────────────────────────────────
        private readonly List<CrossSectionMeshGeometryModel3D> _meshes
            = new List<CrossSectionMeshGeometryModel3D>();

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

        // Half-size of the visual quad in metres — large enough to cover any scene
        private const float QuadHalfSize = 500f;

        // ── Public API ────────────────────────────────────────────────────────

        public bool Enabled
        {
            get => _enabled;
            set { _enabled = value; ApplyAll(); UpdateVisual(); }
        }

        public SectionAxis Axis
        {
            get => _axis;
            set { _axis = value; ApplyAll(); UpdateVisual(); }
        }

        /// <summary>Signed distance from origin along the chosen axis (metres).</summary>
        public float Offset
        {
            get => _offset;
            set { _offset = value; ApplyAll(); UpdateVisual(); }
        }

        // ── Visual attachment ─────────────────────────────────────────────────

        /// <summary>
        /// Create the blue transparent plane visual and add it to <paramref name="parent"/>.
        /// Must be called on the WPF UI thread (creates DependencyObjects).
        /// </summary>
        public void AttachVisual(GroupModel3D parent)
        {
            if (_planeVisual != null) return; // already attached
            _visualParent = parent;

            // Build a flat unit quad in the XZ plane, centred at origin.
            // We'll rotate/translate it via Transform to match the current axis+offset.
            var mesh = new MeshGeometry3D();
            BuildQuad(mesh, QuadHalfSize);

            var mat = new PhongMaterial
            {
                DiffuseColor      = new Color4(0.20f, 0.55f, 1.00f, 0.30f), // blue, 30% opaque
                AmbientColor      = new Color4(0.10f, 0.30f, 0.60f, 0.15f),
                SpecularColor     = new Color4(0f, 0f, 0f, 0f),
                SpecularShininess = 1f,
            };

            _planeVisual = new MeshGeometryModel3D
            {
                Geometry      = mesh,
                Material      = mat,
                IsTransparent = true,
                // Depth bias so the quad doesn't z-fight with nearby geometry
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

        /// <summary>Register a mesh so the section plane is applied to it.</summary>
        public void Register(CrossSectionMeshGeometryModel3D mesh)
        {
            if (mesh == null || _meshes.Contains(mesh)) return;
            _meshes.Add(mesh);
            Apply(mesh);
        }

        /// <summary>Unregister a mesh (e.g. when a model is removed).</summary>
        public void Unregister(CrossSectionMeshGeometryModel3D mesh)
        {
            if (mesh == null) return;
            _meshes.Remove(mesh);
            mesh.EnablePlane1 = false;
        }

        /// <summary>Register all CrossSectionMeshGeometryModel3D children of a group (recursive).</summary>
        public void RegisterGroup(GroupModel3D group)
        {
            if (group == null) return;
            foreach (var child in group.Children)
            {
                if (child is CrossSectionMeshGeometryModel3D csm)
                    Register(csm);
                else if (child is GroupModel3D sub)
                    RegisterGroup(sub);
            }
        }

        /// <summary>Unregister all meshes in a group (recursive).</summary>
        public void UnregisterGroup(GroupModel3D group)
        {
            if (group == null) return;
            foreach (var child in group.Children)
            {
                if (child is CrossSectionMeshGeometryModel3D csm)
                    Unregister(csm);
                else if (child is GroupModel3D sub)
                    UnregisterGroup(sub);
            }
        }

        // ── Internal — clipping ───────────────────────────────────────────────

        private void ApplyAll()
        {
            foreach (var mesh in _meshes)
                Apply(mesh);
        }

        private void Apply(CrossSectionMeshGeometryModel3D mesh)
        {
            if (!_enabled)
            {
                mesh.EnablePlane1 = false;
                return;
            }

            // Build a SharpDX.Plane: dot(normal, p) + d = 0
            // CrossSection clips the side where dot(normal,p)+d > 0 (Subtract).
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

            // The quad geometry lies in the XZ plane (Y = 0) by default.
            // Rotate it so it becomes perpendicular to the chosen axis, then
            // translate it to the offset position.
            Media3D.Transform3DGroup xform = new Media3D.Transform3DGroup();

            switch (_axis)
            {
                case SectionAxis.X:
                    // Rotate 90° around Z so the quad normal points along +X
                    xform.Children.Add(new Media3D.RotateTransform3D(
                        new Media3D.AxisAngleRotation3D(
                            new Media3D.Vector3D(0, 0, 1), 90)));
                    xform.Children.Add(new Media3D.TranslateTransform3D(_offset, 0, 0));
                    break;

                case SectionAxis.Y:
                    // Quad already lies in XZ (normal = Y) — just translate
                    xform.Children.Add(new Media3D.TranslateTransform3D(0, _offset, 0));
                    break;

                default: // Z
                    // Rotate 90° around X so the quad normal points along +Z
                    xform.Children.Add(new Media3D.RotateTransform3D(
                        new Media3D.AxisAngleRotation3D(
                            new Media3D.Vector3D(1, 0, 0), 90)));
                    xform.Children.Add(new Media3D.TranslateTransform3D(0, 0, _offset));
                    break;
            }

            _planeVisual.Transform = xform;
        }

        // ── Geometry helpers ──────────────────────────────────────────────────

        /// <summary>Build a double-sided flat quad in the XZ plane, centred at origin.</summary>
        private static void BuildQuad(MeshGeometry3D mesh, float half)
        {
            // 4 corners in XZ plane (Y = 0)
            var positions = new Vector3Collection
            {
                new Vector3(-half, 0,  half),
                new Vector3( half, 0,  half),
                new Vector3( half, 0, -half),
                new Vector3(-half, 0, -half),
            };

            // Normals pointing up (+Y) — duplicated for the back face below
            var normals = new Vector3Collection
            {
                new Vector3(0, 1, 0),
                new Vector3(0, 1, 0),
                new Vector3(0, 1, 0),
                new Vector3(0, 1, 0),
            };

            // Front face (winding: CCW from above = +Y normal)
            // Back face  (winding reversed so both sides are visible)
            var indices = new IntCollection
            {
                0, 1, 2,   0, 2, 3,   // front (+Y side)
                0, 2, 1,   0, 3, 2,   // back  (-Y side)
            };

            mesh.Positions = positions;
            mesh.Normals   = normals;
            mesh.Indices   = indices;
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
