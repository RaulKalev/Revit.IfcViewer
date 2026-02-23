using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Manages a single axis-aligned section plane applied to every
    /// <see cref="CrossSectionMeshGeometryModel3D"/> in the scene.
    ///
    /// The plane is always Plane1 on each mesh.  Normal and offset are
    /// derived from <see cref="Axis"/> and <see cref="Offset"/>.
    ///
    /// Usage:
    ///   sectionMgr.Axis    = SectionAxis.Z;
    ///   sectionMgr.Offset  = 2.5f;     // metres along axis
    ///   sectionMgr.Enabled = true;
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

        // ── Public API ────────────────────────────────────────────────────────

        public bool Enabled
        {
            get => _enabled;
            set { _enabled = value; ApplyAll(); }
        }

        public SectionAxis Axis
        {
            get => _axis;
            set { _axis = value; ApplyAll(); }
        }

        /// <summary>Signed distance from origin along the chosen axis (metres).</summary>
        public float Offset
        {
            get => _offset;
            set { _offset = value; ApplyAll(); }
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
            // Disable plane on the removed mesh
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

        // ── Internal ─────────────────────────────────────────────────────────

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

            // Build a SharpDX.Plane from normal + d.
            // SharpDX.Plane(normal, d): equation is dot(normal, p) + d = 0
            // The cross-section clips the side where dot(normal,p) + d > 0.
            // We want to keep the negative side (below/behind the plane), so
            // the normal points in the cut-away direction.
            Vector3 normal = AxisNormal(_axis);

            // dot(normal, p) + d = 0  →  d = -dot(normal, planePoint)
            // planePoint is _offset along the axis
            Vector3 planePoint = normal * _offset;
            float d = -Vector3.Dot(normal, planePoint);

            mesh.Plane1       = new Plane(normal, d);
            mesh.EnablePlane1 = true;
            mesh.CuttingOperation = CuttingOperation.Subtract;
        }

        private static Vector3 AxisNormal(SectionAxis axis)
        {
            switch (axis)
            {
                case SectionAxis.X:  return new Vector3(1, 0, 0);
                case SectionAxis.Y:  return new Vector3(0, 1, 0);
                default:             return new Vector3(0, 0, 1);  // Z
            }
        }
    }

    public enum SectionAxis { X, Y, Z }
}
