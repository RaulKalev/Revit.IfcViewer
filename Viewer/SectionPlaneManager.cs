using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System.Collections.Generic;

namespace IfcViewer.Viewer
{
    /// <summary>One section plane: a world-space normal pointing INTO the removed
    /// half-space, and a point on the plane.</summary>
    public struct SectionPlaneDef
    {
        public readonly Vector3 Normal;
        public readonly Vector3 Point;

        public SectionPlaneDef(Vector3 normal, Vector3 point)
        {
            Normal = normal;
            Point  = point;
        }
    }

    /// <summary>
    /// Manages up to four arbitrary section planes applied to every mesh in the scene.
    /// Call <see cref="SetPlanes"/> whenever the plane set changes.
    ///
    /// Meshes are loaded as plain <see cref="MeshGeometryModel3D"/> (cheap Blinn-Phong
    /// shader). When sectioning is enabled, this manager swaps each registered mesh
    /// to a <see cref="CrossSectionMeshGeometryModel3D"/> in-place inside its parent group,
    /// then applies the clip planes. When disabled it swaps them back to plain meshes.
    ///
    /// The interactive workflow and plane visuals live in
    /// <see cref="SectionPlaneController"/>; this class owns only the clipping.
    /// </summary>
    public sealed class SectionPlaneManager
    {
        /// <summary>Maximum number of simultaneous section planes.</summary>
        public const int MaxPlanes = 4;

        // ── Registered mesh entries ───────────────────────────────────────────
        private sealed class MeshEntry
        {
            public GroupModel3D        Parent;
            public MeshGeometryModel3D LiveMesh;
            public bool                IsCrossSection;
        }

        private readonly List<MeshEntry> _entries = new List<MeshEntry>();

        // ── Backing fields ────────────────────────────────────────────────────
        private bool _enabled;
        private readonly List<SectionPlaneDef> _planes = new List<SectionPlaneDef>(MaxPlanes);

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
            }
        }

        /// <summary>
        /// Replace the active plane set (first <see cref="MaxPlanes"/> are used).
        /// Each plane clips everything on the side its normal points towards; with
        /// several planes the kept region is the intersection of the kept sides.
        /// </summary>
        public void SetPlanes(IReadOnlyList<SectionPlaneDef> planes)
        {
            _planes.Clear();
            if (planes != null)
            {
                for (int i = 0; i < planes.Count && i < MaxPlanes; i++)
                    _planes.Add(planes[i]);
            }
            if (_enabled) ApplyAll();
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
                // Neutral cap colour drawn where the cut slices through closed
                // solids (stencil fill; Helix's default is bright green).
                CrossSectionColor = System.Windows.Media.Color.FromRgb(0xB4, 0xB4, 0xB4),
                // Tag carries the MergedMeshInfo used to resolve hit-tests back to
                // elements — it must survive the node swap.
                Tag           = plain.Tag,
                Visibility    = plain.Visibility,
                // Per-model depth bias (anti z-fighting between overlapping models)
                // must survive the node swap too.
                DepthBias             = plain.DepthBias,
                SlopeScaledDepthBias  = plain.SlopeScaledDepthBias,
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
                Tag           = cs.Tag,
                Visibility    = cs.Visibility,
                DepthBias             = cs.DepthBias,
                SlopeScaledDepthBias  = cs.SlopeScaledDepthBias,
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
            mesh.CuttingOperation = CuttingOperation.Intersect;

            mesh.EnablePlane1 = _planes.Count > 0;
            if (_planes.Count > 0) mesh.Plane1 = ToShaderPlane(_planes[0]);
            mesh.EnablePlane2 = _planes.Count > 1;
            if (_planes.Count > 1) mesh.Plane2 = ToShaderPlane(_planes[1]);
            mesh.EnablePlane3 = _planes.Count > 2;
            if (_planes.Count > 2) mesh.Plane3 = ToShaderPlane(_planes[2]);
            mesh.EnablePlane4 = _planes.Count > 3;
            if (_planes.Count > 3) mesh.Plane4 = ToShaderPlane(_planes[3]);
        }

        /// <summary>
        /// Helix's clip-plane shader does NOT use SharpDX's Hessian convention
        /// (dot(n,x) + d = 0). It reconstructs the plane point as normal * D
        /// (vsMeshDefault.hlsl: clipDistance = dot(n, wp - n * D)), i.e. it
        /// expects D = +dot(n, pointOnPlane). Passing the Hessian d mirrors
        /// the plane across the world origin — cuts then only look right for
        /// planes that happen to mirror back into the model.
        ///
        /// Hardware clipping discards pixels with negative clip distance, so
        /// the KEPT half-space is the side the shader normal points to. The
        /// def normal points into the REMOVED half-space, so hand the shader
        /// the opposite normal, with D measured along it.
        /// </summary>
        private static Plane ToShaderPlane(SectionPlaneDef def)
        {
            var n = -def.Normal;
            return new Plane(n, Vector3.Dot(n, def.Point));
        }
    }
}
