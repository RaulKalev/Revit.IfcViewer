using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// One logical element's contiguous vertex/index range inside a merged mesh.
    /// Ranges never move after build — hidden elements are excluded from the
    /// *visible* index buffer, but the master buffers stay intact — so a
    /// hit-test vertex index can always be mapped back to its element.
    /// </summary>
    public sealed class MergedPart
    {
        public MergedMeshInfo Owner;
        public ElementHandle  Handle;
        public int VertexStart;
        public int VertexCount;
        public int IndexStart;   // offset into Owner.MasterIndices
        public int IndexCount;
    }

    /// <summary>
    /// Handle for one logical element (IFC product or Revit element) whose
    /// geometry lives inside one or more merged meshes (one part per colour).
    /// </summary>
    public sealed class ElementHandle
    {
        /// <summary>IfcElementInfo or RevitElementInfo.</summary>
        public object Info;

        public readonly List<MergedPart> Parts = new List<MergedPart>();

        /// <summary>World-space bounds of all parts (for zoom/orbit-pivot/nearest lookups).</summary>
        public BoundingBox Bounds;

        public bool IsHidden;

        internal bool HasBounds;

        /// <summary>
        /// Total triangle surface area, lazily computed for cross-model duplicate
        /// detection (-1 = not computed yet). Area is tessellation-independent, so
        /// it identifies congruent elements even when two exporters mesh the same
        /// object differently.
        /// </summary>
        internal double CachedSurfaceArea = -1;
    }

    /// <summary>
    /// Per merged-mesh metadata. Stored in <see cref="MeshGeometryModel3D.Tag"/> so it
    /// survives the SectionPlaneManager's plain ↔ cross-section node swaps.
    /// </summary>
    public sealed class MergedMeshInfo
    {
        public MeshGeometry3D Geometry;
        public Color4         Colour;
        public bool           IsTransparent;

        /// <summary>Full index buffer including hidden elements.</summary>
        public int[] MasterIndices;

        /// <summary>Parts sorted by VertexStart (append order).</summary>
        public MergedPart[] Parts;

        private int[] _vertexStarts;

        /// <summary>Resolve a hit-test vertex index back to the element that owns it.</summary>
        public ElementHandle HandleFromVertex(int vertexIndex)
        {
            if (Parts == null || Parts.Length == 0) return null;

            if (_vertexStarts == null)
            {
                _vertexStarts = new int[Parts.Length];
                for (int i = 0; i < Parts.Length; i++)
                    _vertexStarts[i] = Parts[i].VertexStart;
            }

            int idx = Array.BinarySearch(_vertexStarts, vertexIndex);
            if (idx < 0) idx = ~idx - 1;
            if (idx < 0 || idx >= Parts.Length) return null;

            var part = Parts[idx];
            if (vertexIndex >= part.VertexStart + part.VertexCount) return null;
            return part.Handle;
        }

        /// <summary>
        /// Rebuild the visible index buffer from the master buffer, skipping parts whose
        /// element is hidden, then refresh the hit-test octree so hidden elements no
        /// longer occlude picking. Call on the UI thread after toggling IsHidden flags.
        /// </summary>
        public void RebuildVisibleIndices()
        {
            int visible = 0;
            bool anyHidden = false;
            foreach (var p in Parts)
            {
                if (p.Handle.IsHidden) anyHidden = true;
                else visible += p.IndexCount;
            }

            if (!anyHidden)
            {
                Geometry.Indices = new IntCollection(MasterIndices);
            }
            else
            {
                var idx = new int[visible];
                int o = 0;
                foreach (var p in Parts)
                {
                    if (p.Handle.IsHidden) continue;
                    Array.Copy(MasterIndices, p.IndexStart, idx, o, p.IndexCount);
                    o += p.IndexCount;
                }
                Geometry.Indices = new IntCollection(idx);
            }

            Geometry.UpdateOctree(true);
        }
    }

    /// <summary>
    /// Accumulates per-element geometry into one merged mesh per (colour, transparency)
    /// bucket. Rendering thousands of elements as a handful of large meshes removes the
    /// per-node CPU cost (frustum tests, state changes, draw calls) that made the
    /// one-mesh-per-element scene CPU-bound at low frame rates.
    /// </summary>
    public sealed class MergedSceneBuilder
    {
        private sealed class Accum
        {
            public Color4 Colour;
            public bool   IsTransparent;
            public readonly List<Vector3>    Positions = new List<Vector3>();
            public readonly List<Vector3>    Normals   = new List<Vector3>();
            public readonly List<int>        Indices   = new List<int>();
            public readonly List<MergedPart> Parts     = new List<MergedPart>();
        }

        private readonly Dictionary<(Color4 colour, bool transparent), Accum> _accums
            = new Dictionary<(Color4, bool), Accum>();

        /// <summary>
        /// Append one element sub-mesh. Safe to call on a background thread.
        /// Also grows <paramref name="handle"/>.Bounds to cover the new vertices.
        /// </summary>
        public void AddElement(ElementHandle handle, Color4 colour, bool isTransparent,
                               IList<Vector3> positions, IList<Vector3> normals, IList<int> indices)
        {
            if (indices == null || indices.Count == 0 || positions == null || positions.Count == 0)
                return;

            var key = (colour, isTransparent);
            if (!_accums.TryGetValue(key, out Accum a))
            {
                a = new Accum { Colour = colour, IsTransparent = isTransparent };
                _accums[key] = a;
            }

            int vBase = a.Positions.Count;
            int iBase = a.Indices.Count;

            var min = new Vector3(float.MaxValue);
            var max = new Vector3(float.MinValue);
            for (int i = 0; i < positions.Count; i++)
            {
                Vector3 p = positions[i];
                a.Positions.Add(p);
                min = Vector3.Min(min, p);
                max = Vector3.Max(max, p);
            }
            for (int i = 0; i < positions.Count; i++)
                a.Normals.Add(i < normals.Count ? normals[i] : new Vector3(0, 1, 0));
            for (int i = 0; i < indices.Count; i++)
                a.Indices.Add(vBase + indices[i]);

            var part = new MergedPart
            {
                Handle      = handle,
                VertexStart = vBase,
                VertexCount = positions.Count,
                IndexStart  = iBase,
                IndexCount  = indices.Count,
            };
            a.Parts.Add(part);
            handle.Parts.Add(part);

            var partBounds = new BoundingBox(min, max);
            handle.Bounds = handle.HasBounds
                ? BoundingBox.Merge(handle.Bounds, partBounds)
                : partBounds;
            handle.HasBounds = true;
        }

        /// <summary>
        /// Materialise the merged geometries and build their hit-test octrees.
        /// Heavy (array copies + octree build) — safe to run on a background thread
        /// because MeshGeometry3D is not a DispatcherObject.
        /// </summary>
        public List<MergedMeshInfo> BuildGeometries()
        {
            var result = new List<MergedMeshInfo>(_accums.Count);
            foreach (var kv in _accums)
            {
                Accum a = kv.Value;
                if (a.Indices.Count == 0) continue;

                var geom = new MeshGeometry3D
                {
                    Positions = new Vector3Collection(a.Positions),
                    Normals   = new Vector3Collection(a.Normals),
                    Indices   = new IntCollection(a.Indices),
                };
                geom.UpdateOctree();

                var info = new MergedMeshInfo
                {
                    Geometry      = geom,
                    Colour        = a.Colour,
                    IsTransparent = a.IsTransparent,
                    MasterIndices = a.Indices.ToArray(),
                    Parts         = a.Parts.ToArray(),
                };
                foreach (var p in info.Parts)
                    p.Owner = info;

                result.Add(info);
            }
            return result;
        }

        /// <summary>
        /// Wrap a merged geometry in a scene node. Must run on the UI thread
        /// (MeshGeometryModel3D is a DispatcherObject).
        /// </summary>
        public static MeshGeometryModel3D CreateMeshNode(MergedMeshInfo info)
        {
            var mat = new PhongMaterial
            {
                DiffuseColor      = info.Colour,
                AmbientColor      = new Color4(0.15f, 0.15f, 0.15f, 1f),
                SpecularColor     = new Color4(0.05f, 0.05f, 0.05f, 1f),
                SpecularShininess = 4f,
                ReflectiveColor   = new Color4(0f, 0f, 0f, 0f),
            };

            return new MeshGeometryModel3D
            {
                Geometry      = info.Geometry,
                Material      = mat,
                IsTransparent = info.IsTransparent,
                Tag           = info,
            };
        }

        /// <summary>
        /// Build a small standalone highlight mesh for one part (used as the selection
        /// overlay). Slices the part's vertices out of the merged buffers with rebased
        /// indices so the overlay uploads only element-sized data.
        /// </summary>
        public static MeshGeometryModel3D CreateHighlightNode(MergedPart part, Color4 emissive)
        {
            MergedMeshInfo owner = part.Owner;

            var pos  = new Vector3Collection(part.VertexCount);
            var norm = new Vector3Collection(part.VertexCount);
            for (int i = 0; i < part.VertexCount; i++)
            {
                pos.Add(owner.Geometry.Positions[part.VertexStart + i]);
                norm.Add(owner.Geometry.Normals[part.VertexStart + i]);
            }

            var idx = new IntCollection(part.IndexCount);
            for (int i = 0; i < part.IndexCount; i++)
                idx.Add(owner.MasterIndices[part.IndexStart + i] - part.VertexStart);

            var geom = new MeshGeometry3D { Positions = pos, Normals = norm, Indices = idx };

            var mat = new PhongMaterial
            {
                DiffuseColor      = owner.Colour,
                AmbientColor      = new Color4(0.15f, 0.15f, 0.15f, 1f),
                SpecularColor     = new Color4(0.05f, 0.05f, 0.05f, 1f),
                SpecularShininess = 4f,
                ReflectiveColor   = new Color4(0f, 0f, 0f, 0f),
                EmissiveColor     = emissive,
            };

            return new MeshGeometryModel3D
            {
                Geometry         = geom,
                Material         = mat,
                IsTransparent    = owner.IsTransparent,
                IsHitTestVisible = false,
                // Negative bias pulls the overlay slightly toward the camera so it
                // wins the depth test against the identical triangles underneath.
                DepthBias        = -100,
            };
        }
    }
}
