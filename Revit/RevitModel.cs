using Autodesk.Revit.DB;
using HelixToolkit.Wpf.SharpDX;
using IfcViewer.Viewer;
using SharpDX;
using System.Collections.Generic;

namespace IfcViewer.Revit
{
    /// <summary>
    /// Lightweight result from a single Revit geometry export.
    /// Geometry is merged by colour into a handful of large meshes; per-element
    /// identity is preserved through <see cref="Handles"/> (vertex-range mapping).
    /// The raw per-element buckets are retained so incremental syncs can re-merge
    /// without re-tessellating unchanged elements.
    /// </summary>
    public sealed class RevitModel
    {
        /// <summary>Short label shown in the UI (e.g. "Active View").</summary>
        public string DisplayName { get; }

        /// <summary>Root group containing the merged colour meshes.</summary>
        public GroupModel3D SceneGroup { get; }

        /// <summary>Axis-aligned bounding box of all exported geometry.</summary>
        public BoundingBox Bounds { get; }

        /// <summary>Number of exported elements with geometry.</summary>
        public int MeshCount { get; }

        /// <summary>Total triangle count across all meshes.</summary>
        public int TriangleCount { get; }

        /// <summary>
        /// Maps each exported Revit ElementId to its element handle (bounds + vertex
        /// ranges inside the merged meshes). Used for follow-selection and picking.
        /// </summary>
        public IReadOnlyDictionary<ElementId, ElementHandle> Handles { get; }

        /// <summary>
        /// Maps each exported Revit ElementId to its extracted element properties.
        /// Used by the properties panel when the user clicks an element.
        /// </summary>
        public IReadOnlyDictionary<ElementId, RevitElementInfo> ElementInfos { get; }

        /// <summary>
        /// Raw per-element geometry, kept so <c>RevitExporter.ExportIncremental</c>
        /// can merge dirty elements with unchanged ones without a full re-export.
        /// </summary>
        internal Dictionary<ElementId, ElementBucket> Buckets { get; }

        /// <summary>
        /// Geometry from linked Revit models visible in the exported view, keyed by
        /// (link traversal, linked ElementId) — linked ids can collide with host ids,
        /// so they are kept out of <see cref="Buckets"/>. Linked geometry is exported
        /// on full syncs only and carried over unchanged through incremental patches.
        /// </summary>
        internal Dictionary<LinkedElementKey, ElementBucket> LinkedBuckets { get; }

        /// <summary>Element properties for linked elements, parallel to <see cref="LinkedBuckets"/>.</summary>
        internal Dictionary<LinkedElementKey, RevitElementInfo> LinkedInfos { get; }

        internal RevitModel(string displayName, GroupModel3D sceneGroup,
                            BoundingBox bounds, int meshCount, int triangleCount,
                            IReadOnlyDictionary<ElementId, ElementHandle> handles,
                            IReadOnlyDictionary<ElementId, RevitElementInfo> elementInfos,
                            Dictionary<ElementId, ElementBucket> buckets,
                            Dictionary<LinkedElementKey, ElementBucket> linkedBuckets,
                            Dictionary<LinkedElementKey, RevitElementInfo> linkedInfos)
        {
            DisplayName   = displayName;
            SceneGroup    = sceneGroup;
            Bounds        = bounds;
            MeshCount     = meshCount;
            TriangleCount = triangleCount;
            Handles       = handles       ?? new Dictionary<ElementId, ElementHandle>();
            ElementInfos  = elementInfos  ?? new Dictionary<ElementId, RevitElementInfo>();
            Buckets       = buckets       ?? new Dictionary<ElementId, ElementBucket>();
            LinkedBuckets = linkedBuckets ?? new Dictionary<LinkedElementKey, ElementBucket>();
            LinkedInfos   = linkedInfos   ?? new Dictionary<LinkedElementKey, RevitElementInfo>();
        }

        public override string ToString() => DisplayName;
    }
}
