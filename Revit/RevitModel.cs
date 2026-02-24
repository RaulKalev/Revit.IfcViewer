using Autodesk.Revit.DB;
using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System.Collections.Generic;

namespace IfcViewer.Revit
{
    /// <summary>
    /// Lightweight result from a single Revit geometry export.
    /// Holds the Helix GroupModel3D, per-element mesh map, and metadata for display.
    /// </summary>
    public sealed class RevitModel
    {
        /// <summary>Short label shown in the UI (e.g. "Active View").</summary>
        public string DisplayName { get; }

        /// <summary>Root group containing all Helix meshes for the exported geometry.</summary>
        public GroupModel3D SceneGroup { get; }

        /// <summary>Axis-aligned bounding box of all exported geometry.</summary>
        public BoundingBox Bounds { get; }

        /// <summary>Number of element meshes in the scene.</summary>
        public int MeshCount { get; }

        /// <summary>Total triangle count across all meshes.</summary>
        public int TriangleCount { get; }

        /// <summary>
        /// Maps each exported Revit ElementId to its Helix mesh.
        /// Used for incremental updates — individual element meshes can be replaced
        /// in-place without a full re-export.
        /// </summary>
        public IReadOnlyDictionary<ElementId, MeshGeometryModel3D> ElementMeshes { get; }

        /// <summary>
        /// Maps each exported Revit ElementId to its extracted element properties.
        /// Used by the properties panel when the user clicks an element.
        /// </summary>
        public IReadOnlyDictionary<ElementId, RevitElementInfo> ElementInfos { get; }

        public RevitModel(string displayName, GroupModel3D sceneGroup,
                          BoundingBox bounds, int meshCount, int triangleCount,
                          IReadOnlyDictionary<ElementId, MeshGeometryModel3D> elementMeshes = null,
                          IReadOnlyDictionary<ElementId, RevitElementInfo> elementInfos = null)
        {
            DisplayName   = displayName;
            SceneGroup    = sceneGroup;
            Bounds        = bounds;
            MeshCount     = meshCount;
            TriangleCount = triangleCount;
            ElementMeshes = elementMeshes ?? new Dictionary<ElementId, MeshGeometryModel3D>();
            ElementInfos  = elementInfos  ?? new Dictionary<ElementId, RevitElementInfo>();
        }

        public override string ToString() => DisplayName;
    }
}
