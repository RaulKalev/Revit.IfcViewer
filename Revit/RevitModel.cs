using HelixToolkit.Wpf.SharpDX;
using SharpDX;

namespace IfcViewer.Revit
{
    /// <summary>
    /// Lightweight result from a single Revit geometry export.
    /// Holds the Helix GroupModel3D and metadata for display.
    /// </summary>
    public sealed class RevitModel
    {
        /// <summary>Short label shown in the UI (e.g. "Active View").</summary>
        public string DisplayName { get; }

        /// <summary>Root group containing all Helix meshes for the exported geometry.</summary>
        public GroupModel3D SceneGroup { get; }

        /// <summary>Axis-aligned bounding box of all exported geometry.</summary>
        public BoundingBox Bounds { get; }

        /// <summary>Number of mesh buckets (one per colour) in the scene.</summary>
        public int MeshCount { get; }

        /// <summary>Total triangle count across all meshes.</summary>
        public int TriangleCount { get; }

        public RevitModel(string displayName, GroupModel3D sceneGroup,
                          BoundingBox bounds, int meshCount, int triangleCount)
        {
            DisplayName = displayName;
            SceneGroup  = sceneGroup;
            Bounds      = bounds;
            MeshCount   = meshCount;
            TriangleCount = triangleCount;
        }

        public override string ToString() => DisplayName;
    }
}
