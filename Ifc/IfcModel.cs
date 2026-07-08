using HelixToolkit.Wpf.SharpDX;
using IfcViewer.Viewer;
using SharpDX;
using System.Collections.Generic;
using System.IO;

namespace IfcViewer.Ifc
{
    /// <summary>
    /// Lightweight record representing one loaded IFC file.
    /// Holds the Helix GroupModel3D (geometry) and metadata for display.
    /// Geometry is merged by colour into a handful of large meshes; per-element
    /// identity is preserved through <see cref="Handles"/> (vertex-range mapping).
    /// </summary>
    public sealed class IfcModel
    {
        /// <summary>Full path to the source .ifc file.</summary>
        public string FilePath { get; }

        /// <summary>File name without extension, shown in the model list.</summary>
        public string DisplayName => Path.GetFileNameWithoutExtension(FilePath);

        /// <summary>Root group containing the merged colour meshes for this file.</summary>
        public GroupModel3D SceneGroup { get; }

        /// <summary>Axis-aligned bounding box of all loaded geometry.</summary>
        public BoundingBox Bounds { get; }

        /// <summary>Number of IFC products that produced geometry.</summary>
        public int MeshCount { get; }

        /// <summary>Total triangle count across all meshes in this model.</summary>
        public int TriangleCount { get; }

        /// <summary>
        /// One handle per IFC product with geometry. Each handle carries the element's
        /// IfcElementInfo, its bounds, and its vertex ranges inside the merged meshes —
        /// used for click-selection, hide/unhide, and Revit follow-selection.
        /// </summary>
        public IReadOnlyList<ElementHandle> Handles { get; }

        public IfcModel(string filePath, GroupModel3D sceneGroup,
                        BoundingBox bounds, int meshCount, int triangleCount,
                        IReadOnlyList<ElementHandle> handles)
        {
            FilePath      = filePath;
            SceneGroup    = sceneGroup;
            Bounds        = bounds;
            MeshCount     = meshCount;
            TriangleCount = triangleCount;
            Handles       = handles ?? new List<ElementHandle>();
        }

        public override string ToString() => DisplayName;
    }
}
