using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System.Collections.Generic;
using System.IO;

namespace IfcViewer.Ifc
{
    /// <summary>
    /// Lightweight record representing one loaded IFC file.
    /// Holds the Helix GroupModel3D (geometry) and metadata for display.
    /// </summary>
    public sealed class IfcModel
    {
        /// <summary>Full path to the source .ifc file.</summary>
        public string FilePath { get; }

        /// <summary>File name without extension, shown in the model list.</summary>
        public string DisplayName => Path.GetFileNameWithoutExtension(FilePath);

        /// <summary>Root group containing all Helix meshes for this file.</summary>
        public GroupModel3D SceneGroup { get; }

        /// <summary>Axis-aligned bounding box of all loaded geometry.</summary>
        public BoundingBox Bounds { get; }

        /// <summary>Number of IFC products that produced geometry.</summary>
        public int MeshCount { get; }

        /// <summary>Total triangle count across all meshes in this model.</summary>
        public int TriangleCount { get; }

        /// <summary>
        /// Maps each per-element mesh to its extracted IFC properties.
        /// Used by the properties panel when the user clicks an element.
        /// </summary>
        public IReadOnlyDictionary<MeshGeometryModel3D, IfcElementInfo> ElementMap { get; }

        public IfcModel(string filePath, GroupModel3D sceneGroup,
                        BoundingBox bounds, int meshCount, int triangleCount,
                        IReadOnlyDictionary<MeshGeometryModel3D, IfcElementInfo> elementMap)
        {
            FilePath      = filePath;
            SceneGroup    = sceneGroup;
            Bounds        = bounds;
            MeshCount     = meshCount;
            TriangleCount = triangleCount;
            ElementMap    = elementMap
                ?? new Dictionary<MeshGeometryModel3D, IfcElementInfo>();
        }

        public override string ToString() => DisplayName;
    }
}
