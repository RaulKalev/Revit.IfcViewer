using System;
using System.Collections.Generic;

namespace IfcCore
{
    /// <summary>
    /// Result of an IFC model load operation.
    /// Contains portable geometry data and element metadata that
    /// the rendering layer can use to build scene objects.
    /// </summary>
    public class IfcLoadResult
    {
        /// <summary>Whether the load completed successfully.</summary>
        public bool Success { get; set; }

        /// <summary>Full path to the loaded IFC file.</summary>
        public string FilePath { get; set; }

        /// <summary>IFC schema version (e.g. "IFC2X3", "IFC4").</summary>
        public string SchemaVersion { get; set; }

        /// <summary>Storage provider used (e.g. "MemoryModel", "Esent").</summary>
        public string StorageProvider { get; set; }

        /// <summary>Total number of IFC entities in the model.</summary>
        public int EntityCount { get; set; }

        /// <summary>Number of geometric products extracted.</summary>
        public int ProductCount { get; set; }

        /// <summary>Total wall-clock duration of the load operation.</summary>
        public TimeSpan Duration { get; set; }

        /// <summary>
        /// Per-element geometry data keyed by product label.
        /// Each entry contains the renderable mesh data for one IFC product.
        /// </summary>
        public Dictionary<int, IfcGeometryData> Elements { get; set; }
            = new Dictionary<int, IfcGeometryData>();

        /// <summary>
        /// Per-element metadata keyed by product label.
        /// </summary>
        public Dictionary<int, IfcElementData> ElementInfo { get; set; }
            = new Dictionary<int, IfcElementData>();

        /// <summary>
        /// Non-fatal warnings encountered during loading.
        /// </summary>
        public List<string> Warnings { get; set; } = new List<string>();

        /// <summary>Error message if <see cref="Success"/> is false.</summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// Portable geometry data for a single IFC product element.
    /// Coordinate system: Y-up (Helix convention), scaled to metres.
    /// </summary>
    public class IfcGeometryData
    {
        /// <summary>Vertex positions as flat array [x0,y0,z0, x1,y1,z1, ...].</summary>
        public float[] Positions { get; set; }

        /// <summary>Vertex normals as flat array [nx0,ny0,nz0, ...].</summary>
        public float[] Normals { get; set; }

        /// <summary>Triangle indices.</summary>
        public int[] Indices { get; set; }

        /// <summary>Diffuse colour RGBA (0–1 range).</summary>
        public float ColourR { get; set; }
        public float ColourG { get; set; }
        public float ColourB { get; set; }
        public float ColourA { get; set; } = 1f;

        /// <summary>Whether this element should render as transparent.</summary>
        public bool IsTransparent { get; set; }
    }
}
