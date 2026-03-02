using System.Collections.Generic;

namespace IfcCore
{
    /// <summary>
    /// Portable element metadata extracted from an IFC product.
    /// Framework-agnostic — no HelixToolkit or SharpDX dependencies.
    /// </summary>
    public class IfcElementData
    {
        /// <summary>IFC Name attribute, e.g. "Basic Wall:Generic - 200mm:123456".</summary>
        public string Name { get; set; } = "(unnamed)";

        /// <summary>IFC express type name, e.g. "IfcWallStandardCase".</summary>
        public string Type { get; set; } = "Unknown";

        /// <summary>IFC GlobalId (GUID), uniquely identifies this element in the file.</summary>
        public string GlobalId { get; set; } = "";

        /// <summary>
        /// Internal product label used as key in geometry dictionaries.
        /// </summary>
        public int ProductLabel { get; set; }

        /// <summary>
        /// Property sets keyed by pset name.
        /// Each inner dictionary maps property name → formatted string value.
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> PropertySets { get; set; }
            = new Dictionary<string, Dictionary<string, string>>();
    }
}
