using System.Collections.Generic;

namespace IfcViewer.Ifc
{
    /// <summary>
    /// Properties extracted from an IFC product during loading.
    /// Stored in <see cref="IfcModel.ElementMap"/> alongside each mesh
    /// so the properties panel can display them when the user clicks an element.
    /// </summary>
    public class IfcElementInfo
    {
        /// <summary>IFC Name attribute, e.g. "Basic Wall:Generic - 200mm:123456".</summary>
        public string Name { get; set; } = "(unnamed)";

        /// <summary>IFC express type name, e.g. "IfcWallStandardCase".</summary>
        public string Type { get; set; } = "Unknown";

        /// <summary>IFC GlobalId (GUID), uniquely identifies this element in the file.</summary>
        public string GlobalId { get; set; } = "";

        /// <summary>
        /// Property sets keyed by pset name.
        /// Each inner dictionary maps property name → formatted string value.
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> PropertySets { get; }
            = new Dictionary<string, Dictionary<string, string>>();
    }
}
