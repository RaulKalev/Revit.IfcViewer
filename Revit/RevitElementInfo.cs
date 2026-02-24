using System.Collections.Generic;

namespace IfcViewer.Revit
{
    /// <summary>
    /// Revit element properties extracted during export.
    /// Stored alongside each mesh so the properties panel can display them
    /// when the user clicks a Revit element in the viewport.
    /// </summary>
    public class RevitElementInfo
    {
        public string Name       { get; set; } = "(unnamed)";
        public string Category   { get; set; } = "";
        public string FamilyName { get; set; } = "";
        public string TypeName   { get; set; } = "";
        public string ElementId  { get; set; } = "";

        /// <summary>
        /// Parameters grouped by their Revit parameter group label
        /// (e.g. "Identity Data", "Constraints", "Dimensions").
        /// Inner dictionary maps parameter name → formatted value string.
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> PropertySets { get; }
            = new Dictionary<string, Dictionary<string, string>>();
    }
}
