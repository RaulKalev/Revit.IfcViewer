using System.Collections.Generic;

namespace ProSchedules.Models
{
    public class ScheduleData
    {
        public Autodesk.Revit.DB.ElementId ScheduleId { get; set; } = Autodesk.Revit.DB.ElementId.InvalidElementId;
        public List<string> Columns { get; set; } = new List<string>();
        public List<Autodesk.Revit.DB.ElementId> ParameterIds { get; set; } = new List<Autodesk.Revit.DB.ElementId>();
        public List<List<string>> Rows { get; set; } = new List<List<string>>();
        public Dictionary<string, bool> IsTypeParameter { get; set; } = new Dictionary<string, bool>();
    }
}
