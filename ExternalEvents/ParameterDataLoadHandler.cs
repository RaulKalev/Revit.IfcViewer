using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;
using ProSchedules.Services;
using System;
using System.Collections.Generic;

namespace ProSchedules.ExternalEvents
{
    public class ParameterDataLoadHandler : IExternalEventHandler
    {
        public ElementId ScheduleId { get; set; } = ElementId.InvalidElementId;
        
        public event Action<List<ParameterItem>, List<ParameterItem>, string> OnDataLoaded; // available, scheduled, categoryName

        public void Execute(UIApplication app)
        {
            if (ScheduleId == ElementId.InvalidElementId) return;

            Document doc = app.ActiveUIDocument.Document;
            var service = new RevitService(doc);

            try
            {
                ViewSchedule schedule = doc.GetElement(ScheduleId) as ViewSchedule;
                if (schedule != null)
                {
                    var available = service.GetAvailableParameters(schedule);
                    var scheduled = service.GetScheduledParameters(schedule);
                    
                    string categoryName = "Multi-Category";
                    if (schedule.Definition.CategoryId != ElementId.InvalidElementId)
                    {
                        var cat = Category.GetCategory(doc, schedule.Definition.CategoryId);
                        if (cat != null) categoryName = cat.Name;
                    }
                    
                    OnDataLoaded?.Invoke(available, scheduled, categoryName);
                }
            }
            catch (Exception)
            {
                // Handle error or just invoke null
                OnDataLoaded?.Invoke(new List<ParameterItem>(), new List<ParameterItem>(), "Error");
            }
        }

        public string GetName()
        {
            return "Load Parameter Data";
        }
    }
}
