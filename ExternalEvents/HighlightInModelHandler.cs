using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ProSchedules.ExternalEvents
{
    public class HighlightInModelHandler : IExternalEventHandler
    {
        public List<ElementId> ElementIds { get; set; } = new List<ElementId>();
        public bool Success { get; private set; }
        public string ErrorMessage { get; private set; }

        // State to track last view type for toggling
        private enum ToggleViewType { None, ThreeD, FloorPlan }
        private ToggleViewType _lastViewType = ToggleViewType.None;

        public void Execute(UIApplication app)
        {
            Success = false;
            ErrorMessage = string.Empty;

            try
            {
                var doc = app.ActiveUIDocument.Document;
                var uidoc = app.ActiveUIDocument;

                if (ElementIds == null || ElementIds.Count == 0)
                {
                    ErrorMessage = "No elements to highlight.";
                    return;
                }

                // Filter out invalid element IDs
                var validIds = ElementIds.Where(id => id != null && id != ElementId.InvalidElementId && doc.GetElement(id) != null).ToList();

                if (validIds.Count == 0)
                {
                    ErrorMessage = "No valid elements found in the project.";
                    return;
                }

                // Determine target view type
                View targetView = null;
                var currentView = doc.ActiveView;

                // Logic:
                // 1. If current view shows elements and matches desired toggle, keep it?
                // 2. User wants: Floor Plan First -> then 3D -> then Plan...
                
                // If last was None (fresh start), try Plan first.
                // If last was Plan, try 3D.
                // If last was 3D, try Plan.
                
                // However, we should also respect what the user is currently looking at if they manually switched.
                // But specifically for the button click:
                
                bool preferPlan = (_lastViewType == ToggleViewType.None || _lastViewType == ToggleViewType.ThreeD);
                
                if (preferPlan)
                {
                    targetView = FindSuitablePlanView(doc, validIds);
                    if (targetView != null)
                    {
                        _lastViewType = ToggleViewType.FloorPlan;
                    }
                    else
                    {
                        // Fallback to 3D if no plan found
                        targetView = Get3DView(doc);
                        _lastViewType = ToggleViewType.ThreeD;
                    }
                }
                else
                {
                    // Prefer 3D
                    targetView = Get3DView(doc);
                    _lastViewType = ToggleViewType.ThreeD;
                }

                if (targetView == null)
                {
                    ErrorMessage = "Could not find a suitable view (Plan or 3D).";
                    return;
                }

                // Switch View is required
                if (uidoc.ActiveView.Id != targetView.Id)
                {
                    uidoc.ActiveView = targetView;
                }

                // Removed Isolation Logic per user request (just highlight and focus)
                /*
                using (Transaction trans = new Transaction(doc, "Highlight Elements"))
                {
                    trans.Start();
                    try
                    {
                        targetView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                        targetView.IsolateElementsTemporary(validIds);
                        trans.Commit();
                    }
                    catch
                    {
                        trans.RollBack();
                        throw;
                    }
                }
                */

                // Set selection to highlight the elements
                uidoc.Selection.SetElementIds(validIds);

                // Zoom to fit the elements
                uidoc.ShowElements(validIds);

                Success = true;
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                Success = false;
            }
        }

        private View Get3DView(Document doc)
        {
            // Try to find an existing 3D view (preferably {3D})
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate);

            // Prefer the default {3D} view
            var defaultView = collector.FirstOrDefault(v => v.Name == "{3D}");
            if (defaultView != null)
                return defaultView;

            // Otherwise return the first available 3D view
            return collector.FirstOrDefault();
        }

        private View FindSuitablePlanView(Document doc, List<ElementId> elementIds)
        {
            // Implementation Strategy:
            // 1. Get the level of the first element (most common use case).
            // 2. Find a floor plan associated with that level.
            
            Element firstElem = doc.GetElement(elementIds.First());
            if (firstElem == null) return null;

            Level level = null;
            
            // Try to get level property
            if (firstElem.LevelId != ElementId.InvalidElementId)
            {
                level = doc.GetElement(firstElem.LevelId) as Level;
            }
            
            // If no direct level, try parameter
            if (level == null)
            {
                Parameter levelParam = firstElem.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) ?? 
                                       firstElem.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM) ??
                                       firstElem.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
                                       
                if (levelParam != null && levelParam.HasValue)
                {
                     level = doc.GetElement(levelParam.AsElementId()) as Level;
                }
            }

            if (level == null) return null; // Can't determine level

            // Find plan view for this level
            var planView = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(v => !v.IsTemplate && v.ViewType == ViewType.FloorPlan && v.GenLevel != null && v.GenLevel.Id == level.Id)
                .FirstOrDefault();

            return planView;
        }

        public string GetName()
        {
            return "HighlightInModelHandler";
        }
    }
}
