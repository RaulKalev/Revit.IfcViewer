using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ProSchedules.ExternalEvents
{
    public class ScheduleFieldsHandler : IExternalEventHandler
    {
        public ElementId ScheduleId { get; set; } = ElementId.InvalidElementId;
        public List<ParameterItem> NewFields { get; set; }

        public event Action<int, string> OnUpdateFinished; // success count (dummy), error msg

        public void Execute(UIApplication app)
        {
            if (ScheduleId == ElementId.InvalidElementId || NewFields == null) return;

            Document doc = app.ActiveUIDocument.Document;
            string errorMsg = "";

            using (Transaction t = new Transaction(doc, "Update Schedule Fields"))
            {
                try
                {
                    t.Start();

                    ViewSchedule schedule = doc.GetElement(ScheduleId) as ViewSchedule;
                    if (schedule == null)
                    {
                        errorMsg = "Schedule not found.";
                        OnUpdateFinished?.Invoke(0, errorMsg);
                        return;
                    }

                    ScheduleDefinition def = schedule.Definition;
                    
                    // 1. Get current field IDs
                    var currentFieldIds = def.GetFieldOrder(); // List<ScheduleFieldId>
                    var newFieldIdsOrdered = new List<ScheduleFieldId>();

                    // 2. Process NewFields list to build the target order and add new fields
                    foreach (var item in NewFields)
                    {
                        if (item.IsFieldId)
                        {
                            // Existing field
                            if (item.Id is ScheduleFieldId sfId)
                            {
                                newFieldIdsOrdered.Add(sfId);
                            }
                        }
                        else
                        {
                            // New field to add
                            if (item.Id is SchedulableField schedulableField)
                            {
                                ScheduleField newField = def.AddField(schedulableField);
                                newFieldIdsOrdered.Add(newField.FieldId);
                                
                                // Update item to be an existing field for future reference if needed
                                item.Id = newField.FieldId;
                                item.IsFieldId = true;
                                item.IsScheduled = true;
                            }
                        }
                    }

                    // 3. Remove fields that are no longer in the list
                    // We need to check all previously existing fields. 
                    // If a currentFieldId is NOT present in newFieldIdsOrdered, remove it.
                    // Note: ScheduleFieldId is a value type or object? It's a class. 
                    // But we can compare IDs? No, wait. 
                    // If I re-add a field using AddField, it gets a NEW ScheduleFieldId.
                    // But here I'm reusing ScheduleFieldId for existing ones.
                    // So if I didn't include an old ScheduleFieldId in newFieldIdsOrdered, it means I want to remove it.
                    
                    foreach (var oldId in currentFieldIds)
                    {
                        if (!newFieldIdsOrdered.Contains(oldId))
                        {
                            // Verify it's not hidden? RemoveField removes it from schedule.
                            // If it works like API doc says...
                            if (def.GetField(oldId) != null) // Check if valid
                            {
                                def.RemoveField(oldId);
                            }
                        }
                    }

                    // 4. Set the new order
                    def.SetFieldOrder(newFieldIdsOrdered);

                    t.Commit();
                    OnUpdateFinished?.Invoke(newFieldIdsOrdered.Count, null);
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    errorMsg = ex.ToString();
                    OnUpdateFinished?.Invoke(0, errorMsg);
                }
            }
        }

        public string GetName()
        {
            return "Update Schedule Fields";
        }
    }
}
