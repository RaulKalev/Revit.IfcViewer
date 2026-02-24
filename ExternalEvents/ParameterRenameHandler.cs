using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;

namespace ProSchedules.ExternalEvents
{
    /// <summary>
    /// ExternalEventHandler for applying parameter rename operations to Revit elements.
    /// </summary>
    public class ParameterRenameHandler : IExternalEventHandler
    {
        /// <summary>
        /// List of rename items to process.
        /// </summary>
        public List<ScheduleRenameItem> RenameItems { get; set; } = new List<ScheduleRenameItem>();

        /// <summary>
        /// Event fired when the rename operation completes.
        /// Parameters: successCount, failCount, errorMessage
        /// </summary>
        public event Action<int, int, string> OnRenameFinished;

        public void Execute(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            if (uidoc == null) return;
            var doc = uidoc.Document;

            int successCount = 0;
            int failCount = 0;
            string errorMsg = "";

            using (Transaction t = new Transaction(doc, "Batch Rename Parameters"))
            {
                t.Start();

                try
                {
                    // Group by ElementId to avoid redundant lookups
                    var groupedByElement = new Dictionary<long, List<ScheduleRenameItem>>();
                    foreach (var item in RenameItems)
                    {
                        if (item.Original == item.New) continue; // Skip unchanged

                        long idVal = item.ElementId.Value;
                        if (!groupedByElement.ContainsKey(idVal))
                        {
                            groupedByElement[idVal] = new List<ScheduleRenameItem>();
                        }
                        groupedByElement[idVal].Add(item);
                    }

                    foreach (var kvp in groupedByElement)
                    {
                        ElementId elemId = new ElementId(kvp.Key);
                        Element element = doc.GetElement(elemId);
                        if (element == null)
                        {
                            failCount += kvp.Value.Count;
                            errorMsg = $"Element with ID {kvp.Key} not found";
                            continue;
                        }

                        foreach (var renameItem in kvp.Value)
                        {
                            try
                            {
                                bool success = SetParameterValue(doc, element, renameItem);
                                if (success)
                                {
                                    successCount++;
                                }
                                else
                                {
                                    failCount++;
                                    errorMsg = $"Failed to set parameter '{renameItem.ParameterName}' on element {kvp.Key}";
                                }
                            }
                            catch (Exception ex)
                            {
                                failCount++;
                                errorMsg = ex.Message;
                            }
                        }
                    }

                    t.Commit();
                }
                catch (Exception ex)
                {
                    if (t.GetStatus() == TransactionStatus.Started)
                    {
                        t.RollBack();
                    }
                    failCount++;
                    errorMsg = "Critical Error: " + ex.Message;
                }
            }

            OnRenameFinished?.Invoke(successCount, failCount, errorMsg);
        }

        private bool SetParameterValue(Document doc, Element element, ScheduleRenameItem renameItem)
        {
            Parameter param = null;

            // Try to find the parameter on the instance first
            param = element.LookupParameter(renameItem.ParameterName);

            // If it's a type parameter or not found on instance, try the type
            if ((param == null || renameItem.IsTypeParameter) && element.GetTypeId() != ElementId.InvalidElementId)
            {
                Element typeElem = doc.GetElement(element.GetTypeId());
                if (typeElem != null)
                {
                    var typeParam = typeElem.LookupParameter(renameItem.ParameterName);
                    if (typeParam != null)
                    {
                        param = typeParam;
                    }
                }
            }

            if (param == null)
            {
                return false;
            }

            if (param.IsReadOnly)
            {
                return false;
            }

            // Set the value based on storage type
            switch (param.StorageType)
            {
                case StorageType.String:
                    param.Set(renameItem.New);
                    return true;

                case StorageType.Integer:
                    if (int.TryParse(renameItem.New, out int intVal))
                    {
                        param.Set(intVal);
                        return true;
                    }
                    return false;

                case StorageType.Double:
                    if (double.TryParse(renameItem.New, out double dblVal))
                    {
                        param.Set(dblVal);
                        return true;
                    }
                    return false;

                default:
                    return false;
            }
        }

        public string GetName()
        {
            return "Parameter Rename Handler";
        }
    }
}
