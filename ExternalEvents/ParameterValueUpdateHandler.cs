using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

namespace ProSchedules.ExternalEvents
{
    public class ParameterBatchData
    {
        public string ElementIdStr { get; set; }
        public ElementId ParameterId { get; set; }
        public string Value { get; set; }
    }

    public class ParameterValueUpdateHandler : IExternalEventHandler
    {
        // Single mode properties
        public string ElementIdStr { get; set; }
        public string ParameterIdStr { get; set; }
        public string NewValue { get; set; }
        
        // Batch mode properties
        public bool IsBatchMode { get; set; }
        public List<ParameterBatchData> BatchData { get; set; } = new List<ParameterBatchData>();

        // Result properties
        public bool Success { get; private set; }
        public string ErrorMessage { get; private set; }

        // Event to notify UI when done
        public event Action<int, int, string> OnUpdateFinished;

        public void Execute(UIApplication app)
        {
            Success = false;
            ErrorMessage = "";
            int successCount = 0;
            int failCount = 0;

            try
            {
                Document doc = app.ActiveUIDocument.Document;

                using (Transaction trans = new Transaction(doc, "Update Parameter Values"))
                {
                    trans.Start();

                    if (IsBatchMode)
                    {
                        if (BatchData == null || BatchData.Count == 0)
                        {
                            ErrorMessage = "No batch data provided.";
                            OnUpdateFinished?.Invoke(0, 0, ErrorMessage);
                            return;
                        }

                        foreach (var item in BatchData)
                        {
                            if (UpdateSingleParameter(doc, item.ElementIdStr, item.ParameterId, item.Value))
                            {
                                successCount++;
                            }
                            else
                            {
                                failCount++;
                            }
                        }
                    }
                    else
                    {
                        // Single Mode
                        if (string.IsNullOrEmpty(ElementIdStr))
                        {
                            ErrorMessage = "Invalid element ID";
                            OnUpdateFinished?.Invoke(0, 1, ErrorMessage);
                            return;
                        }

                        if (!long.TryParse(ParameterIdStr, out long paramIdValue))
                        {
                            ErrorMessage = "Invalid parameter ID";
                            OnUpdateFinished?.Invoke(0, 1, ErrorMessage);
                            return;
                        }

                        // Support comma-separated IDs even in single mode (for grouped cells)
                        string[] idStrings = ElementIdStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        ElementId paramId = new ElementId(paramIdValue);
                        
                        foreach(var id in idStrings)
                        {
                            if (UpdateSingleParameter(doc, id, paramId, NewValue))
                            {
                                successCount++;
                            }
                            else
                            {
                                failCount++;
                            }
                        }
                    }

                    if (successCount > 0)
                    {
                        trans.Commit();
                        Success = true;
                    }
                    else
                    {
                        trans.RollBack();
                        if (string.IsNullOrEmpty(ErrorMessage))
                        {
                            ErrorMessage = "Failed to update any parameters.";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                failCount++; // Count the whole batch as fail if exception? Or just report error.
            }
            finally
            {
                // Always notify
                OnUpdateFinished?.Invoke(successCount, failCount, ErrorMessage);
            }
        }

        private bool UpdateSingleParameter(Document doc, string elementIdStr, ElementId parameterId, string value)
        {
            if (!long.TryParse(elementIdStr, out long elemIdValue)) return false;

            ElementId elementId = new ElementId(elemIdValue);
            Element element = doc.GetElement(elementId);
            
            if (element != null)
            {
                return SetParameterValue(doc, element, parameterId, value);
            }
            return false;
        }

        private bool SetParameterValue(Document doc, Element element, ElementId parameterId, string value)
        {
            Parameter p = null;
            long idValue = parameterId.Value;

            // Try to get parameter from instance
            if (idValue < 0)
            {
                p = element.get_Parameter((BuiltInParameter)(int)idValue);
            }
            else
            {
                try
                {
                    var paramElem = doc.GetElement(parameterId);
                    if (paramElem != null)
                    {
                        p = element.LookupParameter(paramElem.Name);
                    }
                }
                catch { }
            }

            // Fallback: iterate through instance parameters
            if (p == null)
            {
                foreach (Parameter param in element.Parameters)
                {
                    if (param.Id.Value == parameterId.Value)
                    {
                        p = param;
                        break;
                    }
                }
            }

            // If not found on instance, try type
            if (p == null)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    Element typeElem = doc.GetElement(typeId);
                    if (typeElem != null)
                    {
                        if (idValue < 0)
                        {
                            p = typeElem.get_Parameter((BuiltInParameter)(int)idValue);
                        }
                        else
                        {
                            var paramElem = doc.GetElement(parameterId);
                            if (paramElem != null)
                            {
                                p = typeElem.LookupParameter(paramElem.Name);
                            }
                        }

                        if (p == null)
                        {
                            foreach (Parameter param in typeElem.Parameters)
                            {
                                if (param.Id.Value == parameterId.Value)
                                {
                                    p = param;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            if (p == null || p.IsReadOnly)
                return false;

            try
            {
                switch (p.StorageType)
                {
                    case StorageType.String:
                        p.Set(value);
                        return true;

                    case StorageType.Integer:
                        if (int.TryParse(value, out int intVal))
                        {
                            p.Set(intVal);
                            return true;
                        }
                        break;

                    case StorageType.Double:
                        // First try SetValueString to handle Project Units (e.g. "1000" mm -> "3.28" ft)
                        if (p.SetValueString(value))
                        {
                            return true;
                        }

                        // Fallback to internal units conversion if string parsing fails/isn't applicable
                        if (double.TryParse(value, out double dblVal))
                        {
                            p.Set(dblVal);
                            return true;
                        }
                        break;

                    case StorageType.ElementId:
                        // For now, skip ElementId type parameters
                        return false;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        public string GetName()
        {
            return "ParameterValueUpdateHandler";
        }
    }
}
