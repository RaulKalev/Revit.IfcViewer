using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ProSchedules.Services
{
    public class RevitService
    {
        private Document _doc;

        public RevitService(Document doc)
        {
            _doc = doc;
        }

        public List<SheetItem> GetSheets()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsTemplate)
                .Select(s => new SheetItem(s))
                .OrderBy(s => s.SheetNumber)
                .ToList();
        }

        public List<ViewSchedule> GetSchedules()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(s => !s.IsTemplate && !s.IsInternalKeynoteSchedule && !s.IsTitleblockRevisionSchedule)
                .OrderBy(s => s.Name)
                .ToList();
        }

        public ScheduleData GetScheduleData(ViewSchedule schedule)
        {
            var data = new Models.ScheduleData();
            data.ScheduleId = schedule.Id;
            if (schedule == null) return data;

            ScheduleDefinition def = schedule.Definition;
            ElementId categoryId = def.CategoryId;

            var fields = new List<ScheduleField>();
            var fieldIds = def.GetFieldOrder();

            foreach (var id in fieldIds)
            {
                ScheduleField field = def.GetField(id);
                if (!field.IsHidden)
                {
                    fields.Add(field);
                    data.Columns.Add(field.GetName());
                    data.ParameterIds.Add(field.ParameterId);
                }
            }

            IList<Element> elements = new List<Element>();
            if (categoryId != ElementId.InvalidElementId)
            {
                elements = new FilteredElementCollector(_doc)
                    .OfCategoryId(categoryId)
                    .WhereElementIsNotElementType()
                    .ToElements();
            }

            if (!data.Columns.Contains("ElementId"))
            {
                data.Columns.Insert(0, "ElementId");
                data.ParameterIds.Insert(0, ElementId.InvalidElementId); // Placeholder
            }
            if (!data.Columns.Contains("TypeName"))
            {
                data.Columns.Insert(1, "TypeName");
                data.ParameterIds.Insert(1, ElementId.InvalidElementId); // Placeholder
            }

            foreach (Element el in elements)
            {
                var rowData = new List<string>();
                rowData.Add(el.Id.Value.ToString());
                
                ElementId typeId = el.GetTypeId();
                string typeName = "";
                if (typeId != ElementId.InvalidElementId)
                {
                    Element typeElem = _doc.GetElement(typeId);
                    if (typeElem != null) typeName = typeElem.Name;
                }
                rowData.Add(typeName);

                foreach (ScheduleField field in fields)
                {
                    string val = "";
                    bool isType = false;
                    
                    if (field.ParameterId == ElementId.InvalidElementId)
                    {
                        val = "";
                    }
                    else
                    {
                        var result = GetParameterValue(el, field.ParameterId);
                        val = result.Item1;
                        isType = result.Item2;
                    }
                    
                    if (val == null) val = "";
                    rowData.Add(val);

                    string colName = field.GetName();
                    if (!data.IsTypeParameter.ContainsKey(colName))
                    {
                        data.IsTypeParameter[colName] = isType;
                    }
                    if (isType) data.IsTypeParameter[colName] = true;
                }
                data.Rows.Add(rowData);
            }

            return data;
        }

        private (string, bool) GetParameterValue(Element el, ElementId parameterId)
        {
            Parameter p = null;
            bool isType = false;
            
            long idValue = parameterId.Value;

            // Try to get from instance first
            if (idValue < 0)
            {
                p = el.get_Parameter((BuiltInParameter)(int)idValue);
            }
            else
            {
                try
                {
                    var paramElem = _doc.GetElement(parameterId);
                    if (paramElem != null)
                    {
                        p = el.LookupParameter(paramElem.Name);
                    }
                }
                catch { }
            }
            
            // Fallback: iterate through instance parameters
            if (p == null)
            {
                foreach (Parameter param in el.Parameters)
                {
                    if (param.Id.Value == parameterId.Value)
                    {
                        p = param;
                        break;
                    }
                }
            }

            // If not found on instance, check the element's type
            if (p == null)
            {
                ElementId typeId = el.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    Element typeElem = _doc.GetElement(typeId);
                    if (typeElem != null)
                    {
                         // Try direct lookup first
                         if (idValue < 0)
                         {
                            p = typeElem.get_Parameter((BuiltInParameter)(int)idValue);
                         }
                         else
                         {
                             var paramElem = _doc.GetElement(parameterId);
                             if (paramElem != null)
                             {
                                 p = typeElem.LookupParameter(paramElem.Name);
                             }
                         }
                         
                         // If still not found, iterate through all parameters on type
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
                         
                         if (p != null) isType = true;
                    }
                }
            }

            if (p != null)
            {
                string val = "";
                try
                {
                    if (p.HasValue)
                    {
                        switch (p.StorageType)
                        {
                            case StorageType.String:
                                val = p.AsString() ?? "";
                                break;
                            case StorageType.Integer:
                                val = p.AsInteger().ToString();
                                break;
                            case StorageType.Double:
                                val = p.AsValueString() ?? p.AsDouble().ToString();
                                break;
                            case StorageType.ElementId:
                                ElementId elemId = p.AsElementId();
                                if (elemId != null && elemId != ElementId.InvalidElementId)
                                {
                                    Element refElem = _doc.GetElement(elemId);
                                    val = refElem != null ? refElem.Name : "";
                                }
                                break;
                            default:
                                val = p.AsValueString() ?? "";
                                break;
                        }
                    }
                }
                catch
                {
                    // Last resort: try AsValueString
                    try { val = p.AsValueString() ?? ""; } catch { }
                }
                
                return (val, isType);
            }

            return ("", false);
        }

        public List<ParameterItem> GetScheduledParameters(ViewSchedule schedule)
        {
            var items = new List<ParameterItem>();
            if (schedule == null) return items;

            ScheduleDefinition def = schedule.Definition;
            var fieldIds = def.GetFieldOrder();

            foreach (var id in fieldIds)
            {
                ScheduleField field = def.GetField(id);
                if (!field.IsHidden)
                {
                    items.Add(new ParameterItem
                    {
                        Name = field.GetName(),
                        Id = id, // ScheduleFieldId
                        IsScheduled = true,
                        IsFieldId = true
                    });
                }
            }
            return items;
        }

        private bool IsParameterBoundToCategory(Document doc, SchedulableField sf, ElementId categoryId)
        {
            // If categoryId is invalid (e.g. Multi-Category), we can't filter strictly by single category
            if (categoryId == ElementId.InvalidElementId) return true;

            ElementId paramId = sf.ParameterId;
            if (paramId == ElementId.InvalidElementId) return true; // Keep intrinsic fields like "Count"

            // 1. Check if it's a BuiltInParameter
            if (paramId.Value < 0)
            {
                // Hard to check built-in parameter category binding easily without checking all definitions.
                // But usually GetSchedulableFields returns valid ones. 
                // However, things like "Room: Name" appear as BuiltIn but are related to a different category.
                // We want to exclude "Related" fields if we want strict filtering.
                // SchedulableField has no "IsRelated" property?
                // Actually, if the field is from a linked element (e.g. Room), the name usually contains ":".
                // But that's a heuristic.
                return true; 
            }

            // 2. Check Project/Shared Parameter Binding
            try
            {
                var bindingMap = doc.ParameterBindings;
                // check by iterating? No, direct lookup if we have definition.
                ParameterElement pe = doc.GetElement(paramId) as ParameterElement;
                if (pe != null)
                {
                    Definition def = pe.GetDefinition();
                    if (bindingMap.Contains(def))
                    {
                            InstanceBinding ib = bindingMap.get_Item(def) as InstanceBinding;
                            if (ib != null)
                            {
                                var cat = Category.GetCategory(doc, categoryId);
                                return ib.Categories.Contains(cat);
                            }
                            TypeBinding tb = bindingMap.get_Item(def) as TypeBinding;
                            if (tb != null)
                            {
                                var cat = Category.GetCategory(doc, categoryId);
                                return tb.Categories.Contains(cat);
                            }
                        }
                    }
                }
            catch { }

            // If we can't determine, include it to be safe.
            return true;
        }

        private bool IsParameterAvailableForCategory(Document doc, ElementId paramId, ElementId categoryId)
        {
            try
            {
                if (paramId.Value < 0) return false; // Built-in params handled separately
                if (categoryId == ElementId.InvalidElementId) return true; // Multi-category schedule

                // Find a sample element of this category and check if it has the parameter
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                
                // Try instance elements first
                Element sampleElement = collector
                    .OfCategoryId(categoryId)
                    .WhereElementIsNotElementType()
                    .FirstElement();
                
                if (sampleElement != null)
                {
                    foreach (Parameter param in sampleElement.Parameters)
                    {
                        if (param.Id == paramId) return true;
                    }
                }
                
                // If no instance found or parameter not on instances, try element types
                collector = new FilteredElementCollector(doc);
                ElementType sampleType = collector
                    .OfCategoryId(categoryId)
                    .WhereElementIsElementType()
                    .FirstElement() as ElementType;
                
                if (sampleType != null)
                {
                    foreach (Parameter param in sampleType.Parameters)
                    {
                        if (param.Id == paramId) return true;
                    }
                }
                
                // No elements found or parameter not available
                return false;
            }
            catch
            {
                return false;
            }
        }

        public List<ParameterItem> GetAvailableParameters(ViewSchedule schedule)
        {
            var items = new List<ParameterItem>();
            if (schedule == null) return items;

            ScheduleDefinition def = schedule.Definition;
            IList<SchedulableField> schedulableFields = def.GetSchedulableFields();
            
            ElementId scheduleCategoryId = def.CategoryId;
            
            // Setup robust filtering using ParameterFilterUtilities
            HashSet<ElementId> validParameterIds = null;
            if (scheduleCategoryId != ElementId.InvalidElementId)
            {
                try
                {
                    validParameterIds = new HashSet<ElementId>(
                        ParameterFilterUtilities.GetFilterableParametersInCommon(_doc, new List<ElementId> { scheduleCategoryId })
                    );
                }
                catch { }
            }

            // Get existing parameter Ids to filter duplicates
            var existingParamIds = new HashSet<ElementId>();
            var fieldIds = def.GetFieldOrder();
            foreach (var id in fieldIds)
            {
                var f = def.GetField(id);
                if (f != null) existingParamIds.Add(f.ParameterId);
            }

            foreach (var sf in schedulableFields)
            {
                // Filter out parameters that are already scheduled
                if (existingParamIds.Contains(sf.ParameterId)) continue;
                
                // Allow "Invalid" parameters (e.g. Count)
                if (sf.ParameterId == ElementId.InvalidElementId)
                {
                     items.Add(new ParameterItem
                    {
                        Name = sf.GetName(_doc),
                        Id = sf,
                        IsScheduled = false,
                        IsFieldId = false
                    });
                    continue;
                }

                bool keep = false;
                
                // 1. Built-In Parameters: Strict check against FilterUtils
                if (sf.ParameterId.Value < 0)
                {
                    if (validParameterIds != null && validParameterIds.Contains(sf.ParameterId))
                    {
                        keep = true;
                    }
                    else if (validParameterIds == null)
                    {
                        keep = true; // Fallback
                    }
                }
                else
                {
                    // 2. Project/Shared Parameters: Explicit Binding Check
                    //    Must be bound to the specific category.
                    if (scheduleCategoryId == ElementId.InvalidElementId)
                    {
                        keep = true; // Multi-Category schedule -> keep all
                    }
                    else
                    {
                        // First try: Element-based check
                        if (IsParameterAvailableForCategory(_doc, sf.ParameterId, scheduleCategoryId))
                        {
                            keep = true;
                        }
                        // Fallback: Also check ParameterFilterUtilities (covers edge cases)
                        else if (validParameterIds != null && validParameterIds.Contains(sf.ParameterId))
                        {
                            keep = true;
                        }
                    }
                }

                if (keep)
                {
                    items.Add(new ParameterItem
                    {
                        Name = sf.GetName(_doc),
                        Id = sf, // SchedulableField
                        IsScheduled = false,
                        IsFieldId = false
                    });
                }
            }
            
            return items.OrderBy(x => x.Name).ToList();
        }
    }
}
