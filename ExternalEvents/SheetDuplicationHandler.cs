using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;

namespace ProSchedules.ExternalEvents
{
    // Mode definition
    public enum SheetDuplicateMode
    {
        EmptySheet,
        WithSheetDetailing,
        WithViews
    }

    public class SheetDuplicationHandler : IExternalEventHandler
    {
        public List<SheetItem> SheetsToDuplicate { get; set; } = new List<SheetItem>();
        public int NumberOfCopies { get; set; } = 1;
        public SheetDuplicateMode DuplicateMode { get; set; } = SheetDuplicateMode.EmptySheet;

        public bool KeepLegends { get; set; }
        public bool KeepSchedules { get; set; }
        public bool CopyRevisions { get; set; }
        public bool CopyParameters { get; set; }

        // New property for pending sheet data
        public List<SheetItem> PendingSheetData { get; set; } = new List<SheetItem>();

        public event Action<int, int, string, List<ElementId>> OnDuplicationFinished;

        public void Execute(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            if (uidoc == null) return;
            var doc = uidoc.Document;

            // Check if we have pending data (new workflow) or old workflow
            bool usePendingData = PendingSheetData != null && PendingSheetData.Count > 0;

            if (!usePendingData && (SheetsToDuplicate == null || SheetsToDuplicate.Count == 0))
            {
                TaskDialog.Show("Debug", "Handler: No sheets to duplicate.");
                return;
            }

            using (Transaction t = new Transaction(doc, "Pro Schedules"))
            {
                t.Start();

                int successCount = 0;
                int failCount = 0;
                string errorMsg = "";
                List<ElementId> newSheetIds = new List<ElementId>();

                try
                {
                    if (usePendingData)
                    {
                        // New workflow: Use pending sheet data
                        foreach (var pendingData in PendingSheetData)
                        {
                            try
                            {
                                ViewSheet sourceSheet = doc.GetElement(pendingData.SourceSheetId) as ViewSheet;
                                if (sourceSheet == null)
                                {
                                    failCount++;
                                    continue;
                                }

                                // Duplicate the sheet using the options stored in pendingData
                                ViewSheet newSheet = DuplicateSheetLogic(doc, sourceSheet, pendingData.DuplicateOptions);

                                if (newSheet != null)
                                {
                                    // Apply custom names from pending data
                                    newSheet.SheetNumber = pendingData.SheetNumber;
                                    newSheet.Name = pendingData.Name;

                                    newSheetIds.Add(newSheet.Id);
                                    successCount++;
                                }
                                else
                                {
                                    failCount++;
                                    errorMsg = $"Failed to create sheet '{pendingData.Name}'";
                                }
                            }
                            catch (Exception ex)
                            {
                                failCount++;
                                errorMsg = ex.Message;
                                System.Diagnostics.Debug.WriteLine($"Error creating pending sheet: {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        // Old workflow: Use SheetsToDuplicate with NumberOfCopies
                        foreach (var sheetItem in SheetsToDuplicate)
                        {
                            ViewSheet sourceSheet = doc.GetElement(sheetItem.Id) as ViewSheet;
                            if (sourceSheet == null)
                            {
                                failCount++;
                                continue;
                            }

                            for (int i = 0; i < NumberOfCopies; i++)
                            {
                                try
                                {
                                    var options = new DuplicateOptions
                                    {
                                        DuplicateMode = this.DuplicateMode,
                                        KeepLegends = this.KeepLegends,
                                        KeepSchedules = this.KeepSchedules,
                                        CopyRevisions = this.CopyRevisions,
                                        CopyParameters = this.CopyParameters
                                    };

                                    ViewSheet newSheet = DuplicateSheetLogic(doc, sourceSheet, options);

                                    if (newSheet != null)
                                    {
                                        // Auto-generate names
                                        try
                                        {
                                            string baseName = sourceSheet.Name;
                                            string newName = $"{baseName} - Copy {i + 1}";
                                            newSheet.Name = newName;
                                            newSheet.SheetNumber = sourceSheet.SheetNumber + $"-C{i + 1}";
                                        }
                                        catch
                                        {
                                            // If setting name fails, let Revit keep default
                                        }

                                        newSheetIds.Add(newSheet.Id);
                                        successCount++;
                                    }
                                    else
                                    {
                                        failCount++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    failCount++;
                                    errorMsg = ex.Message;
                                    System.Diagnostics.Debug.WriteLine($"Error duplicating sheet {sourceSheet.Name}: {ex.Message}");
                                }
                            }
                        }
                    }

                    t.Commit();
                }
                catch (Exception ex)
                {
                    failCount++;
                    errorMsg = "Critical Error: " + ex.Message;
                }
                finally
                {
                    OnDuplicationFinished?.Invoke(successCount, failCount, errorMsg, newSheetIds);
                }
            }
        }

        private ViewSheet DuplicateSheetLogic(Document doc, ViewSheet sourceSheet, DuplicateOptions options)
        {
            ViewSheet newSheet = null;

            // 1. Determine Revit ViewDuplicateOption for the SHEET
            ViewDuplicateOption sheetOption = ViewDuplicateOption.WithDetailing;

            // 2. Try Standard Duplication
            bool standardDuplicateSuccess = false;

            // Check supports
            if (sourceSheet.CanViewBeDuplicated(sheetOption))
            {
                ElementId newSheetId = sourceSheet.Duplicate(sheetOption);
                if (newSheetId != ElementId.InvalidElementId)
                {
                    newSheet = doc.GetElement(newSheetId) as ViewSheet;
                    standardDuplicateSuccess = true;
                }
            }
            else if (sheetOption == ViewDuplicateOption.WithDetailing && sourceSheet.CanViewBeDuplicated(ViewDuplicateOption.Duplicate))
            {
                // Fallback: Try duplicating without detailing if WithDetailing failed
                ElementId newSheetId = sourceSheet.Duplicate(ViewDuplicateOption.Duplicate);
                if (newSheetId != ElementId.InvalidElementId)
                {
                    newSheet = doc.GetElement(newSheetId) as ViewSheet;
                    standardDuplicateSuccess = true;

                    // Manually copy detailing (lines/text) since we fell back to Empty Duplicate
                    CopySheetDetailing(doc, sourceSheet, newSheet);
                }
            }

            // 3. Manual Fallback (if standard failed)
            if (!standardDuplicateSuccess)
            {
                newSheet = ManualDuplicateSheet(doc, sourceSheet);
                if (newSheet != null)
                {
                    // Always try to copy detailing for manual sheets
                    CopySheetDetailing(doc, sourceSheet, newSheet);
                }
            }

            if (newSheet == null)
            {
                return null;
            }

            // 4. Copy Parameters
            if (options.CopyParameters)
            {
                CopyElementParameters(doc, sourceSheet, newSheet);
                CopyTitleBlockInstanceParameters(doc, sourceSheet, newSheet);
            }

            // 5. Copy Revisions
            if (options.CopyRevisions)
            {
                var revs = sourceSheet.GetAdditionalRevisionIds();
                if (revs.Count > 0)
                {
                    newSheet.SetAdditionalRevisionIds(revs);
                }
            }

            // 6. Handle Views (Legends, Schedules, Model Views)
            bool includeModelViews = (options.DuplicateMode == SheetDuplicateMode.WithViews);
            DuplicateAndProSchedules(doc, sourceSheet, newSheet, includeModelViews, options.KeepLegends, options.KeepSchedules);

            return newSheet;
        }

        private void CopyElementParameters(Document doc, Element source, Element target)
        {
            foreach (Parameter p in source.Parameters)
            {
                if (p.IsReadOnly) continue;
                if (p.Definition.Name == "Sheet Number" || p.Definition.Name == "Sheet Name") continue;

                Parameter targetParam = target.LookupParameter(p.Definition.Name);
                if (targetParam != null && !targetParam.IsReadOnly)
                {
                   SetParameterValue(targetParam, p);
                }
            }
        }

        private void CopyTitleBlockInstanceParameters(Document doc, ViewSheet source, ViewSheet target)
        {
            var sourceTB = new FilteredElementCollector(doc, source.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).FirstElement();
            var targetTB = new FilteredElementCollector(doc, target.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).FirstElement();

            if (sourceTB != null && targetTB != null)
            {
                 CopyElementParameters(doc, sourceTB, targetTB);
            }
        }

        private void SetParameterValue(Parameter target, Parameter source)
        {
            switch (source.StorageType)
            {
                case StorageType.Double:
                    target.Set(source.AsDouble());
                    break;
                case StorageType.Integer:
                    target.Set(source.AsInteger());
                    break;
                case StorageType.String:
                    target.Set(source.AsString());
                    break;
                case StorageType.ElementId:
                    target.Set(source.AsElementId());
                    break;
            }
        }

        private void DuplicateAndProSchedules(Document doc, ViewSheet source, ViewSheet target, bool includeModelViews, bool keepLegends, bool keepSchedules)
        {
            // 1. Handle Schedules (Explicitly via ScheduleSheetInstance)
            if (keepSchedules)
            {
                var scheduleInstances = new FilteredElementCollector(doc, source.Id)
                    .OfClass(typeof(ScheduleSheetInstance))
                    .Cast<ScheduleSheetInstance>()
                    .ToList();

                foreach (var ssi in scheduleInstances)
                {
                    try
                    {
                        if (ssi.ScheduleId != ElementId.InvalidElementId)
                        {
                            // Create new instance of the SAME schedule
                            ScheduleSheetInstance.Create(doc, target.Id, ssi.ScheduleId, ssi.Point);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error placing schedule: {ex.Message}");
                    }
                }
            }

            // 2. Handle Legends and Model Views
            foreach (ElementId viewId in source.GetAllPlacedViews())
            {
                try {
                View sourceView = doc.GetElement(viewId) as View;
                if (sourceView == null) continue;

                // Skip Schedules here as they are handled above
                if (sourceView.ViewType == ViewType.Schedule) continue;

                // A. Legends
                if (sourceView.ViewType == ViewType.Legend)
                {
                    if (keepLegends)
                    {
                        Viewport.Create(doc, target.Id, sourceView.Id, GetViewportCenter(doc, source, sourceView.Id));
                    }
                }
                // C. Model Views (Plans, Sections etc) & Drafting Views
                else
                {
                    if (includeModelViews)
                    {
                        if (sourceView.CanViewBeDuplicated(ViewDuplicateOption.WithDetailing))
                        {
                            ElementId newViewId = sourceView.Duplicate(ViewDuplicateOption.WithDetailing);
                            Viewport.Create(doc, target.Id, newViewId, GetViewportCenter(doc, source, sourceView.Id));
                        }
                    }
                }
                } catch(Exception innerEx) {
                     System.Diagnostics.Debug.WriteLine($"Error placing view: {innerEx.Message}");
                }
            }
        }

        private XYZ GetViewportCenter(Document doc, ViewSheet sheet, ElementId viewId)
        {
            // 1. Try finding Viewport
            var viewport = new FilteredElementCollector(doc, sheet.Id)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .FirstOrDefault(vp => vp.ViewId == viewId);
            
            if (viewport != null) return viewport.GetBoxCenter();
            
            // 2. Try finding ScheduleSheetInstance
            var scheduleInstance = new FilteredElementCollector(doc, sheet.Id)
                .OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>()
                .FirstOrDefault(ssi => ssi.ScheduleId == viewId);
                
            if (scheduleInstance != null) return scheduleInstance.Point;
            
            return XYZ.Zero;
        }

        private ViewSheet ManualDuplicateSheet(Document doc, ViewSheet source)
        {
            // 1. Find Title Block
            var sourceTB = new FilteredElementCollector(doc, source.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .FirstElement() as FamilyInstance;

            ElementId titleBlockId = ElementId.InvalidElementId;
            if (sourceTB != null)
            {
                titleBlockId = sourceTB.GetTypeId();
            }
            
            // 2. Create Sheet
            ViewSheet newSheet;
            if (titleBlockId != ElementId.InvalidElementId)
            {
                newSheet = ViewSheet.Create(doc, titleBlockId);
            }
            else
            {
                newSheet = ViewSheet.Create(doc, ElementId.InvalidElementId);
            }
            
            // 3. Try setting name/number automatically (Revit does this, but we can try to match pattern if needed)
            // For now, let Revit handle default naming "Unnamed", "A101" etc.
            
            return newSheet;
        }

        private void CopySheetDetailing(Document doc, ViewSheet source, ViewSheet target)
        {
            // Copy Detail Lines, Text Notes, Generic Annotations directly on the sheet
            // Note: ElementId.Value is available since Revit 2024.
            
            var viewSpecificElements = new FilteredElementCollector(doc, source.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null)
                .Where(e => {
                    long catId = e.Category.Id.Value;
                    return catId == (long)BuiltInCategory.OST_Lines ||
                           catId == (long)BuiltInCategory.OST_TextNotes ||
                           catId == (long)BuiltInCategory.OST_GenericAnnotation ||
                           catId == (long)BuiltInCategory.OST_Dimensions ||
                           catId == (long)BuiltInCategory.OST_Tags;
                })
                .Select(e => e.Id)
                .ToList();

            if (viewSpecificElements.Count > 0)
            {
                ElementTransformUtils.CopyElements(source, viewSpecificElements, target, Transform.Identity, new CopyPasteOptions());
            }
        }

        public string GetName()
        {
            return "Sheet Duplication Handler";
        }
    }
}
