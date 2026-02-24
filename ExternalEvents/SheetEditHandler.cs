using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;

namespace ProSchedules.ExternalEvents
{
    public class SheetEditHandler : IExternalEventHandler
    {
        public List<SheetItem> SheetsToEdit { get; set; } = new List<SheetItem>();
        public event Action<int, int, string> OnEditFinished;

        public void Execute(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            if (uidoc == null) return;
            var doc = uidoc.Document;

            int successCount = 0;
            int failCount = 0;
            string errorMsg = "";

            using (Transaction t = new Transaction(doc, "Update Sheet Names/Numbers"))
            {
                t.Start();

                try
                {
                    foreach (var sheetItem in SheetsToEdit)
                    {
                        try
                        {
                            var sheet = doc.GetElement(sheetItem.Id) as ViewSheet;
                            if (sheet != null)
                            {
                                sheet.SheetNumber = sheetItem.SheetNumber;
                                sheet.Name = sheetItem.Name;
                                successCount++;
                            }
                            else
                            {
                                failCount++;
                                errorMsg = $"Could not find sheet with ID {sheetItem.Id}";
                            }
                        }
                        catch (Exception ex)
                        {
                            failCount++;
                            errorMsg = ex.Message;
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
                    OnEditFinished?.Invoke(successCount, failCount, errorMsg);
                }
            }
        }

        public string GetName()
        {
            return "Sheet Edit Handler";
        }
    }
}
