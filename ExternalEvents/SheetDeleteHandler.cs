using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ProSchedules.ExternalEvents
{
    public class SheetDeleteHandler : IExternalEventHandler
    {
        public List<ElementId> SheetIdsToDelete { get; set; } = new List<ElementId>();
        public event Action<int, int, string, List<ElementId>> OnDeleteFinished;

        public void Execute(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            if (uidoc == null) return;
            var doc = uidoc.Document;

            int successCount = 0;
            int failCount = 0;
            string errorMsg = "";
            var deletedIds = new List<ElementId>();

            if (SheetIdsToDelete == null || SheetIdsToDelete.Count == 0)
            {
                OnDeleteFinished?.Invoke(0, 0, "", deletedIds);
                return;
            }

            using (var t = new Transaction(doc, "Delete Sheets"))
            {
                t.Start();
                try
                {
                    foreach (var id in SheetIdsToDelete)
                    {
                        try
                        {
                            var sheet = doc.GetElement(id) as ViewSheet;
                            if (sheet == null)
                            {
                                failCount++;
                                errorMsg = $"Could not find sheet with ID {id}";
                                continue;
                            }

                            doc.Delete(id);
                            deletedIds.Add(id);
                            successCount++;
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
                    OnDeleteFinished?.Invoke(successCount, failCount, errorMsg, deletedIds);
                }
            }
        }

        public string GetName()
        {
            return "Sheet Delete Handler";
        }
    }
}
