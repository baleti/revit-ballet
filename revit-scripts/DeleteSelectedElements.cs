using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DeleteSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get selected elements
                ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
                
                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("No Selection", "Please select elements before running this command.");
                    return Result.Cancelled;
                }

                // Delete the selected elements
                int count = selectedIds.Count;
                
                using (Transaction trans = new Transaction(doc, "Delete Selected Elements"))
                {
                    trans.Start();
                    
                    ICollection<ElementId> deletedIds = doc.Delete(selectedIds);
                    
                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
