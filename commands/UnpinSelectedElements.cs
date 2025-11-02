using System.Collections.Generic;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace MyRevitCommands
{
    public class UnpinSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the active document and the selected element IDs
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();

            if (selectedIds.Count == 0)
            {
                message = "Please select at least one element.";
                return Result.Failed;
            }

            // Start a transaction to modify the document
            using (Transaction tx = new Transaction(doc, "Unpin Selected Elements"))
            {
                tx.Start();

                foreach (ElementId id in selectedIds)
                {
                    Element element = doc.GetElement(id);
                    
                    // Check if the element is pinned and unpin it
                    if (element.Pinned)
                    {
                        element.Pinned = false;
                    }
                }
                
                tx.Commit();
            }

            return Result.Succeeded;
        }
    }
}
