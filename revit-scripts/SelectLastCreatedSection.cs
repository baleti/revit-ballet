using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using System.Collections.Generic;

namespace RevitAPICommands
{
    [Transaction(TransactionMode.ReadOnly)]
    public class SelectLastCreatedSection : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Collect only section elements (ViewSection) and order them by ElementId descending.
            var sections = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSection))
                .Cast<ViewSection>()
                .OrderByDescending(vs => vs.Id.IntegerValue);

            // Get the section with the highest ElementId.
            ViewSection lastSection = sections.FirstOrDefault();

            if (lastSection == null)
            {
                TaskDialog.Show("Select Last Section", "No section element found in the document.");
                return Result.Failed;
            }

            // Add the found section to the current selection.
            ICollection<ElementId> currentSelection = uidoc.Selection.GetElementIds();
            List<ElementId> newSelection = currentSelection.ToList();
            newSelection.Add(lastSection.Id);
            uidoc.Selection.SetElementIds(newSelection);

            return Result.Succeeded;
        }
    }
}
