using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace YourNamespace
{
    [Transaction(TransactionMode.ReadOnly)]
    public class SelectViewportsOfViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the current UIDocument and Document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get the currently selected elements
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                message = "Please select one or more views in the project browser or drawing.";
                return Result.Failed;
            }

            // Filter selection to collect view elements (excluding sheets)
            HashSet<ElementId> selectedViewIds = new HashSet<ElementId>();
            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);
                // Check if the element is a view but not a ViewSheet
                if (element is View view && !(view is ViewSheet))
                {
                    selectedViewIds.Add(view.Id);
                }
            }

            if (selectedViewIds.Count == 0)
            {
                message = "No valid views selected.";
                return Result.Failed;
            }

            // Collect all viewport elements from the document (placed on sheets)
            FilteredElementCollector viewportCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport));

            List<ElementId> viewportsToSelect = new List<ElementId>();

            // For each viewport, check if its associated view is in the selection
            foreach (Viewport vp in viewportCollector)
            {
                if (selectedViewIds.Contains(vp.ViewId))
                {
                    viewportsToSelect.Add(vp.Id);
                }
            }

            if (viewportsToSelect.Count == 0)
            {
                message = "None of the selected views are placed on any sheet.";
                return Result.Failed;
            }

            // Set the selection in the UIDocument to the found viewport elements
            uidoc.Selection.SetElementIds(viewportsToSelect);

            return Result.Succeeded;
        }
    }
}
