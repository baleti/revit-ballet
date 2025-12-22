using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace YourNamespace
{
    [Transaction(TransactionMode.ReadOnly)]
    public class SelectViewsOfSelectedViewports : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the current UIDocument and Document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get the currently selected elements
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            if (selectedIds.Count == 0)
            {
                message = "Please select one or more viewports on a sheet.";
                return Result.Failed;
            }

            // Filter selection to collect viewport elements
            HashSet<ElementId> selectedViewportIds = new HashSet<ElementId>();
            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);
                // Check if the element is a viewport
                if (element is Viewport)
                {
                    selectedViewportIds.Add(element.Id);
                }
            }

            if (selectedViewportIds.Count == 0)
            {
                message = "No valid viewports selected.";
                return Result.Failed;
            }

            // Collect the views referenced by the selected viewports
            List<ElementId> viewsToSelect = new List<ElementId>();

            foreach (ElementId vpId in selectedViewportIds)
            {
                Viewport vp = doc.GetElement(vpId) as Viewport;
                if (vp != null)
                {
                    // Get the view referenced by this viewport
                    ElementId viewId = vp.ViewId;
                    if (viewId != null && viewId != ElementId.InvalidElementId)
                    {
                        viewsToSelect.Add(viewId);
                    }
                }
            }

            if (viewsToSelect.Count == 0)
            {
                message = "No valid views found for the selected viewports.";
                return Result.Failed;
            }

            // Set the selection in the UIDocument to the found view elements
            uidoc.SetSelectionIds(viewsToSelect);

            return Result.Succeeded;
        }
    }
}
