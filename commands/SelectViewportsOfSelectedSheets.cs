using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace RevitBallet.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class SelectViewportsOfSelectedSheets : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get current selection using SelectionModeManager
            ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();

            if (currentSelection.Count == 0)
            {
                TaskDialog.Show("Select Viewports of Selected Sheets", "No elements selected. Please select one or more sheets.");
                return Result.Cancelled;
            }

            // Filter to only sheets
            var sheets = currentSelection
                .Select(id => doc.GetElement(id))
                .OfType<ViewSheet>()
                .ToList();

            if (sheets.Count == 0)
            {
                TaskDialog.Show("Select Viewports of Selected Sheets", "No sheets found in selection. Please select one or more sheets.");
                return Result.Cancelled;
            }

            // Collect all viewports from selected sheets
            var viewportIds = new List<ElementId>();
            foreach (ViewSheet sheet in sheets)
            {
                // Get all viewports on this sheet
                var sheetViewports = sheet.GetAllViewports();
                viewportIds.AddRange(sheetViewports);
            }

            if (viewportIds.Count == 0)
            {
                TaskDialog.Show("Select Viewports of Selected Sheets",
                    $"No viewports found on the selected {sheets.Count} sheet(s).");
                return Result.Cancelled;
            }

            // Update selection to viewports using SelectionModeManager
            uidoc.SetSelectionIds(viewportIds);

            return Result.Succeeded;
        }
    }
}
