using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ListSheetsWithSelectedLegend : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the active Revit document
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the currently selected legend
        var selectedIds = uidoc.Selection.GetElementIds();
        if (selectedIds.Count != 1)
        {
            TaskDialog.Show("Error", "Please select one legend.");
            return Result.Failed;
        }

        // Get the selected viewport
        ElementId viewportId = selectedIds.First();
        Viewport viewport = doc.GetElement(viewportId) as Viewport;

        if (viewport == null)
        {
            TaskDialog.Show("Error", "The selected element is not a viewport.");
            return Result.Failed;
        }

        // Get the view (legend) associated with the viewport
        ElementId legendId = viewport.ViewId;
        View legendView = doc.GetElement(legendId) as View;

        if (legendView == null || legendView.ViewType != ViewType.Legend)
        {
            TaskDialog.Show("Error", "The selected viewport does not contain a valid legend.");
            return Result.Failed;
        }

        // Collect all sheets that contain the selected legend
        IEnumerable<ViewSheet> filteredSheets = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Sheets)
            .WhereElementIsNotElementType()
            .Cast<ViewSheet>()
            .Where(viewSheet => viewSheet.GetAllPlacedViews().Contains(legendId));

        if (!filteredSheets.Any())
        {
            TaskDialog.Show("No Sheets", "The selected legend is not placed on any sheets.");
            return Result.Succeeded;
        }

        // Prepare the data for the DataGrid
        List<Dictionary<string, object>> sheetEntries = new List<Dictionary<string, object>>();
        foreach (ViewSheet sheet in filteredSheets)
        {
            Dictionary<string, object> entry = new Dictionary<string, object>
            {
                { "Sheet Name", sheet.Name },
                { "Sheet Number", sheet.SheetNumber },
                { "Id", sheet.Id.IntegerValue }
            };
            sheetEntries.Add(entry);
        }

        List<string> propertyNames = new List<string> { "Sheet Name", "Sheet Number" };

        // Display the DataGrid and get the user's selection
        var selectedSheetEntries = CustomGUIs.DataGrid(sheetEntries, propertyNames, false);
        if (selectedSheetEntries == null || selectedSheetEntries.Count == 0)
        {
            TaskDialog.Show("Error", "No sheet selected.");
            return Result.Cancelled;
        }

        // Get the selected sheet's ID
        int selectedSheetId = (int)selectedSheetEntries.First()["Id"];
        ElementId selectedSheetElementId = new ElementId(selectedSheetId);

        // Open the selected sheet
        ViewSheet selectedSheet = doc.GetElement(selectedSheetElementId) as ViewSheet;
        if (selectedSheet != null)
        {
            uidoc.ActiveView = selectedSheet;
        }
        else
        {
            TaskDialog.Show("Error", "Failed to open the selected sheet.");
            return Result.Failed;
        }

        return Result.Succeeded;
    }
}
