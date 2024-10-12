using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class MoveSelectedViewsToSheet : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Check if the active view is a ViewSheet
        ViewSheet activeSheet = doc.ActiveView as ViewSheet;
        if (activeSheet == null)
        {
            message = "Please run this command on a sheet view.";
            return Result.Failed;
        }

        // Get selected viewports from the active sheet
        List<Viewport> selectedViewports = GetSelectedViewports(uidoc);
        if (selectedViewports == null || selectedViewports.Count == 0)
        {
            message = "Please select one or more viewports on the sheet.";
            return Result.Failed;
        }

        // Get the list of available sheets to move the views to
        List<ViewSheet> sheets = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet))
            .Cast<ViewSheet>()
            .Where(sheet => sheet.Id != activeSheet.Id)  // Exclude the current sheet
            .ToList();

        // Find the index of the current active sheet
        int activeSheetIndex = sheets.FindIndex(sheet => sheet.Id == activeSheet.Id);

        // Prepare data for the GUI
        List<Dictionary<string, object>> sheetEntries = sheets.Select(sheet => new Dictionary<string, object>
        {
            {"Sheet Name", sheet.Name },
            {"Sheet Number", sheet.SheetNumber }
        }).ToList();

        // Display the sheet selection dialog with the active sheet pre-selected
        var selectedSheets = CustomGUIs.DataGrid(sheetEntries, new List<string> { "Sheet Name", "Sheet Number" }, false, new List<int> { activeSheetIndex });

        if (selectedSheets == null || selectedSheets.Count == 0)
        {
            message = "No sheet was selected.";
            return Result.Cancelled;
        }

        // Get the selected sheet name and number from the selection
        string selectedSheetName = selectedSheets[0]["Sheet Name"].ToString();
        string selectedSheetNumber = selectedSheets[0]["Sheet Number"].ToString();

        // Retrieve the corresponding ViewSheet from the document
        ViewSheet targetSheet = sheets.FirstOrDefault(sheet => sheet.Name == selectedSheetName && sheet.SheetNumber == selectedSheetNumber);
        if (targetSheet == null)
        {
            message = "Failed to find the selected sheet.";
            return Result.Failed;
        }

        // Move all selected views to the new sheet at their original locations
        using (Transaction t = new Transaction(doc, "Move Selected Views to Sheet"))
        {
            t.Start();

            foreach (var viewport in selectedViewports)
            {
                // Get the view from the selected viewport
                View viewToMove = doc.GetElement(viewport.ViewId) as View;
                if (viewToMove == null)
                {
                    message = "One or more selected views are not valid.";
                    return Result.Failed;
                }

                // Get the location of the viewport on the active sheet
                XYZ originalLocation = viewport.GetBoxCenter();

                // Delete the old viewport
                doc.Delete(viewport.Id);

                // Create a new viewport on the target sheet at the same location
                Viewport newViewport = Viewport.Create(doc, targetSheet.Id, viewToMove.Id, originalLocation);
            }

            t.Commit();
        }

        return Result.Succeeded;
    }

    // Helper method to get the selected viewports
    private List<Viewport> GetSelectedViewports(UIDocument uidoc)
    {
        // Get selected elements
        ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
        List<Viewport> viewports = new List<Viewport>();

        foreach (ElementId id in selectedIds)
        {
            Element element = uidoc.Document.GetElement(id);
            if (element is Viewport viewport)
            {
                viewports.Add(viewport);
            }
        }

        return viewports;
    }
}
