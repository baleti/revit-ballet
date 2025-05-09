using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[TransactionAttribute(TransactionMode.Manual)]
public class SelectViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Create a mapping for views that are placed on sheets (non-sheet views)
        Dictionary<ElementId, ViewSheet> viewToSheetMap = new Dictionary<ElementId, ViewSheet>();
        FilteredElementCollector sheetCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet));
        foreach (ViewSheet sheet in sheetCollector)
        {
            foreach (ElementId viewportId in sheet.GetAllViewports())
            {
                Viewport viewport = doc.GetElement(viewportId) as Viewport;
                if (viewport != null)
                {
                    viewToSheetMap[viewport.ViewId] = sheet;
                }
            }
        }

        // Get all views in the project, including view sheets.
        FilteredElementCollector viewCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(View));

        // Prepare data for the data grid and map view titles to view objects.
        List<Dictionary<string, object>> viewData = new List<Dictionary<string, object>>();
        Dictionary<string, View> titleToViewMap = new Dictionary<string, View>();

        foreach (View view in viewCollector.Cast<View>())
        {
            // Skip view templates, legends, schedules, project browser, and system browser.
            if (view.IsTemplate || 
                view.ViewType == ViewType.Legend || 
                view.ViewType == ViewType.Schedule ||
                view.ViewType == ViewType.ProjectBrowser ||
                view.ViewType == ViewType.SystemBrowser)
                continue;

            string sheetInfo = string.Empty;
            if (view is ViewSheet)
            {
                // For a sheet, show its own sheet number and name.
                ViewSheet viewSheet = view as ViewSheet;
                sheetInfo = $"{viewSheet.SheetNumber} - {viewSheet.Name}";
            }
            else if (viewToSheetMap.TryGetValue(view.Id, out ViewSheet sheet))
            {
                // For non-sheet views placed on a sheet, display the sheet info.
                sheetInfo = $"{sheet.SheetNumber} - {sheet.Name}";
            }
            else
            {
                sheetInfo = "Not Placed";
            }

            // Assuming titles are unique; otherwise, you might need to use a different key.
            titleToViewMap[view.Title] = view;
            Dictionary<string, object> viewInfo = new Dictionary<string, object>
            {
                { "Title", view.Title },
                { "Sheet", sheetInfo }
            };
            viewData.Add(viewInfo);
        }

        // Define the column headers.
        List<string> columns = new List<string> { "Title", "Sheet" };

        // Show the selection dialog (using your custom GUI).
        List<Dictionary<string, object>> selectedViews = CustomGUIs.DataGrid(
            viewData,
            columns,
            false  // Don't span all screens.
        );

        // If the user made a selection, add those elements to the current selection.
        if (selectedViews != null && selectedViews.Any())
        {
            // Get the current selection
            ICollection<ElementId> currentSelectionIds = uidoc.Selection.GetElementIds();
            
            // Get the ElementIds of the views selected in the dialog
            List<ElementId> newViewIds = selectedViews
                .Select(v => titleToViewMap[v["Title"].ToString()].Id)
                .ToList();
                
            // Add the new views to the current selection
            foreach (ElementId id in newViewIds)
            {
                if (!currentSelectionIds.Contains(id))
                {
                    currentSelectionIds.Add(id);
                }
            }
            
            // Update the selection with the combined set of elements
            uidoc.Selection.SetElementIds(currentSelectionIds);
            
            return Result.Succeeded;
        }

        return Result.Cancelled;
    }
}
