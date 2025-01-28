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

        // Create view-to-sheet mapping first
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

        // Get all views in the project
        FilteredElementCollector viewCollector = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Views)
            .WhereElementIsNotElementType();

        // Create a list to store view data and a dictionary to map titles to views
        List<Dictionary<string, object>> viewData = new List<Dictionary<string, object>>();
        Dictionary<string, View> titleToViewMap = new Dictionary<string, View>();

        // Collect view data
        foreach (View view in viewCollector.Cast<View>())
        {
            // Skip view templates, legends, and schedules
            if (view.IsTemplate || view.ViewType == ViewType.Legend || view.ViewType == ViewType.Schedule)
                continue;

            // Get sheet info using the mapping
            string sheetInfo = "Not Placed";
            if (viewToSheetMap.TryGetValue(view.Id, out ViewSheet sheet))
            {
                sheetInfo = $"{sheet.SheetNumber} - {sheet.Name}";
            }

            // Store the view in our title mapping
            titleToViewMap[view.Title] = view;

            // Create dictionary for current view
            Dictionary<string, object> viewInfo = new Dictionary<string, object>
            {
                { "Title", view.Title },
                { "Sheet", sheetInfo }
            };

            viewData.Add(viewInfo);
        }

        // Define column headers
        List<string> columns = new List<string> { "Title", "Sheet" };

        // Show selection dialog
        List<Dictionary<string, object>> selectedViews = CustomGUIs.DataGrid(
            viewData,
            columns,
            false  // Don't span all screens
        );

        // If user selected views, add them to selection
        if (selectedViews != null && selectedViews.Any())
        {
            // Get the ElementIds using the title mapping
            List<ElementId> viewIds = selectedViews
                .Select(v => titleToViewMap[v["Title"].ToString()].Id)
                .ToList();

            // Set the selection
            uidoc.Selection.SetElementIds(viewIds);

            return Result.Succeeded;
        }

        return Result.Cancelled;
    }
}
