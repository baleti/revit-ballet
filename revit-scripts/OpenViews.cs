using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenViews : IExternalCommand
{
    // A simple wrapper class to expose View properties and a custom parameter for the DataGrid
    public class ViewInfo
    {
        public string Title { get; set; }
        public string SheetFolder { get; set; }
        // Keep a reference to the underlying Revit View so we can open it.
        public View RevitView { get; set; }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Grab all non-template, non-browser views.
        List<View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => !v.IsTemplate && v.Title != "Project Browser" && v.Title != "System Browser")
            .OrderBy(v => v.Title) // Sort views alphabetically by Title
            .ToList();

        // Convert each View to a ViewInfo, including the "Sheet Folder" parameter (if present).
        List<ViewInfo> viewInfoList = allViews.Select(v =>
        {
            // Attempt to read a parameter called "Sheet Folder"
            Parameter sheetFolderParam = v.LookupParameter("Sheet Folder");
            string sheetFolderValue = sheetFolderParam?.AsString() ?? string.Empty;

            return new ViewInfo
            {
                Title       = v.Title,
                SheetFolder = sheetFolderValue,
                RevitView   = v
            };
        }).ToList();

        // Determine the index of the currently active view in the new list.
        ElementId currentViewId = uidoc.ActiveView.Id;
        int selectedIndex = viewInfoList.FindIndex(v => v.RevitView.Id == currentViewId);

        // Create a single initial selection if we found a match.
        List<int> initialSelectionIndices = selectedIndex >= 0 ? new List<int> { selectedIndex } : new List<int>();

        // Specify which properties to show in the DataGrid:
        List<string> properties = new List<string> { "Title", "SheetFolder" };

        // Call your DataGrid, which presumably uses reflection to display columns for each property
        var selectedItems = CustomGUIs.DataGrid<ViewInfo>(viewInfoList, properties, initialSelectionIndices);

        // Request a view change for each selected item
        selectedItems.ForEach(vInfo => uidoc.RequestViewChange(vInfo.RevitView));

        return Result.Succeeded;
    }
}
