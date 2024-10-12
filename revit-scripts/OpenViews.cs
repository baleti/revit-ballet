using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get all views in the project
        List<View> views = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => !v.IsTemplate && v.Title != "Project Browser" && v.Title != "System Browser")
            .OrderBy(v => v.Title) // Sort views alphabetically by Title
            .ToList();

        List<string> properties = new List<string> { "Title", "ViewType" };

        // Determine the index of the currently active view in the sorted list
        ElementId currentViewId = uidoc.ActiveView.Id;
        int selectedIndex = views.FindIndex(v => v.Id == currentViewId);

        // Adjusted call to DataGrid to use a list with a single index for initial selection
        List<int> initialSelectionIndices = selectedIndex >= 0 ? new List<int> { selectedIndex } : new List<int>();

        var selectedViews = CustomGUIs.DataGrid<View>(views, properties, initialSelectionIndices);

        selectedViews.ForEach(view => uidoc.RequestViewChange(view));

        return Result.Succeeded;
    }
}
