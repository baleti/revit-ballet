using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SwitchViewByHistory : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Document doc = commandData.Application.ActiveUIDocument.Document;
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        string projectName = doc != null ? doc.Title : "UnknownProject";

        string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string logFilePath = Path.Combine(appDataPath, "revit-scripts", "LogViewChanges", $"{projectName}");

        if (!File.Exists(logFilePath))
        {
            TaskDialog.Show("Error", "Log file does not exist.");
            return Result.Failed;
        }

        var viewEntries = File.ReadAllLines(logFilePath)
            .Reverse()
            .Distinct()
            .Skip(1)
            .ToList();

        var viewTitles = new List<string>();

        foreach (var entry in viewEntries)
        {
            var parts = entry.Split(new[] { ' ' }, 2); // Split into two parts: ID and Title
            if (parts.Length == 2)
            {
                viewTitles.Add(parts[1]);
            }
        }

        var allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .ToList();

        var views = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => viewTitles.Contains(v.Title))
            .OrderBy(v => viewTitles.IndexOf(v.Title))
            .ToList();

        var propertyNames = new List<string> { "Title", "ViewType" };
        ElementId currentViewId = uidoc.ActiveView.Id;

        // Determine the index of the currently active view in the list
        int selectedIndex = views.FindIndex(v => v.Id == currentViewId);

        // Adjusted call to DataGrid to use a list with a single index for initial selection
        List<int> initialSelectionIndices = selectedIndex >= 0 ? new List<int> { selectedIndex } : new List<int>();

        List<View> selectedViews = CustomGUIs.DataGrid(views, propertyNames, initialSelectionIndices); // Null for sorting as we've already sorted

        if (selectedViews.Count == 0)
            return Result.Failed;

        uidoc.ActiveView = selectedViews.First();
        return Result.Succeeded;
    }
}
