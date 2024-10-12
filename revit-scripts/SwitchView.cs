using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SwitchView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Document doc = commandData.Application.ActiveUIDocument.Document;
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        View activeView = uidoc.ActiveView;
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

        var views = allViews
            .Where(v => viewTitles.Contains(v.Title))
            .OrderBy(v => v.Title) // Sort views alphabetically by Title
            .ToList();

        var propertyNames = new List<string> { "Title", "ViewType" };

        // Check if the active view is on a sheet (i.e., it is a Viewport on a ViewSheet)
        ElementId currentViewId = activeView.Id;
        int selectedIndex = -1;

        if (activeView is ViewSheet)
        {
            // If it's a sheet, find the active sheet in the view list
            selectedIndex = views.FindIndex(v => v.Id == currentViewId);
        }
        else
        {
            // Check if it's a view on a sheet (i.e., a Viewport)
            var viewports = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .Where(vp => vp.ViewId == currentViewId)
                .ToList();

            if (viewports.Count > 0)
            {
                // If we find a viewport, get the associated sheet
                Viewport activeViewport = viewports.First();
                ViewSheet containingSheet = doc.GetElement(activeViewport.SheetId) as ViewSheet;
                if (containingSheet != null)
                {
                    selectedIndex = views.FindIndex(v => v.Id == containingSheet.Id);
                }
            }
            else
            {
                // If it's not on a sheet, just select the active view
                selectedIndex = views.FindIndex(v => v.Id == currentViewId);
            }
        }

        List<int> initialSelectionIndices = selectedIndex >= 0 ? new List<int> { selectedIndex } : new List<int>();

        List<View> selectedViews = CustomGUIs.DataGrid(views, propertyNames, initialSelectionIndices);

        if (selectedViews.Count == 0)
            return Result.Failed;

        uidoc.ActiveView = selectedViews.First();
        return Result.Succeeded;
    }
}
