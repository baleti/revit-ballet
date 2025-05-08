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
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document  doc   = uidoc.Document;
        View      activeView = uidoc.ActiveView;

        string projectName = doc != null ? doc.Title : "UnknownProject";

        string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string logFilePath = Path.Combine(appDataPath, "revit-scripts", "LogViewChanges", projectName);

        /* ─────────────────────────────┐
           Validate log-file existence  │
           ─────────────────────────────┘ */
        if (!File.Exists(logFilePath))
        {
            TaskDialog.Show("View Switch", "Log file does not exist for this project.");
            return Result.Failed;
        }

        /* ─────────────────────────────┐
           Load & de-duplicate entries  │
           ─────────────────────────────┘ */
        var viewEntries = File.ReadAllLines(logFilePath)
                              .Reverse()         // newest first
                              .Select(l => l.Trim())
                              .Where(l => l.Length > 0)
                              .Distinct()
                              .ToList();

        /* Terminate if the log is empty */
        if (viewEntries.Count == 0)
        {
            TaskDialog.Show("View Switch", "The log file is empty – nothing to switch to.");
            return Result.Failed;
        }

        /* ─────────────────────────────┐
           Extract titles from entries  │
           ─────────────────────────────┘ */
        var viewTitles = new List<string>();
        foreach (string entry in viewEntries)
        {
            var parts = entry.Split(new[] { ' ' }, 2);  // "ID  Title"
            if (parts.Length == 2)
                viewTitles.Add(parts[1]);
        }

        /* ─────────────────────────────┐
           Collect views in the model   │
           ─────────────────────────────┘ */
        var allViews = new FilteredElementCollector(doc)
                       .OfClass(typeof(View))
                       .Cast<View>()
                       .ToList();

        var views = allViews
                    .Where(v => viewTitles.Contains(v.Title))
                    .OrderBy(v => v.Title)
                    .ToList();

        /* If nothing matches the log, stop */
        if (views.Count == 0)
        {
            TaskDialog.Show("View Switch", "No matching views found in this model.");
            return Result.Failed;
        }

        /* ─────────────────────────────┐
           Pre-select the current view  │
           ─────────────────────────────┘ */
        int selectedIndex = -1;
        ElementId currentViewId = activeView.Id;

        if (activeView is ViewSheet)
        {
            selectedIndex = views.FindIndex(v => v.Id == currentViewId);
        }
        else
        {
            // Check if the active view is placed on a sheet
            var viewports = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .Where(vp => vp.ViewId == currentViewId)
                            .ToList();

            if (viewports.Count > 0)
            {
                ViewSheet sheet = doc.GetElement(viewports.First().SheetId) as ViewSheet;
                if (sheet != null)
                    selectedIndex = views.FindIndex(v => v.Id == sheet.Id);
            }
            else
            {
                selectedIndex = views.FindIndex(v => v.Id == currentViewId);
            }
        }

        var propertyNames           = new List<string> { "Title", "ViewType" };
        var initialSelectionIndices = selectedIndex >= 0
                                        ? new List<int> { selectedIndex }
                                        : new List<int>();

        /* ─────────────────────────────┐
           Show picker & handle result  │
           ─────────────────────────────┘ */
        List<View> chosen = CustomGUIs.DataGrid(views, propertyNames, initialSelectionIndices);

        if (chosen.Count == 0)
            return Result.Failed;

        uidoc.ActiveView = chosen.First();
        return Result.Succeeded;
    }
}
