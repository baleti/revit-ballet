using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectOpenedViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData,
                          ref string            message,
                          ElementSet            elements)
    {
        /* ───────────────────────────────────────
           Basic handles
           ───────────────────────────────────── */
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;

        string projectName = doc?.Title ?? "UnknownProject";

        string logFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "revit-scripts",
            "LogViewChanges",
            projectName);

        /* ───────────────────────────────────────
           Validate the log
           ───────────────────────────────────── */
        if (!File.Exists(logFilePath))
        {
            TaskDialog.Show("Select Opened Views",
                            "No log file was found for this project.");
            return Result.Failed;
        }

        /* ───────────────────────────────────────
           Load & parse log entries
           Each line:  "<ElementId> <Title>"
           newest entry is last in the file
           ───────────────────────────────────── */
        HashSet<ElementId> loggedIds   = new HashSet<ElementId>();
        HashSet<string>    loggedTitles = new HashSet<string>();

        foreach (string raw in File.ReadAllLines(logFilePath).Reverse()) // newest first
        {
            string line = raw.Trim();
            if (line.Length == 0) continue;

            string[] parts = line.Split(new[] { ' ' }, 2);
            if (parts.Length != 2) continue;

            if (int.TryParse(parts[0], out int idInt))
                loggedIds.Add(new ElementId(idInt));

            loggedTitles.Add(parts[1]);
        }

        if (loggedIds.Count == 0)
        {
            TaskDialog.Show("Select Opened Views",
                            "The log is empty – nothing to pick.");
            return Result.Failed;
        }

        /* ───────────────────────────────────────
           Map non-sheet views to the sheets they’re placed on
           ───────────────────────────────────── */
        Dictionary<ElementId, ViewSheet> viewToSheetMap = new Dictionary<ElementId, ViewSheet>();

        foreach (ViewSheet sheet in new FilteredElementCollector(doc)
                                     .OfClass(typeof(ViewSheet))
                                     .Cast<ViewSheet>())
        {
            foreach (ElementId vpId in sheet.GetAllViewports())
            {
                if (doc.GetElement(vpId) is Viewport vp)
                    viewToSheetMap[vp.ViewId] = sheet;
            }
        }

        /* ───────────────────────────────────────
           Collect all document views
           and keep only those in the log
           ───────────────────────────────────── */
        List<Dictionary<string, object>> gridData       = new List<Dictionary<string, object>>();
        Dictionary<string, View>         titleToViewMap = new Dictionary<string, View>();

        foreach (View v in new FilteredElementCollector(doc)
                            .OfClass(typeof(View))
                            .Cast<View>())
        {
            // Skip templates & non-selectable view types
            if (v.IsTemplate ||
                v.ViewType == ViewType.Legend ||
                v.ViewType == ViewType.Schedule ||
                v.ViewType == ViewType.ProjectBrowser ||
                v.ViewType == ViewType.SystemBrowser)
                continue;

            // Keep only logged views
            if (!loggedIds.Contains(v.Id) && !loggedTitles.Contains(v.Title))
                continue;

            // Sheet information
            string sheetInfo;
            if (v is ViewSheet sheetView)
            {
                sheetInfo = $"{sheetView.SheetNumber} – {sheetView.Name}";
            }
            else if (viewToSheetMap.TryGetValue(v.Id, out ViewSheet hostSheet))
            {
                sheetInfo = $"{hostSheet.SheetNumber} – {hostSheet.Name}";
            }
            else
            {
                sheetInfo = "Not Placed";
            }

            // Compose row for the grid
            titleToViewMap[v.Title] = v;    // assumes titles are unique
            gridData.Add(new Dictionary<string, object>
            {
                { "Title",  v.Title     },
                { "Sheet",  sheetInfo   }
            });
        }

        if (gridData.Count == 0)
        {
            TaskDialog.Show("Select Opened Views",
                            "None of the logged views exist in this model.");
            return Result.Failed;
        }

        /* ───────────────────────────────────────
           Let the user pick
           ───────────────────────────────────── */
        List<string> columns = new List<string> { "Title", "Sheet" };

        // Returns rows the user checked, or null / empty if cancelled
        List<Dictionary<string, object>> chosen =
            CustomGUIs.DataGrid(gridData, columns, spanAllScreens: false);

        if (chosen == null || !chosen.Any())
            return Result.Cancelled;

        /* ───────────────────────────────────────
           Add chosen views to the current selection
           ───────────────────────────────────── */
        ICollection<ElementId> current = uidoc.GetSelectionIds();

        foreach (Dictionary<string, object> row in chosen)
        {
            string t   = row["Title"].ToString();
            ElementId id = titleToViewMap[t].Id;
            if (!current.Contains(id))
                current.Add(id);
        }

        uidoc.SetSelectionIds(current);
        return Result.Succeeded;
    }
}
