using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SwitchView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document  doc   = uidoc.Document;
        View      activeView = uidoc.ActiveView;

        string projectName = doc != null ? doc.Title : "UnknownProject";

        string logFilePath = PathHelper.GetLogViewChangesPath(projectName);

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
                              .AsEnumerable()
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
            // Try ID match first, fallback to title match (handles detached models with changed IDs)
            selectedIndex = views.FindIndex(v => v.Id == currentViewId);
            if (selectedIndex < 0)
                selectedIndex = views.FindIndex(v => v.Title == activeView.Title);
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
                {
                    selectedIndex = views.FindIndex(v => v.Id == sheet.Id);
                    // Fallback to title match if ID doesn't match
                    if (selectedIndex < 0)
                        selectedIndex = views.FindIndex(v => v.Title == sheet.Title);
                }
            }
            else
            {
                selectedIndex = views.FindIndex(v => v.Id == currentViewId);
                // Fallback to title match if ID doesn't match
                if (selectedIndex < 0)
                    selectedIndex = views.FindIndex(v => v.Title == activeView.Title);
            }
        }

        var propertyNames           = new List<string> { "SheetNumber", "Name", "ViewType" };
        var initialSelectionIndices = selectedIndex >= 0
                                        ? new List<int> { selectedIndex }
                                        : new List<int>();

        /* ─────────────────────────────────────────────────────┐
           Convert views to DataGrid format with custom logic   │
           (split sheet number and name into separate columns)  │
           ─────────────────────────────────────────────────────┘ */
        var viewDicts = new List<Dictionary<string, object>>();
        foreach (var view in views)
        {
            var dict = new Dictionary<string, object>();

            if (view is ViewSheet sheet)
            {
                dict["SheetNumber"] = sheet.SheetNumber;
                dict["Name"] = sheet.Name;
            }
            else
            {
                dict["SheetNumber"] = ""; // Empty for non-sheet views
                dict["Name"] = view.Name;
            }

            dict["ViewType"] = view.ViewType;
            dict["ElementIdObject"] = view.Id; // Required for edit functionality
            dict["__OriginalObject"] = view; // Store original object for extraction

            viewDicts.Add(dict);
        }

        /* ─────────────────────────────┐
           Show picker & handle result  │
           ─────────────────────────────┘ */
        CustomGUIs.SetCurrentUIDocument(uidoc);
        var selectedDicts = CustomGUIs.DataGrid(viewDicts, propertyNames, false, initialSelectionIndices);
        List<View> chosen = CustomGUIs.ExtractOriginalObjects<View>(selectedDicts);

        // Apply any pending edits to Revit elements
        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits())
        {
            CustomGUIs.ApplyCellEditsToEntities();

            // Update LogViewChanges to reflect renamed views/sheets
            UpdateLogViewChangesAfterRenames(doc);

            editsWereApplied = true;
        }

        if (chosen.Count == 0)
            return Result.Failed;

        // Only switch view if no edits were made (stay in current view if editing)
        if (!editsWereApplied)
        {
            uidoc.ActiveView = chosen.First();
        }
        return Result.Succeeded;
    }

    /// <summary>
    /// Updates LogViewChanges file to reflect renamed views/sheets.
    /// Call this after ApplyCellEditsToEntities() to keep the log in sync.
    /// </summary>
    private static void UpdateLogViewChangesAfterRenames(Document doc)
    {
        string projectName = doc != null ? doc.Title : "UnknownProject";
        string logFilePath = PathHelper.GetLogViewChangesPath(projectName);

        if (!System.IO.File.Exists(logFilePath))
            return;

        // Read all log entries
        var logEntries = System.IO.File.ReadAllLines(logFilePath).ToList();
        bool modified = false;

        // Update each entry with current view title
        for (int i = 0; i < logEntries.Count; i++)
        {
            string entry = logEntries[i].Trim();
            if (string.IsNullOrEmpty(entry))
                continue;

            // Parse: "ElementId Title"
            var parts = entry.Split(new[] { ' ' }, 2);
            if (parts.Length != 2)
                continue;

            // Try to parse element ID
            if (!long.TryParse(parts[0], out long elementIdValue))
                continue;

            // Get the current view/sheet from document
            ElementId elemId = elementIdValue.ToElementId();
            Element elem = doc.GetElement(elemId);

            if (elem is View view)
            {
                // Update entry with current title
                string currentTitle = view.Title;
                string newEntry = $"{elementIdValue} {currentTitle}";

                if (newEntry != entry)
                {
                    logEntries[i] = newEntry;
                    modified = true;
                }
            }
        }

        // Write back if modified
        if (modified)
        {
            System.IO.File.WriteAllLines(logFilePath, logEntries);
        }
    }
}
