using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SwitchViewByHistory : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Document doc = commandData.Application.ActiveUIDocument.Document;
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        string projectName = doc != null ? doc.Title : "UnknownProject";

        string logFilePath = PathHelper.GetLogViewChangesPath(projectName);

        if (!File.Exists(logFilePath))
        {
            TaskDialog.Show("Error", "Log file does not exist.");
            return Result.Failed;
        }

        var viewEntries = File.ReadAllLines(logFilePath)
            .AsEnumerable()
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

        var propertyNames = new List<string> { "SheetNumber", "Name", "ViewType" };
        ElementId currentViewId = uidoc.ActiveView.Id;

        // Determine the index of the currently active view in the list
        // Try ID match first, fallback to title match (handles detached models with changed IDs)
        int selectedIndex = views.FindIndex(v => v.Id == currentViewId);
        if (selectedIndex < 0)
            selectedIndex = views.FindIndex(v => v.Title == uidoc.ActiveView.Title);

        // Adjusted call to DataGrid to use a list with a single index for initial selection
        List<int> initialSelectionIndices = selectedIndex >= 0 ? new List<int> { selectedIndex } : new List<int>();

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

        CustomGUIs.SetCurrentUIDocument(uidoc);
        var selectedDicts = CustomGUIs.DataGrid(viewDicts, propertyNames, false, initialSelectionIndices);
        List<View> selectedViews = CustomGUIs.ExtractOriginalObjects<View>(selectedDicts);

        // Apply any pending edits to Revit elements
        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits())
        {
            CustomGUIs.ApplyCellEditsToEntities();

            // Update LogViewChanges to reflect renamed views/sheets
            UpdateLogViewChangesAfterRenames(doc);

            editsWereApplied = true;
        }

        if (selectedViews.Count == 0)
            return Result.Failed;

        // Only switch view if no edits were made (stay in current view if editing)
        if (!editsWereApplied)
        {
            uidoc.ActiveView = selectedViews.First();
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
