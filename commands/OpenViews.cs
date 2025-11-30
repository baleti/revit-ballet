using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;
        View       activeView = uidoc.ActiveView;

        // ─────────────────────────────────────────────────────────────
        // 1. Collect every non-template, non-browser view (incl. sheets)
        // ─────────────────────────────────────────────────────────────
        List<View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v =>
                   !v.IsTemplate &&
                   v.ViewType != ViewType.ProjectBrowser &&
                   v.ViewType != ViewType.SystemBrowser)
            .OrderBy(v => v.Title)
            .ToList();

        // ─────────────────────────────────────────────────────────────
        // 2. Prepare data for the grid with SheetNumber, Name, ViewType
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> gridData =
            new List<Dictionary<string, object>>();

        foreach (View v in allViews)
        {
            var dict = new Dictionary<string, object>();

            if (v is ViewSheet sheet)
            {
                dict["SheetNumber"] = sheet.SheetNumber;
                dict["Name"] = sheet.Name;
            }
            else
            {
                dict["SheetNumber"] = ""; // Empty for non-sheet views
                dict["Name"] = v.Name;
            }

            dict["ViewType"] = v.ViewType;
            dict["ElementIdObject"] = v.Id; // Required for edit functionality
            dict["__OriginalObject"] = v; // Store original object for extraction

            gridData.Add(dict);
        }

        // Column headers (order determines column order)
        List<string> columns = new List<string> { "SheetNumber", "Name", "ViewType" };

        // ─────────────────────────────────────────────────────────────
        // 3. Figure out which row should be pre-selected
        // ─────────────────────────────────────────────────────────────
        int selectedIndex = -1;

        if (activeView is ViewSheet)
        {
            selectedIndex = allViews.FindIndex(v => v.Id == activeView.Id);
        }
        else
        {
            Viewport vp = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .FirstOrDefault(vpt => vpt.ViewId == activeView.Id);

            if (vp != null)
            {
                ViewSheet containingSheet = doc.GetElement(vp.SheetId) as ViewSheet;
                if (containingSheet != null)
                    selectedIndex = allViews.FindIndex(v => v.Id == containingSheet.Id);
            }

            if (selectedIndex == -1) // not on a sheet
                selectedIndex = allViews.FindIndex(v => v.Id == activeView.Id);
        }

        List<int> initialSelectionIndices = selectedIndex >= 0
            ? new List<int> { selectedIndex }
            : new List<int>();

        // ─────────────────────────────────────────────────────────────
        // 4. Show the grid
        // ─────────────────────────────────────────────────────────────
        CustomGUIs.SetCurrentUIDocument(uidoc);
        List<Dictionary<string, object>> selectedRows =
            CustomGUIs.DataGrid(gridData, columns, false, initialSelectionIndices);

        // ─────────────────────────────────────────────────────────────
        // 5. Apply any pending edits to Revit elements
        // ─────────────────────────────────────────────────────────────
        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits())
        {
            CustomGUIs.ApplyCellEditsToEntities();

            // Update LogViewChanges to reflect renamed views/sheets
            UpdateLogViewChangesAfterRenames(doc);

            editsWereApplied = true;
        }

        // ─────────────────────────────────────────────────────────────
        // 6. Open every selected view (sheet or model view)
        //    BUT skip if user made edits (stay in current view instead)
        // ─────────────────────────────────────────────────────────────
        if (!editsWereApplied)
        {
            List<View> selectedViews = CustomGUIs.ExtractOriginalObjects<View>(selectedRows);

            if (selectedViews != null && selectedViews.Any())
            {
                foreach (View view in selectedViews)
                {
                    uidoc.RequestViewChange(view);
                }
            }
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
