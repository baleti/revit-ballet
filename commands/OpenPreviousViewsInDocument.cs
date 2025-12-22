using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenPreviousViewsInDocument : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        // Get DocumentTitle for querying history
        string currentSessionId = RevitBallet.LogViewChanges.GetSessionId();
        string documentTitle = doc.Title;

        // Get ALL view history from database for this document (all sessions, sorted by timestamp DESC)
        var history = LogViewChangesDatabase.GetAllViewHistoryForDocument(documentTitle, limit: 1000);

        if (history.Count == 0)
        {
            TaskDialog.Show("Info", "No previously opened views found in this document.");
            return Result.Failed;
        }

        // Build list of view entries from history (DO NOT deduplicate - keep all session records)
        var viewEntries = new List<(View view, ViewHistoryEntry entry)>();

        foreach (var entry in history)
        {
            ElementId viewId;
            try
            {
                viewId = entry.ViewId.ToElementId();
            }
            catch (Exception)
            {
                // Skip invalid ViewId entries (e.g., 0, -1, or corrupted data)
                continue;
            }

            Element viewElement = doc.GetElement(viewId);

            // If element not found by ID, try finding by title
            if (viewElement == null || !(viewElement is View))
            {
                viewElement = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .FirstOrDefault(v => v.Title == entry.ViewTitle);
            }

            if (viewElement != null && viewElement is View view &&
                !view.IsTemplate &&
                view.ViewType != ViewType.ProjectBrowser &&
                view.ViewType != ViewType.SystemBrowser)
            {
                viewEntries.Add((view, entry));
            }
        }

        if (viewEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No matching views found in current document.");
            return Result.Failed;
        }

        // Get browser organization columns
        List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
            BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, viewEntries.Select(ve => ve.view).Distinct().ToList());

        // ─────────────────────────────────────────────────────────────
        // Prepare data for the grid
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> gridData =
            new List<Dictionary<string, object>>();

        foreach (var (view, entry) in viewEntries)
        {
            var dict = new Dictionary<string, object>();

            // Add browser organization columns first
            BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, view, doc, browserColumns);

            // Add session tracking column
            bool isCurrentSession = entry.SessionId == currentSessionId;
            dict["Session"] = isCurrentSession ? "Current" : entry.SessionId.Substring(0, Math.Min(8, entry.SessionId.Length));

            // Add timestamp column
            dict["Last Opened"] = entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss");
            dict["__Timestamp"] = entry.Timestamp; // Store for sorting

            // Then add standard columns
            if (view is ViewSheet sheet)
            {
                dict["SheetNumber"] = sheet.SheetNumber;
                dict["Name"] = sheet.Name;
                dict["Sheet"] = ""; // Empty for sheets
            }
            else
            {
                dict["SheetNumber"] = ""; // Empty for non-sheet views
                dict["Name"] = view.Name;

                // Check if view is placed on a sheet
                var viewport = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .FirstOrDefault(vp => vp.ViewId == view.Id);

                if (viewport != null)
                {
                    ViewSheet containingSheet = doc.GetElement(viewport.SheetId) as ViewSheet;
                    dict["Sheet"] = containingSheet != null ? containingSheet.Title : "";
                }
                else
                {
                    dict["Sheet"] = ""; // Empty for views not on sheets
                }
            }

            dict["ElementIdObject"] = view.Id; // Required for edit functionality
            dict["__OriginalObject"] = view; // Store original object for extraction

            gridData.Add(dict);
        }

        // Sort by timestamp descending (most recently opened first) - already sorted by database query
        gridData = gridData.OrderByDescending(row =>
        {
            if (row.ContainsKey("__Timestamp"))
                return (DateTime)row["__Timestamp"];
            return DateTime.MinValue;
        }).ToList();

        // Column headers (order determines column order) - browser columns first, then Session, then timestamp, then standard
        List<string> columns = new List<string>();
        columns.AddRange(browserColumns.Select(bc => bc.Name));
        columns.Add("Session");
        columns.Add("Last Opened");
        columns.Add("SheetNumber");
        columns.Add("Name");
        columns.Add("Sheet");

        // ─────────────────────────────────────────────────────────────
        // Figure out which row should be pre-selected (active view)
        // ─────────────────────────────────────────────────────────────
        int selectedIndex = -1;
        ElementId targetViewId = null;

        if (activeView is ViewSheet)
        {
            targetViewId = activeView.Id;
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
                    targetViewId = containingSheet.Id;
            }

            if (targetViewId == null) // not on a sheet
                targetViewId = activeView.Id;
        }

        // Find the index in sorted gridData
        if (targetViewId != null)
        {
            selectedIndex = gridData.FindIndex(row =>
            {
                if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View view)
                    return view.Id == targetViewId;
                return false;
            });
        }

        List<int> initialSelectionIndices = selectedIndex >= 0
            ? new List<int> { selectedIndex }
            : new List<int>();

        // ─────────────────────────────────────────────────────────────
        // Show the grid
        // ─────────────────────────────────────────────────────────────
        CustomGUIs.SetCurrentUIDocument(uidoc);
        List<Dictionary<string, object>> selectedRows =
            CustomGUIs.DataGrid(gridData, columns, false, initialSelectionIndices);

        // ─────────────────────────────────────────────────────────────
        // Apply any pending edits to Revit elements
        // ─────────────────────────────────────────────────────────────
        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits())
        {
            CustomGUIs.ApplyCellEditsToEntities();
            editsWereApplied = true;
        }

        // ─────────────────────────────────────────────────────────────
        // Open every selected view (sheet or model view)
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
}
