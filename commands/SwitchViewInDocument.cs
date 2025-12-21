using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SwitchViewInDocument : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        // Get SessionId (ProcessId) and DocumentTitle for querying history
        string sessionId = RevitBallet.LogViewChanges.GetSessionId();
        string documentTitle = doc.Title;

        // Get view history from database
        var history = LogViewChangesDatabase.GetViewHistoryForDocument(sessionId, documentTitle, limit: 1000);

        // Build list of views from history
        var viewsInHistory = new List<View>();
        var seenViewIds = new HashSet<ElementId>();

        if (history.Count == 0)
        {
            // No history yet - document just opened and first view hasn't been logged yet
            // Fallback: show only the currently active view
            viewsInHistory.Add(activeView);
            seenViewIds.Add(activeView.Id);
        }
        else
        {
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

            // Skip duplicates
            if (seenViewIds.Contains(viewId))
                continue;

            Element viewElement = doc.GetElement(viewId);

            // If element not found by ID, try finding by title
            if (viewElement == null || !(viewElement is View))
            {
                viewElement = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .FirstOrDefault(v => v.Title == entry.ViewTitle);
            }

            if (viewElement != null && viewElement is View view)
            {
                viewsInHistory.Add(view);
                seenViewIds.Add(view.Id);
            }
        }

            if (viewsInHistory.Count == 0)
            {
                // History exists but no matching views found - fallback to active view
                viewsInHistory.Add(activeView);
                seenViewIds.Add(activeView.Id);
            }
        }

        // Get browser organization columns
        List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
            BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, viewsInHistory);

        // Pre-select the current view
        int selectedIndex = -1;
        ElementId currentViewId = activeView.Id;

        if (activeView is ViewSheet)
        {
            selectedIndex = viewsInHistory.FindIndex(v => v.Id == currentViewId);
            if (selectedIndex < 0)
                selectedIndex = viewsInHistory.FindIndex(v => v.Title == activeView.Title);
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
                    selectedIndex = viewsInHistory.FindIndex(v => v.Id == sheet.Id);
                    if (selectedIndex < 0)
                        selectedIndex = viewsInHistory.FindIndex(v => v.Title == sheet.Title);
                }
            }
            else
            {
                selectedIndex = viewsInHistory.FindIndex(v => v.Id == currentViewId);
                if (selectedIndex < 0)
                    selectedIndex = viewsInHistory.FindIndex(v => v.Title == activeView.Title);
            }
        }

        var initialSelectionIndices = selectedIndex >= 0
                                        ? new List<int> { selectedIndex }
                                        : new List<int>();

        // Convert views to DataGrid format
        var viewDicts = new List<Dictionary<string, object>>();
        foreach (var view in viewsInHistory)
        {
            var dict = new Dictionary<string, object>();

            // Add browser organization columns first
            BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, view, doc, browserColumns);

            // Then add standard columns
            if (view is ViewSheet sheet)
            {
                dict["SheetNumber"] = sheet.SheetNumber;
                dict["Name"] = sheet.Name;
                dict["Sheet"] = ""; // Empty for sheets
            }
            else
            {
                dict["SheetNumber"] = "";
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

            dict["ViewType"] = view.ViewType;
            dict["ElementIdObject"] = view.Id;
            dict["__OriginalObject"] = view;

            viewDicts.Add(dict);
        }

        // Sort by browser columns if available
        if (browserColumns != null && browserColumns.Count > 0)
        {
            viewDicts = BrowserOrganizationHelper.SortByBrowserColumns(viewDicts, browserColumns);

            // Recalculate selectedIndex after sorting to match the active view or its sheet
            ElementId targetViewId = null;
            if (activeView is ViewSheet)
            {
                targetViewId = activeView.Id;
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
                        targetViewId = sheet.Id;
                }
                else
                {
                    targetViewId = activeView.Id;
                }
            }

            if (targetViewId != null)
            {
                selectedIndex = viewDicts.FindIndex(row =>
                {
                    if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View v)
                        return v.Id == targetViewId;
                    return false;
                });

                initialSelectionIndices = selectedIndex >= 0
                    ? new List<int> { selectedIndex }
                    : new List<int>();
            }
        }

        // Build property names - browser columns first, then standard columns
        var propertyNames = new List<string>();
        propertyNames.AddRange(browserColumns.Select(bc => bc.Name));
        propertyNames.Add("SheetNumber");
        propertyNames.Add("Name");
        propertyNames.Add("ViewType");
        propertyNames.Add("Sheet");

        // Show picker & handle result
        CustomGUIs.SetCurrentUIDocument(uidoc);
        var selectedDicts = CustomGUIs.DataGrid(viewDicts, propertyNames, false, initialSelectionIndices);
        List<View> chosen = CustomGUIs.ExtractOriginalObjects<View>(selectedDicts);

        // Apply any pending edits to Revit elements
        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits())
        {
            CustomGUIs.ApplyCellEditsToEntities();
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
}
