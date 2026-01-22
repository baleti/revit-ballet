using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CloseViewsInDocument : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        // Get SessionId for database operations
        string sessionId = RevitBallet.LogViewChanges.GetSessionId();

        // Get all currently open views
        IList<UIView> UIViews = uidoc.GetOpenUIViews();

        // Build openViews list with special handling:
        // - If a view is placed on a sheet, show the SHEET instead of the view
        // - Avoid duplicates (use HashSet to track added view IDs)
        List<View> openViews = new List<View>();
        HashSet<ElementId> addedViewIds = new HashSet<ElementId>();

        foreach (UIView UIview in UIViews)
        {
            View view = doc.GetElement(UIview.ViewId) as View;
            if (view != null &&
                !view.IsTemplate &&
                view.ViewType != ViewType.ProjectBrowser &&
                view.ViewType != ViewType.SystemBrowser)
            {
                // If this is a non-sheet view, check if it's placed on a sheet
                if (!(view is ViewSheet))
                {
                    var viewport = new FilteredElementCollector(doc)
                        .OfClass(typeof(Viewport))
                        .Cast<Viewport>()
                        .FirstOrDefault(vp => vp.ViewId.Equals(view.Id));

                    if (viewport != null)
                    {
                        // View is placed on a sheet - add the SHEET instead
                        ViewSheet containingSheet = doc.GetElement(viewport.SheetId) as ViewSheet;
                        if (containingSheet != null && !addedViewIds.Contains(containingSheet.Id))
                        {
                            openViews.Add(containingSheet);
                            addedViewIds.Add(containingSheet.Id);
                        }
                        // Skip adding the view itself
                        continue;
                    }
                }

                // Add view/sheet (if not already added)
                if (!addedViewIds.Contains(view.Id))
                {
                    openViews.Add(view);
                    addedViewIds.Add(view.Id);
                }
            }
        }

        if (openViews.Count == 0)
        {
            TaskDialog.Show("Info", "No open views to close.");
            return Result.Failed;
        }

        // Get browser organization columns
        List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
            BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, openViews);

        // Prepare data for the grid
        List<Dictionary<string, object>> gridData = new List<Dictionary<string, object>>();

        foreach (View v in openViews)
        {
            var dict = new Dictionary<string, object>();

            // Add browser organization columns first
            BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, v, doc, browserColumns);

            // Then add standard columns
            if (v is ViewSheet sheet)
            {
                dict["SheetNumber"] = sheet.SheetNumber;
                dict["Name"] = sheet.Name;
                dict["Sheet"] = ""; // Empty for sheets
            }
            else
            {
                dict["SheetNumber"] = "";
                dict["Name"] = v.Name;

                // Check if view is placed on a sheet
                var viewport = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .FirstOrDefault(vp => vp.ViewId.Equals(v.Id));

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

            dict["ElementIdObject"] = v.Id;
            dict["__OriginalObject"] = v;

            gridData.Add(dict);
        }

        // Sort by browser organization columns (if any), otherwise by Title
        // Use "SheetNumber" as tiebreaker for multi-column sorting (sheets sort by number, views group together)
        if (browserColumns != null && browserColumns.Count > 0)
        {
            gridData = BrowserOrganizationHelper.SortByBrowserColumns(gridData, browserColumns, tiebreakerColumn: "SheetNumber");
        }
        else
        {
            gridData = gridData.OrderBy(row =>
            {
                if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View v)
                    return v.Title;
                return "";
            }).ToList();
        }

        // Column headers
        List<string> columns = new List<string>();
        columns.AddRange(browserColumns.Select(bc => bc.Name));
        columns.Add("SheetNumber");
        columns.Add("Name");
        columns.Add("Sheet");

        // ─────────────────────────────────────────────────────────────
        // Figure out which row should be pre-selected (current view or its sheet)
        // Note: openViews already replaces views-on-sheets with their containing sheets
        // ─────────────────────────────────────────────────────────────
        int selectedIndex = -1;
        ElementId targetViewId = null;

        if (activeView is ViewSheet)
        {
            targetViewId = activeView.Id;
        }
        else
        {
            // Check if active view is placed on a sheet
            Viewport vp = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .FirstOrDefault(vpt => vpt.ViewId.Equals(activeView.Id));

            if (vp != null)
            {
                ViewSheet containingSheet = doc.GetElement(vp.SheetId) as ViewSheet;
                if (containingSheet != null)
                {
                    // View is on a sheet - select the sheet
                    // (The sheet will be in openViews list because we replaced the view with the sheet)
                    targetViewId = containingSheet.Id;
                }
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
                    return view.Id.Equals(targetViewId);
                return false;
            });
        }

        List<int> initialSelectionIndices = selectedIndex >= 0
            ? new List<int> { selectedIndex }
            : new List<int>();

        // Show the grid
        CustomGUIs.SetCurrentUIDocument(uidoc);
        var selectedDicts = CustomGUIs.DataGrid(gridData, columns, false, initialSelectionIndices);
        List<View> selectedViews = CustomGUIs.ExtractOriginalObjects<View>(selectedDicts);

        if (selectedViews.Count == 0)
            return Result.Failed;

        // Close selected views
        foreach (View view in selectedViews)
        {
            foreach (UIView openedUIView in uidoc.GetOpenUIViews())
            {
                if (openedUIView.ViewId.Equals(view.Id))
                {
                    // Remove the closed view from history
                    try
                    {
                        LogViewChangesDatabase.RemoveViewFromHistory(sessionId, doc.Title, view.Title);
                    }
                    catch
                    {
                        // Silently fail - don't interrupt the close operation
                    }

                    openedUIView.Close();
                }
            }
        }

        return Result.Succeeded;
    }
}
