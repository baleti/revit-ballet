using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenViewsInDocument : IExternalCommand
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
            .ToList();

        // Get browser organization columns (pass views so it can detect sheets vs views)
        List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
            BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, allViews);

        // ─────────────────────────────────────────────────────────────
        // 2. Prepare data for the grid with SheetNumber, Name, ViewType
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> gridData =
            new List<Dictionary<string, object>>();

        foreach (View v in allViews)
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
                dict["SheetNumber"] = ""; // Empty for non-sheet views
                dict["Name"] = v.Name;

                // Check if view is placed on a sheet
                var viewport = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .FirstOrDefault(vp => vp.ViewId == v.Id);

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

            dict["ElementIdObject"] = v.Id; // Required for edit functionality
            dict["__OriginalObject"] = v; // Store original object for extraction

            gridData.Add(dict);
        }

        // Sort by browser organization columns (if any), otherwise by Title
        if (browserColumns != null && browserColumns.Count > 0)
        {
            gridData = BrowserOrganizationHelper.SortByBrowserColumns(gridData, browserColumns);
        }
        else
        {
            gridData = gridData.OrderBy(row =>
            {
                // Extract view to get Title for sorting
                if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View v)
                    return v.Title;
                return "";
            }).ToList();
        }

        // Column headers (order determines column order) - browser columns first
        List<string> columns = new List<string>();
        columns.AddRange(browserColumns.Select(bc => bc.Name));
        columns.Add("SheetNumber");
        columns.Add("Name");
        columns.Add("Sheet");

        // ─────────────────────────────────────────────────────────────
        // 3. Figure out which row should be pre-selected (after sorting)
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
}
