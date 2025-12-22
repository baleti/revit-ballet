using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[TransactionAttribute(TransactionMode.Manual)]
public class SelectViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the currently active view
        View activeView = doc.ActiveView;

        // Create a mapping for views that are placed on sheets (non-sheet views)
        Dictionary<ElementId, ViewSheet> viewToSheetMap = new Dictionary<ElementId, ViewSheet>();
        FilteredElementCollector sheetCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet));
        foreach (ViewSheet sheet in sheetCollector)
        {
            foreach (ElementId viewportId in sheet.GetAllViewports())
            {
                Viewport viewport = doc.GetElement(viewportId) as Viewport;
                if (viewport != null)
                {
                    viewToSheetMap[viewport.ViewId] = sheet;
                }
            }
        }

        // Get all views in the project, including view sheets.
        List<View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => !v.IsTemplate &&
                        v.ViewType != ViewType.Legend &&
                        v.ViewType != ViewType.Schedule &&
                        v.ViewType != ViewType.ProjectBrowser &&
                        v.ViewType != ViewType.SystemBrowser)
            .ToList();

        // Get browser organization columns (pass views so it can detect sheets vs views)
        List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
            BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, allViews);

        // Prepare data for the data grid and map view titles to view objects.
        List<Dictionary<string, object>> viewData = new List<Dictionary<string, object>>();
        Dictionary<string, View> titleToViewMap = new Dictionary<string, View>();

        foreach (View view in allViews)
        {
            // Assuming titles are unique; otherwise, you might need to use a different key.
            titleToViewMap[view.Title] = view;
            Dictionary<string, object> viewInfo = new Dictionary<string, object>();

            // Add browser organization columns first
            BrowserOrganizationHelper.AddBrowserColumnsToDict(viewInfo, view, doc, browserColumns);

            // Then add standard columns - differentiate between ViewSheet and regular views
            if (view is ViewSheet viewSheet)
            {
                // For a sheet, show its sheet number and name separately
                viewInfo["SheetNumber"] = viewSheet.SheetNumber;
                viewInfo["Name"] = viewSheet.Name;
                viewInfo["Sheet"] = ""; // Empty for sheets
            }
            else
            {
                viewInfo["SheetNumber"] = ""; // Empty for non-sheet views
                viewInfo["Name"] = view.Name;

                // Check if view is placed on a sheet
                if (viewToSheetMap.TryGetValue(view.Id, out ViewSheet sheet))
                {
                    // For non-sheet views placed on a sheet, display the sheet info.
                    viewInfo["Sheet"] = $"{sheet.SheetNumber} - {sheet.Name}";
                }
                else
                {
                    viewInfo["Sheet"] = "Not Placed";
                }
            }

            viewInfo["__OriginalObject"] = view; // Store original object for comparison

            viewData.Add(viewInfo);
        }

        // Sort by browser organization columns (if any), otherwise by Name
        if (browserColumns != null && browserColumns.Count > 0)
        {
            viewData = BrowserOrganizationHelper.SortByBrowserColumns(viewData, browserColumns);
        }
        else
        {
            viewData = viewData.OrderBy(v => v["Name"].ToString()).ToList();
        }

        // Find the index of the active view after sorting
        int sortedActiveViewIndex = -1;
        if (activeView != null)
        {
            sortedActiveViewIndex = viewData.FindIndex(row =>
            {
                if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View v)
                    return v.Id == activeView.Id;
                return false;
            });
        }

        // Define the column headers - browser columns first, then standard columns.
        List<string> columns = new List<string>();
        columns.AddRange(browserColumns.Select(bc => bc.Name));
        columns.Add("SheetNumber");
        columns.Add("Name");
        columns.Add("Sheet");

        // Prepare initial selection indices (if active view was found)
        List<int> initialSelection = null;
        if (sortedActiveViewIndex >= 0)
        {
            initialSelection = new List<int> { sortedActiveViewIndex };
        }

        // Show the selection dialog (using your custom GUI).
        List<Dictionary<string, object>> selectedViews = CustomGUIs.DataGrid(
            viewData,
            columns,
            false,  // Don't span all screens.
            initialSelection  // Pass the initial selection
        );

        // If the user made a selection, add those elements to the current selection.
        if (selectedViews != null && selectedViews.Any())
        {
            // Get the current selection
            ICollection<ElementId> currentSelectionIds = uidoc.GetSelectionIds();

            // Get the ElementIds of the views selected in the dialog using __OriginalObject
            List<ElementId> newViewIds = selectedViews
                .Where(v => v.ContainsKey("__OriginalObject") && v["__OriginalObject"] is View)
                .Select(v => (v["__OriginalObject"] as View).Id)
                .ToList();

            // Add the new views to the current selection
            foreach (ElementId id in newViewIds)
            {
                if (!currentSelectionIds.Contains(id))
                {
                    currentSelectionIds.Add(id);
                }
            }

            // Update the selection with the combined set of elements
            uidoc.SetSelectionIds(currentSelectionIds);

            return Result.Succeeded;
        }

        return Result.Cancelled;
    }
}
