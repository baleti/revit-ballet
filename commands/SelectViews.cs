using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[TransactionAttribute(TransactionMode.Manual)]
[CommandMeta("")]
public class SelectViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the currently active view
        View activeView = doc.ActiveView;

        // Map viewId -> list of (Viewport, ViewSheet) placements. Legends appear on multiple sheets.
        var viewToViewportsMap = new Dictionary<ElementId, List<(Viewport Viewport, ViewSheet Sheet)>>();
        FilteredElementCollector sheetCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet));
        foreach (ViewSheet sheet in sheetCollector)
        {
            foreach (ElementId viewportId in sheet.GetAllViewports())
            {
                Viewport viewport = doc.GetElement(viewportId) as Viewport;
                if (viewport != null)
                {
                    if (!viewToViewportsMap.ContainsKey(viewport.ViewId))
                        viewToViewportsMap[viewport.ViewId] = new List<(Viewport, ViewSheet)>();
                    viewToViewportsMap[viewport.ViewId].Add((viewport, sheet));
                }
            }
        }

        // Get all views in the project, including view sheets and legends.
        List<View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => !v.IsTemplate &&
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
            titleToViewMap[view.Title] = view;

            if (view is ViewSheet viewSheet)
            {
                Dictionary<string, object> viewInfo = new Dictionary<string, object>();
                BrowserOrganizationHelper.AddBrowserColumnsToDict(viewInfo, view, doc, browserColumns);
                viewInfo["SheetNumber"] = viewSheet.SheetNumber;
                viewInfo["Name"] = viewSheet.Name;
                viewInfo["Sheet"] = "";
                viewInfo["ElementIdObject"] = view.Id;
                viewData.Add(viewInfo);
            }
            else if (view.ViewType == ViewType.Legend)
            {
                // Each legend placement on a sheet is a separate row (keyed by viewport Id).
                if (viewToViewportsMap.TryGetValue(view.Id, out var placements))
                {
                    foreach (var (viewport, sheet) in placements)
                    {
                        Dictionary<string, object> viewInfo = new Dictionary<string, object>();
                        BrowserOrganizationHelper.AddBrowserColumnsToDict(viewInfo, view, doc, browserColumns);
                        viewInfo["SheetNumber"] = sheet.SheetNumber;
                        viewInfo["Name"] = view.Name;
                        viewInfo["Sheet"] = $"{sheet.SheetNumber} - {sheet.Name}";
                        viewInfo["ElementIdObject"] = viewport.Id;
                        viewData.Add(viewInfo);
                    }
                }
                else
                {
                    Dictionary<string, object> viewInfo = new Dictionary<string, object>();
                    BrowserOrganizationHelper.AddBrowserColumnsToDict(viewInfo, view, doc, browserColumns);
                    viewInfo["SheetNumber"] = "";
                    viewInfo["Name"] = view.Name;
                    viewInfo["Sheet"] = "Not Placed";
                    viewInfo["ElementIdObject"] = view.Id;
                    viewData.Add(viewInfo);
                }
            }
            else
            {
                Dictionary<string, object> viewInfo = new Dictionary<string, object>();
                BrowserOrganizationHelper.AddBrowserColumnsToDict(viewInfo, view, doc, browserColumns);
                viewInfo["SheetNumber"] = "";
                viewInfo["Name"] = view.Name;
                if (viewToViewportsMap.TryGetValue(view.Id, out var placements))
                    viewInfo["Sheet"] = $"{placements[0].Sheet.SheetNumber} - {placements[0].Sheet.Name}";
                else
                    viewInfo["Sheet"] = "Not Placed";
                viewInfo["ElementIdObject"] = view.Id;
                viewData.Add(viewInfo);
            }
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
                if (row.ContainsKey("ElementIdObject") && row["ElementIdObject"] is ElementId id)
                    return id == activeView.Id;
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

        // Enable automatic edit application
        CustomGUIs.SetCurrentUIDocument(uidoc);

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

            // Get the ElementIds of the views selected in the dialog using ElementIdObject
            List<ElementId> newViewIds = selectedViews
                .Where(v => v.ContainsKey("ElementIdObject") && v["ElementIdObject"] is ElementId)
                .Select(v => v["ElementIdObject"] as ElementId)
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
