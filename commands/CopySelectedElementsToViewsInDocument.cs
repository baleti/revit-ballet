using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class CopySelectedElementsToViewsInDocument : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        // Get the current document
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = doc.ActiveView;

        // Get the selected elements in the active view
        ICollection<ElementId> selectedElementIds = uidoc.GetSelectionIds();

        if (!selectedElementIds.Any())
        {
            TaskDialog.Show("Error", "No elements selected. Please select elements to copy.");
            return Result.Failed;
        }

        // ─────────────────────────────────────────────────────────────
        // 0. Check if views or viewports are in the selection - if so, use them as targets
        // ─────────────────────────────────────────────────────────────
        List<View> viewsInSelection = new List<View>();
        List<ElementId> nonViewElementIds = new List<ElementId>();

        foreach (ElementId id in selectedElementIds)
        {
            Element elem = doc.GetElement(id);

            // Check if it's a View
            if (elem is View view &&
                !view.IsTemplate &&
                view.ViewType != ViewType.ProjectBrowser &&
                view.ViewType != ViewType.SystemBrowser &&
                view.ViewType != ViewType.Legend && // Legend views can't be target views
                view.Id != activeView.Id) // Exclude current view
            {
                viewsInSelection.Add(view);
            }
            // Check if it's a Viewport
            else if (elem is Viewport viewport)
            {
                View viewportView = doc.GetElement(viewport.ViewId) as View;
                if (viewportView != null &&
                    !viewportView.IsTemplate &&
                    viewportView.ViewType != ViewType.ProjectBrowser &&
                    viewportView.ViewType != ViewType.SystemBrowser &&
                    viewportView.ViewType != ViewType.Legend && // Legend views can't be target views
                    viewportView.Id != activeView.Id)
                {
                    viewsInSelection.Add(viewportView);
                }
                // Don't add viewport itself to elements to copy
            }
            else
            {
                nonViewElementIds.Add(id);
            }
        }

        List<View> selectedViews = null;
        ICollection<ElementId> elementsToCopy = null;

        if (viewsInSelection.Any())
        {
            // Use views from selection as targets
            selectedViews = viewsInSelection;

            // Copy non-view elements, or all elements if no non-view elements
            elementsToCopy = nonViewElementIds.Any() ? nonViewElementIds : selectedElementIds;
        }
        else
        {
            // No views in selection - show DataGrid to select target views
            elementsToCopy = selectedElementIds;
        }

        // ─────────────────────────────────────────────────────────────
        // 1. Collect every non-template, non-browser view (incl. sheets)
        //    (only if we need to show the DataGrid)
        // ─────────────────────────────────────────────────────────────
        if (selectedViews == null)
        {
            List<View> allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v =>
                       !v.IsTemplate &&
                       v.ViewType != ViewType.ProjectBrowser &&
                       v.ViewType != ViewType.SystemBrowser &&
                       v.ViewType != ViewType.Legend) // Legend views can't be target views
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

            if (!selectedRows.Any())
            {
                message = "No target views selected.";
                return Result.Failed;
            }

            // ─────────────────────────────────────────────────────────────
            // 5. Extract the View objects from the selected dictionaries
            // ─────────────────────────────────────────────────────────────
            selectedViews = CustomGUIs.ExtractOriginalObjects<View>(selectedRows);

            if (selectedViews == null || !selectedViews.Any())
            {
                message = "No valid target views selected.";
                return Result.Failed;
            }

        }

        // ─────────────────────────────────────────────────────────────
        // 6. Determine the actual source view for elements
        // ─────────────────────────────────────────────────────────────
        // If we're on a sheet, elements might belong to views placed on the sheet
        View sourceView = activeView;

        if (activeView is ViewSheet)
        {
            // Find which view the elements actually belong to
            ElementId ownerViewId = null;
            foreach (ElementId id in elementsToCopy)
            {
                Element elem = doc.GetElement(id);
                if (elem != null)
                {
                    ElementId ownerId = elem.OwnerViewId;

                    if (ownerId != null && ownerId != ElementId.InvalidElementId)
                    {
                        if (ownerViewId == null)
                        {
                            ownerViewId = ownerId;
                        }
                        else if (ownerViewId != ownerId)
                        {
                            // Elements belong to different views - use first found
                            break;
                        }
                    }
                }
            }

            if (ownerViewId != null && ownerViewId != ElementId.InvalidElementId)
            {
                View ownerView = doc.GetElement(ownerViewId) as View;
                if (ownerView != null)
                {
                    sourceView = ownerView;
                }
            }
        }

        // ─────────────────────────────────────────────────────────────
        // 7. Exclude source view from targets
        // ─────────────────────────────────────────────────────────────
        selectedViews = selectedViews.Where(v => v.Id != sourceView.Id).ToList();

        if (!selectedViews.Any())
        {
            message = "No valid target views selected (source view was excluded).";
            return Result.Failed;
        }

        // ─────────────────────────────────────────────────────────────
        // 8. Start a transaction to copy the elements
        // ─────────────────────────────────────────────────────────────
        using (Transaction transaction = new Transaction(doc, "Copy Selected Elements to Views"))
        {
            transaction.Start();
            CopyPasteOptions options = new CopyPasteOptions();

            foreach (View targetView in selectedViews)
            {
                try
                {
                    ElementTransformUtils.CopyElements(
                        sourceView,
                        elementsToCopy,
                        targetView,
                        Transform.Identity,
                        options);
                }
                catch (Exception ex)
                {
                    message = $"Error copying elements to view {targetView.Title}: {ex.Message}";
                    transaction.RollBack();
                    return Result.Failed;
                }
            }

            transaction.Commit();
        }

        return Result.Succeeded;
    }
}
