using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[TransactionAttribute(TransactionMode.Manual)]
public class CopyCropRegionOfSelectedView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the current selection
        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();

        // Check if at least one element is selected
        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "Please select at least one view or viewport.");
            return Result.Failed;
        }

        // Get all selected views (resolve viewports to their views)
        List<View> selectedViews = new List<View>();
        foreach (ElementId selectedId in selectedIds)
        {
            Element selectedElement = doc.GetElement(selectedId);
            View view = null;

            if (selectedElement is Viewport viewport)
            {
                view = doc.GetElement(viewport.ViewId) as View;
            }
            else if (selectedElement is View v)
            {
                view = v;
            }

            if (view != null)
            {
                selectedViews.Add(view);
            }
        }

        // Check if we have valid views
        if (selectedViews.Count == 0)
        {
            TaskDialog.Show("Error", "No views or viewports found in selection.");
            return Result.Failed;
        }

        // Determine the source view to get crop region from
        View sourceView = null;

        if (selectedViews.Count == 1)
        {
            // Single selection - use it as source
            sourceView = selectedViews[0];
        }
        else
        {
            // Multiple selections - show DataGrid to let user pick the source
            sourceView = PromptUserForSourceView(doc, uidoc, selectedViews);

            if (sourceView == null)
            {
                // User cancelled
                return Result.Cancelled;
            }
        }

        // Check if the source view has a crop region
        if (sourceView == null || !sourceView.CropBoxActive)
        {
            message = "The selected view does not have an active crop region.";
            return Result.Failed;
        }

        // Get source view's crop shape
        ViewCropRegionShapeManager sourceCropManager = sourceView.GetCropRegionShapeManager();
        CurveLoop sourceCropShape = null;
        
        // Try to get the non-rectangular crop shape
        try
        {
            // Check if the view has a non-rectangular crop shape
            if (sourceCropManager.ShapeSet)
            {
                sourceCropShape = sourceCropManager.GetCropShape().FirstOrDefault();
            }
        }
        catch (Autodesk.Revit.Exceptions.InvalidOperationException)
        {
            // The view may not support non-rectangular crop shapes
        }
        
        // If no custom shape or not supported, create a curve loop from the crop box
        if (sourceCropShape == null)
        {
            BoundingBoxXYZ cropBox = sourceView.CropBox;
            
            XYZ pt1 = new XYZ(cropBox.Min.X, cropBox.Min.Y, 0);
            XYZ pt2 = new XYZ(cropBox.Max.X, cropBox.Min.Y, 0);
            XYZ pt3 = new XYZ(cropBox.Max.X, cropBox.Max.Y, 0);
            XYZ pt4 = new XYZ(cropBox.Min.X, cropBox.Max.Y, 0);
            
            CurveLoop cropBoxLoop = new CurveLoop();
            cropBoxLoop.Append(Line.CreateBound(pt1, pt2));
            cropBoxLoop.Append(Line.CreateBound(pt2, pt3));
            cropBoxLoop.Append(Line.CreateBound(pt3, pt4));
            cropBoxLoop.Append(Line.CreateBound(pt4, pt1));
            
            sourceCropShape = cropBoxLoop;
        }

        if (sourceCropShape == null)
        {
            message = "Could not retrieve crop shape from the selected view.";
            return Result.Failed;
        }

        // Get all OTHER selected views (excluding source) as target views
        List<View> targetViews = selectedViews.Where(v => v.Id != sourceView.Id).ToList();

        if (targetViews.Count == 0)
        {
            TaskDialog.Show("Info", "Only one view was selected. No target views to copy crop region to.");
            return Result.Cancelled;
        }

        // Apply the crop region to target views
        using (Transaction trans = new Transaction(doc, "Copy Crop Region to Selected Views"))
        {
            trans.Start();

            int successCount = 0;
            foreach (View targetView in targetViews)
            {
                try
                {
                    // Store the original crop visibility state
                    bool originalCropVisibility = targetView.CropBoxVisible;

                    // Activate crop box on target view
                    targetView.CropBoxActive = true;

                    // Try to apply the crop shape using ViewCropRegionShapeManager
                    bool shapeApplied = false;
                    try
                    {
                        ViewCropRegionShapeManager targetCropManager = targetView.GetCropRegionShapeManager();
                        targetCropManager.SetCropShape(sourceCropShape);
                        shapeApplied = true;
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        // View might not support non-rectangular crop shapes
                        shapeApplied = false;
                    }

                    // If SetCropShape failed or threw an exception, fall back to setting the CropBox
                    if (!shapeApplied)
                    {
                        BoundingBoxXYZ sourceBBox = GetBoundingBox(sourceCropShape);
                        targetView.CropBox = sourceBBox;
                    }

                    // Restore original crop visibility state
                    targetView.CropBoxVisible = originalCropVisibility;
                    successCount++;
                }
                catch (Exception ex)
                {
                    // Log exception but continue with other views
                    TaskDialog.Show("Error", $"Failed to apply crop region to view '{targetView.Title}': {ex.Message}");
                }
            }

            trans.Commit();
        }

        return Result.Succeeded;
    }

    /// <summary>
    /// Shows a DataGrid to let user select the source view from multiple selected views.
    /// Uses the same columns and sorting as SelectViews command.
    /// </summary>
    private View PromptUserForSourceView(Document doc, UIDocument uidoc, List<View> views)
    {
        // Get the currently active view
        View activeView = doc.ActiveView;

        // Create a mapping for views that are placed on sheets
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

        // Get browser organization columns (pass views so it can detect sheets vs views)
        List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
            BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, views);

        // Prepare data for the data grid
        List<Dictionary<string, object>> viewData = new List<Dictionary<string, object>>();

        foreach (View view in views)
        {
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

            viewInfo["__OriginalObject"] = view; // Store original object for retrieval

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

        // Determine initial selection based on active view
        List<int> initialSelection = new List<int>();

        if (activeView != null)
        {
            if (activeView is ViewSheet activeSheet)
            {
                // Active view is a sheet - select all views in the list that are placed on this sheet
                for (int i = 0; i < viewData.Count; i++)
                {
                    if (viewData[i].ContainsKey("__OriginalObject") && viewData[i]["__OriginalObject"] is View v)
                    {
                        // Check if this view is placed on the active sheet
                        if (viewToSheetMap.TryGetValue(v.Id, out ViewSheet sheet) && sheet.Id == activeSheet.Id)
                        {
                            initialSelection.Add(i);
                        }
                    }
                }
            }
            else
            {
                // Active view is not a sheet - find and select it if it's in the list
                for (int i = 0; i < viewData.Count; i++)
                {
                    if (viewData[i].ContainsKey("__OriginalObject") && viewData[i]["__OriginalObject"] is View v)
                    {
                        if (v.Id == activeView.Id)
                        {
                            initialSelection.Add(i);
                            break;
                        }
                    }
                }
            }
        }

        // Define the column headers - browser columns first, then standard columns.
        List<string> columns = new List<string>();
        columns.AddRange(browserColumns.Select(bc => bc.Name));
        columns.Add("SheetNumber");
        columns.Add("Name");
        columns.Add("Sheet");

        // Set UIDocument for selection set support
        CustomGUIs.SetCurrentUIDocument(uidoc);

        // Show the selection dialog - user must select exactly ONE view as source
        List<Dictionary<string, object>> selectedViewData = CustomGUIs.DataGrid(
            viewData,
            columns,
            false,  // Don't span all screens
            initialSelection.Count > 0 ? initialSelection : null  // Pass initial selection
        );

        // Return the selected view (only one should be selected)
        if (selectedViewData != null && selectedViewData.Count > 0)
        {
            // Get the first selected view (use __OriginalObject)
            if (selectedViewData[0].ContainsKey("__OriginalObject") &&
                selectedViewData[0]["__OriginalObject"] is View selectedView)
            {
                return selectedView;
            }
        }

        return null; // User cancelled or no valid selection
    }

    // Helper method to check if a view can be cropped
    private bool CanViewBeCropped(View view)
    {
        // These view types typically support crop regions
        ViewType[] croppableViewTypes = new ViewType[]
        {
            ViewType.FloorPlan,
            ViewType.CeilingPlan,
            ViewType.Elevation,
            ViewType.ThreeD,
            ViewType.Section,
            ViewType.Detail,
            ViewType.AreaPlan,
            ViewType.EngineeringPlan
        };

        return croppableViewTypes.Contains(view.ViewType);
    }

    // Helper method to get a BoundingBoxXYZ from a CurveLoop
    private BoundingBoxXYZ GetBoundingBox(CurveLoop curveLoop)
    {
        double minX = double.MaxValue;
        double minY = double.MaxValue;
        double maxX = double.MinValue;
        double maxY = double.MinValue;

        foreach (Curve curve in curveLoop)
        {
            XYZ p1 = curve.GetEndPoint(0);
            XYZ p2 = curve.GetEndPoint(1);

            minX = Math.Min(minX, Math.Min(p1.X, p2.X));
            minY = Math.Min(minY, Math.Min(p1.Y, p2.Y));
            maxX = Math.Max(maxX, Math.Max(p1.X, p2.X));
            maxY = Math.Max(maxY, Math.Max(p1.Y, p2.Y));
        }

        BoundingBoxXYZ bbox = new BoundingBoxXYZ();
        bbox.Min = new XYZ(minX, minY, 0);
        bbox.Max = new XYZ(maxX, maxY, 0);
        return bbox;
    }
}
