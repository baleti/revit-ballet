using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;

[TransactionAttribute(TransactionMode.Manual)]
public class CopyCropRegionOfSelectedView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the current selection
        ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

        // Check if exactly one element is selected
        if (selectedIds.Count != 1)
        {
            message = "Please select exactly one view or viewport to copy the crop region from.";
            return Result.Failed;
        }

        // Get the selected element
        ElementId selectedId = selectedIds.First();
        Element selectedElement = doc.GetElement(selectedId);

        // Determine the source view to get crop region from
        View sourceView = null;

        if (selectedElement is Viewport)
        {
            // If a viewport is selected, get its view
            Viewport viewport = selectedElement as Viewport;
            sourceView = doc.GetElement(viewport.ViewId) as View;
        }
        else if (selectedElement is View)
        {
            // If a view is directly selected
            sourceView = selectedElement as View;
        }
        else
        {
            message = "The selected element is not a view or viewport.";
            return Result.Failed;
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

        // Get all views in the project
        FilteredElementCollector viewCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(View));

        // Prepare data for the data grid and map view titles to view objects
        List<Dictionary<string, object>> viewData = new List<Dictionary<string, object>>();
        Dictionary<string, View> titleToViewMap = new Dictionary<string, View>();

        foreach (View view in viewCollector.Cast<View>())
        {
            // Skip view templates, legends, schedules, project browser, system browser,
            // views that cannot be cropped, and the source view
            if (view.IsTemplate || 
                view.ViewType == ViewType.Legend || 
                view.ViewType == ViewType.Schedule ||
                view.ViewType == ViewType.ProjectBrowser ||
                view.ViewType == ViewType.SystemBrowser ||
                view.Id.IntegerValue == sourceView.Id.IntegerValue) // Skip the source view
                continue;

            // Skip views that don't support crop box (additional check)
            if (!CanViewBeCropped(view))
                continue;

            string sheetInfo = string.Empty;
            if (view is ViewSheet)
            {
                // For a sheet, show its own sheet number and name
                ViewSheet viewSheet = view as ViewSheet;
                sheetInfo = $"{viewSheet.SheetNumber} - {viewSheet.Name}";
            }
            else if (viewToSheetMap.TryGetValue(view.Id, out ViewSheet sheet))
            {
                // For non-sheet views placed on a sheet, display the sheet info
                sheetInfo = $"{sheet.SheetNumber} - {sheet.Name}";
            }
            else
            {
                sheetInfo = "Not Placed";
            }

            // Add view info to the data collection
            titleToViewMap[view.Title] = view;
            Dictionary<string, object> viewInfo = new Dictionary<string, object>
            {
                { "Title", view.Title },
                { "View Type", view.ViewType.ToString() },
                { "Sheet", sheetInfo }
            };
            viewData.Add(viewInfo);
        }

        // Define the column headers
        List<string> columns = new List<string> { "Title", "View Type", "Sheet" };

        // Show the selection dialog using CustomGUIs.DataGrid
        List<Dictionary<string, object>> selectedViews = CustomGUIs.DataGrid(
            viewData,
            columns,
            false  // Don't span all screens
        );

        // If the user made a selection, apply the crop region to those views
        if (selectedViews != null && selectedViews.Any())
        {
            using (Transaction trans = new Transaction(doc, "Copy Crop Region to Selected Views"))
            {
                trans.Start();

                int successCount = 0;
                foreach (Dictionary<string, object> selectedViewData in selectedViews)
                {
                    string viewTitle = selectedViewData["Title"].ToString();
                    if (titleToViewMap.TryGetValue(viewTitle, out View targetView))
                    {
                        try
                        {
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

                            // Make crop region visible
                            targetView.CropBoxVisible = true;
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            // Log exception but continue with other views
                            TaskDialog.Show("Error", $"Failed to apply crop region to view '{viewTitle}': {ex.Message}");
                        }
                    }
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }

        return Result.Cancelled;
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
