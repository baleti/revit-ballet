using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System;

[Transaction(TransactionMode.Manual)]
public class ViewsContainingElement : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the currently selected element
        var selection = uiDoc.Selection.GetElementIds();
        if (selection.Count == 0)
        {
            TaskDialog.Show("Error", "Please select an element first.");
            return Result.Failed;
        }

        ElementId selectedElementId = selection.First();
        Element selectedElement = doc.GetElement(selectedElementId);
        
        // Get the bounding box of the selected element
        BoundingBoxXYZ elementBB = selectedElement.get_BoundingBox(null);
        if (elementBB == null)
        {
            TaskDialog.Show("Error", "Selected element has no bounding box.");
            return Result.Failed;
        }

        // Define view types to exclude
        var excludedTypes = new HashSet<ViewType>
        {
            ViewType.Schedule,
            ViewType.DrawingSheet,
            ViewType.Legend,
            ViewType.DraftingView,
            ViewType.SystemBrowser
        };

        // Get all non-template views from the document
        List<View> views = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => !v.IsTemplate && !excludedTypes.Contains(v.ViewType))
            .ToList();

        // Build view to sheet mapping
        var viewports = new FilteredElementCollector(doc)
            .OfClass(typeof(Viewport))
            .Cast<Viewport>()
            .Where(vp => vp.SheetId != ElementId.InvalidElementId)
            .ToList();

        var viewSheetMapping = new Dictionary<ElementId, List<ViewSheet>>();
        foreach (var vp in viewports)
        {
            ElementId viewId = vp.ViewId;
            ViewSheet sheet = doc.GetElement(vp.SheetId) as ViewSheet;
            if (sheet != null)
            {
                if (!viewSheetMapping.ContainsKey(viewId))
                {
                    viewSheetMapping[viewId] = new List<ViewSheet>();
                }
                viewSheetMapping[viewId].Add(sheet);
            }
        }

        // Check which views contain the element
        var viewsContainingElement = new List<Dictionary<string, object>>();
        
        foreach (View view in views)
        {
            if (IsElementInViewRange(view, elementBB, doc))
            {
                string sheetNumbers = "";
                string sheetNames = "";
                if (viewSheetMapping.ContainsKey(view.Id))
                {
                    var sheets = viewSheetMapping[view.Id];
                    sheetNumbers = string.Join(", ", sheets.Select(s => s.SheetNumber));
                    sheetNames = string.Join(", ", sheets.Select(s => s.Name));
                }

                viewsContainingElement.Add(new Dictionary<string, object>
                {
                    { "View Name", view.Name },
                    { "View Type", view.ViewType.ToString() },
                    { "Sheet Number", sheetNumbers },
                    { "Sheet Name", sheetNames },
                    { "View Id", view.Id.Value.ToString() }
                });
            }
        }

        // Display results
        if (viewsContainingElement.Count == 0)
        {
            TaskDialog.Show("Result", $"No views contain the selected element: {selectedElement.Name}");
            return Result.Succeeded;
        }

        // Define column headers
        var propertyNames = new List<string>
        {
            "View Name",
            "View Type",
            "Sheet Number",
            "Sheet Name",
            "View Id"
        };

        // Show the results in a data grid
        TaskDialog.Show("Info", $"Found {viewsContainingElement.Count} views containing element: {selectedElement.Name}");
        CustomGUIs.DataGrid(viewsContainingElement, propertyNames, false);

        return Result.Succeeded;
    }

    private bool IsElementInViewRange(View view, BoundingBoxXYZ elementBB, Document doc)
    {
        try
        {
            // Handle 3D views
            if (view is View3D view3D)
            {
                if (!view3D.IsSectionBoxActive)
                {
                    // If no section box is active, the view shows everything
                    return true;
                }

                BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
                return BoundingBoxesIntersect(elementBB, sectionBox);
            }

            // Get crop box
            if (!view.CropBoxActive)
            {
                // If crop box is not active, assume the view shows everything
                return true;
            }

            BoundingBoxXYZ cropBox = view.CropBox;

            // Handle plan views
            if (view.ViewType == ViewType.FloorPlan ||
                view.ViewType == ViewType.CeilingPlan ||
                view.ViewType == ViewType.EngineeringPlan ||
                view.ViewType == ViewType.AreaPlan)
            {
                ViewPlan viewPlan = view as ViewPlan;
                PlanViewRange viewRange = viewPlan.GetViewRange();

                // Get Z extents from view range
                ElementId topLevelId = viewRange.GetLevelId(PlanViewPlane.TopClipPlane);
                ElementId bottomLevelId = viewRange.GetLevelId(PlanViewPlane.BottomClipPlane);
                ElementId viewDepthLevelId = viewRange.GetLevelId(PlanViewPlane.ViewDepthPlane);

                Level topLevel = doc.GetElement(topLevelId) as Level;
                Level bottomLevel = doc.GetElement(bottomLevelId) as Level;
                Level viewDepthLevel = doc.GetElement(viewDepthLevelId) as Level;

                double topOffset = viewRange.GetOffset(PlanViewPlane.TopClipPlane);
                double bottomOffset = viewRange.GetOffset(PlanViewPlane.BottomClipPlane);
                double viewDepthOffset = viewRange.GetOffset(PlanViewPlane.ViewDepthPlane);

                double topElevation = topLevel.ProjectElevation + topOffset;
                double bottomElevation = bottomLevel.ProjectElevation + bottomOffset;
                double viewDepthElevation = viewDepthLevel.ProjectElevation + viewDepthOffset;

                // Check if element's Z is within view range
                if (elementBB.Max.Z < viewDepthElevation || elementBB.Min.Z > topElevation)
                {
                    return false;
                }

                // Transform crop box corners to world coordinates for X,Y check
                var worldPointsXY = new List<XYZ>();
                foreach (double x in new double[] { cropBox.Min.X, cropBox.Max.X })
                {
                    foreach (double y in new double[] { cropBox.Min.Y, cropBox.Max.Y })
                    {
                        XYZ localPt = new XYZ(x, y, 0);
                        XYZ worldPt = cropBox.Transform.OfPoint(localPt);
                        worldPointsXY.Add(worldPt);
                    }
                }

                double minX = worldPointsXY.Min(pt => pt.X);
                double minY = worldPointsXY.Min(pt => pt.Y);
                double maxX = worldPointsXY.Max(pt => pt.X);
                double maxY = worldPointsXY.Max(pt => pt.Y);

                // Check if element's X,Y is within crop box
                return !(elementBB.Max.X < minX || elementBB.Min.X > maxX ||
                        elementBB.Max.Y < minY || elementBB.Min.Y > maxY);
            }
            else
            {
                // For sections, elevations, and other view types
                // Transform crop box to world coordinates
                var worldPoints = new List<XYZ>();
                foreach (double x in new double[] { cropBox.Min.X, cropBox.Max.X })
                {
                    foreach (double y in new double[] { cropBox.Min.Y, cropBox.Max.Y })
                    {
                        foreach (double z in new double[] { cropBox.Min.Z, cropBox.Max.Z })
                        {
                            XYZ localPt = new XYZ(x, y, z);
                            XYZ worldPt = cropBox.Transform.OfPoint(localPt);
                            worldPoints.Add(worldPt);
                        }
                    }
                }

                double minX = worldPoints.Min(pt => pt.X);
                double minY = worldPoints.Min(pt => pt.Y);
                double minZ = worldPoints.Min(pt => pt.Z);
                double maxX = worldPoints.Max(pt => pt.X);
                double maxY = worldPoints.Max(pt => pt.Y);
                double maxZ = worldPoints.Max(pt => pt.Z);

                BoundingBoxXYZ worldCropBox = new BoundingBoxXYZ
                {
                    Min = new XYZ(minX, minY, minZ),
                    Max = new XYZ(maxX, maxY, maxZ)
                };

                return BoundingBoxesIntersect(elementBB, worldCropBox);
            }
        }
        catch
        {
            // If any error occurs, assume the element might be visible
            return true;
        }
    }

    private bool BoundingBoxesIntersect(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        // Check if bounding boxes intersect
        return !(bb1.Max.X < bb2.Min.X || bb1.Min.X > bb2.Max.X ||
                bb1.Max.Y < bb2.Min.Y || bb1.Min.Y > bb2.Max.Y ||
                bb1.Max.Z < bb2.Min.Z || bb1.Min.Z > bb2.Max.Z);
    }
}
