using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SectionBox3DFromViews : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Ensure the active view is a 3D view.
        View3D current3DView = doc.ActiveView as View3D;
        if (current3DView == null)
        {
            TaskDialog.Show("Error", "The current view is not a 3D view.");
            return Result.Failed;
        }

        // Build a mapping from view Id to a list of sheets on which that view is placed.
        var viewports = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .Where(vp => vp.SheetId != ElementId.InvalidElementId)
                            .ToList();

        // Using a Dictionary to store one or more sheets per view.
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

        // Define view types to exclude.
        var excludedTypes = new HashSet<ViewType>
        {
            ViewType.ThreeD,
            ViewType.Schedule,
            ViewType.DrawingSheet,
            ViewType.Legend,
            ViewType.DraftingView,
            ViewType.SystemBrowser
        };

        // Get all non-template views from the document excluding unwanted view types.
        List<View> views = new FilteredElementCollector(doc)
                            .OfClass(typeof(View))
                            .Cast<View>()
                            .Where(v => !v.IsTemplate && !excludedTypes.Contains(v.ViewType))
                            .ToList();

        // Prepare data for the custom GUI.
        // For each view, if it is placed on a sheet (or sheets) via viewports,
        // get a comma‐separated list of sheet numbers, names, and sheet folders.
        var entries = views.Select(v =>
        {
            string sheetNumbers = "";
            string sheetNames = "";
            string sheetFolders = "";
            if (viewSheetMapping.ContainsKey(v.Id))
            {
                var sheets = viewSheetMapping[v.Id];
                sheetNumbers = string.Join(", ", sheets.Select(s => s.SheetNumber));
                sheetNames = string.Join(", ", sheets.Select(s => s.Name));
                sheetFolders = string.Join(", ", sheets.Select(s => s.LookupParameter("Sheet Folder")?.AsString() ?? ""));
            }
            return new Dictionary<string, object>
            {
                { "View Name", v.Name },
                { "View Type", v.ViewType.ToString() },
                { "Sheet Number", sheetNumbers },
                { "Sheet Name", sheetNames },
                { "Sheet Folder", sheetFolders },
                { "View Id", v.Id.IntegerValue.ToString() }
            };
        }).ToList();

        // Define the column headers for the custom GUI.
        var propertyNames = new List<string>
        {
            "View Name",
            "View Type",
            "Sheet Number",
            "Sheet Name",
            "Sheet Folder",
            "View Id"
        };

        // Show the custom GUI and let the user select one view.
        List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, false);
        if (selectedEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No view selected.");
            return Result.Cancelled;
        }

        // Retrieve the selected view.
        Dictionary<string, object> selectedEntry = selectedEntries.First();
        ElementId selectedViewId = new ElementId(int.Parse(selectedEntry["View Id"].ToString()));
        View selectedView = doc.GetElement(selectedViewId) as View;

        // Get the crop region of the selected view.
        BoundingBoxXYZ cropBox = selectedView.CropBox;

        // Initialize min/max values
        double minX, minY, minZ, maxX, maxY, maxZ;

        // Check if the selected view is a plan view
        if (selectedView.ViewType == ViewType.FloorPlan || 
            selectedView.ViewType == ViewType.CeilingPlan || 
            selectedView.ViewType == ViewType.EngineeringPlan ||
            selectedView.ViewType == ViewType.AreaPlan)
        {
            // For plan views, handle X and Y from crop box, but Z from view range
            
            // Get the X and Y extents from the crop box
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
            
            minX = worldPointsXY.Min(pt => pt.X);
            minY = worldPointsXY.Min(pt => pt.Y);
            maxX = worldPointsXY.Max(pt => pt.X);
            maxY = worldPointsXY.Max(pt => pt.Y);

            // Get the view range for Z extents
            ViewPlan viewPlan = selectedView as ViewPlan;
            PlanViewRange viewRange = viewPlan.GetViewRange();
            
            // Get the levels associated with the view range
            ElementId topLevelId = viewRange.GetLevelId(PlanViewPlane.TopClipPlane);
            ElementId bottomLevelId = viewRange.GetLevelId(PlanViewPlane.BottomClipPlane);
            ElementId viewDepthLevelId = viewRange.GetLevelId(PlanViewPlane.ViewDepthPlane);
            
            Level topLevel = doc.GetElement(topLevelId) as Level;
            Level bottomLevel = doc.GetElement(bottomLevelId) as Level;
            Level viewDepthLevel = doc.GetElement(viewDepthLevelId) as Level;
            
            // Get the offsets
            double topOffset = viewRange.GetOffset(PlanViewPlane.TopClipPlane);
            double bottomOffset = viewRange.GetOffset(PlanViewPlane.BottomClipPlane);
            double viewDepthOffset = viewRange.GetOffset(PlanViewPlane.ViewDepthPlane);
            
            // Calculate actual elevations
            double topElevation = topLevel.ProjectElevation + topOffset;
            double bottomElevation = bottomLevel.ProjectElevation + bottomOffset;
            double viewDepthElevation = viewDepthLevel.ProjectElevation + viewDepthOffset;
            
            // Use view depth as the bottom and top as the top
            minZ = viewDepthElevation;
            maxZ = topElevation;
        }
        else
        {
            // For non-plan views (sections, elevations), use the original logic
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

            minX = worldPoints.Min(pt => pt.X);
            minY = worldPoints.Min(pt => pt.Y);
            minZ = worldPoints.Min(pt => pt.Z);
            maxX = worldPoints.Max(pt => pt.X);
            maxY = worldPoints.Max(pt => pt.Y);
            maxZ = worldPoints.Max(pt => pt.Z);
        }

        BoundingBoxXYZ sectionBox = new BoundingBoxXYZ
        {
            Min = new XYZ(minX, minY, minZ),
            Max = new XYZ(maxX, maxY, maxZ),
            Transform = Transform.Identity // Defined in world coordinates.
        };

        // Apply the computed section box to the current 3D view.
        using (Transaction trans = new Transaction(doc, "Set 3D Section Box"))
        {
            trans.Start();
            current3DView.IsSectionBoxActive = true;
            current3DView.SetSectionBox(sectionBox);
            trans.Commit();
        }

        return Result.Succeeded;
    }
}
