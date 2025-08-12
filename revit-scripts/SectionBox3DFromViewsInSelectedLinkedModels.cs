using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SectionBox3DFromViewsInSelectedLinkedModels : IExternalCommand
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

        // Get selected elements
        var selectedIds = uiDoc.GetSelectionIds();
        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "Please select one or more linked models.");
            return Result.Failed;
        }

        // Filter for RevitLinkInstances
        List<RevitLinkInstance> selectedLinks = selectedIds
            .Select(id => doc.GetElement(id))
            .OfType<RevitLinkInstance>()
            .Where(link => link.GetLinkDocument() != null)
            .ToList();

        if (selectedLinks.Count == 0)
        {
            TaskDialog.Show("Error", "No valid linked models selected. Please select RevitLinkInstance elements.");
            return Result.Failed;
        }

        // Define view types to exclude
        var excludedTypes = new HashSet<ViewType>
        {
            ViewType.ThreeD,
            ViewType.Schedule,
            ViewType.DrawingSheet,
            ViewType.Legend,
            ViewType.DraftingView,
            ViewType.SystemBrowser
        };

        // Collect views from all selected linked models
        var allViewEntries = new List<Dictionary<string, object>>();
        
        foreach (var linkInstance in selectedLinks)
        {
            Document linkDoc = linkInstance.GetLinkDocument();
            if (linkDoc == null) continue;

            // Get viewports to build view-sheet mapping
            var viewports = new FilteredElementCollector(linkDoc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .Where(vp => vp.SheetId != ElementId.InvalidElementId)
                .ToList();

            var viewSheetMapping = new Dictionary<ElementId, List<ViewSheet>>();
            foreach (var vp in viewports)
            {
                ElementId viewId = vp.ViewId;
                ViewSheet sheet = linkDoc.GetElement(vp.SheetId) as ViewSheet;
                if (sheet != null)
                {
                    if (!viewSheetMapping.ContainsKey(viewId))
                    {
                        viewSheetMapping[viewId] = new List<ViewSheet>();
                    }
                    viewSheetMapping[viewId].Add(sheet);
                }
            }

            // Get all non-template views from the linked document
            List<View> views = new FilteredElementCollector(linkDoc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && !excludedTypes.Contains(v.ViewType) && v.CropBoxActive)
                .ToList();

            // Create entries for each view
            foreach (var view in views)
            {
                string sheetNumbers = "";
                string sheetNames = "";
                string sheetFolders = "";
                if (viewSheetMapping.ContainsKey(view.Id))
                {
                    var sheets = viewSheetMapping[view.Id];
                    sheetNumbers = string.Join(", ", sheets.Select(s => s.SheetNumber));
                    sheetNames = string.Join(", ", sheets.Select(s => s.Name));
                    sheetFolders = string.Join(", ", sheets.Select(s => s.LookupParameter("Sheet Folder")?.AsString() ?? ""));
                }

                allViewEntries.Add(new Dictionary<string, object>
                {
                    { "View Name", view.Name },
                    { "View Type", view.ViewType.ToString() },
                    { "Sheet Number", sheetNumbers },
                    { "Sheet Name", sheetNames },
                    { "Sheet Folder", sheetFolders },
                    { "View Id", view.Id.IntegerValue.ToString() },
                    { "Link Instance Id", linkInstance.Id.IntegerValue.ToString() },
                    { "Link Name", linkInstance.Name }
                });
            }
        }

        if (allViewEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No views with crop regions found in the selected linked models.");
            return Result.Cancelled;
        }

        // Define column headers for views
        var viewPropertyNames = new List<string>
        {
            "View Name",
            "View Type",
            "Sheet Number",
            "Sheet Name",
            "Sheet Folder",
            "View Id",
            "Link Name"
        };

        // Show DataGrid for view selection
        List<Dictionary<string, object>> selectedViewEntries = CustomGUIs.DataGrid(allViewEntries, viewPropertyNames, false);
        if (selectedViewEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No view selected.");
            return Result.Cancelled;
        }

        // Get the selected view and link instance
        Dictionary<string, object> selectedEntry = selectedViewEntries.First();
        ElementId selectedViewId = new ElementId(int.Parse(selectedEntry["View Id"].ToString()));
        ElementId linkInstanceId = new ElementId(int.Parse(selectedEntry["Link Instance Id"].ToString()));
        
        RevitLinkInstance selectedLinkInstance = doc.GetElement(linkInstanceId) as RevitLinkInstance;
        Document selectedLinkDoc = selectedLinkInstance.GetLinkDocument();
        View selectedView = selectedLinkDoc.GetElement(selectedViewId) as View;

        // Get the transform from the link instance
        Transform linkTransform = selectedLinkInstance.GetTotalTransform();

        // Get the crop region of the selected view
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
                    XYZ viewWorldPt = cropBox.Transform.OfPoint(localPt);
                    // Transform from link coordinates to host coordinates
                    XYZ hostWorldPt = linkTransform.OfPoint(viewWorldPt);
                    worldPointsXY.Add(hostWorldPt);
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

            Level topLevel = selectedLinkDoc.GetElement(topLevelId) as Level;
            Level bottomLevel = selectedLinkDoc.GetElement(bottomLevelId) as Level;
            Level viewDepthLevel = selectedLinkDoc.GetElement(viewDepthLevelId) as Level;

            // Get the offsets
            double topOffset = viewRange.GetOffset(PlanViewPlane.TopClipPlane);
            double bottomOffset = viewRange.GetOffset(PlanViewPlane.BottomClipPlane);
            double viewDepthOffset = viewRange.GetOffset(PlanViewPlane.ViewDepthPlane);

            // Calculate actual elevations in link coordinates
            double topElevation = topLevel.ProjectElevation + topOffset;
            double bottomElevation = bottomLevel.ProjectElevation + bottomOffset;
            double viewDepthElevation = viewDepthLevel.ProjectElevation + viewDepthOffset;

            // Transform Z coordinates to host coordinates
            XYZ bottomPoint = linkTransform.OfPoint(new XYZ(0, 0, viewDepthElevation));
            XYZ topPoint = linkTransform.OfPoint(new XYZ(0, 0, topElevation));

            minZ = bottomPoint.Z;
            maxZ = topPoint.Z;
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
                        XYZ viewWorldPt = cropBox.Transform.OfPoint(localPt);
                        // Transform from link coordinates to host coordinates
                        XYZ hostWorldPt = linkTransform.OfPoint(viewWorldPt);
                        worldPoints.Add(hostWorldPt);
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
            Transform = Transform.Identity // Defined in world coordinates
        };

        // Apply the computed section box to the current 3D view
        using (Transaction trans = new Transaction(doc, "Set 3D Section Box from Selected Linked View"))
        {
            trans.Start();
            current3DView.IsSectionBoxActive = true;
            current3DView.SetSectionBox(sectionBox);
            trans.Commit();
        }

        // Zoom to fit the section box
        UIView uiView = uiDoc
                        .GetOpenUIViews()
                        .FirstOrDefault(v => v.ViewId == current3DView.Id);

        if (uiView != null)
        {
            uiView.ZoomToFit();
        }

        return Result.Succeeded;
    }
}
