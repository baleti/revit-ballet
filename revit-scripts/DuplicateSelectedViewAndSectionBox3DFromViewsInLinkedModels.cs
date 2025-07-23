using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class DuplicateSelectedViewAndSectionBox3DFromViewsInLinkedModels : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Step 1: Get all 3D views in the project for selection
        List<View3D> all3DViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View3D))
            .Cast<View3D>()
            .Where(v => !v.IsTemplate)
            .ToList();

        if (all3DViews.Count == 0)
        {
            TaskDialog.Show("Error", "No 3D views found in the project.");
            return Result.Failed;
        }

        // Prepare data for the first DataGrid (3D view selection)
        var view3DEntries = all3DViews.Select(view => new Dictionary<string, object>
        {
            { "View Name", view.Name },
            { "View Id", view.Id.IntegerValue.ToString() }
        }).ToList();

        // Check if active view is a 3D view to use as initial selection
        View3D activeView3D = doc.ActiveView as View3D;
        List<int> initialSelectionIndices = null;
        if (activeView3D != null && !activeView3D.IsTemplate)
        {
            // Find the index of the active view in the entries list
            int activeViewIndex = view3DEntries.FindIndex(entry => 
                entry["View Id"].ToString() == activeView3D.Id.IntegerValue.ToString());
            
            if (activeViewIndex >= 0)
            {
                initialSelectionIndices = new List<int> { activeViewIndex };
            }
        }

        // Define column headers for 3D views
        var view3DPropertyNames = new List<string> { "View Name" };

        // Show first DataGrid for 3D view selection (single selection)
        List<Dictionary<string, object>> selectedView3DEntries = CustomGUIs.DataGrid(
            view3DEntries, 
            view3DPropertyNames, 
            false, // spanAllScreens parameter
            initialSelectionIndices); // initial selection indices

        if (selectedView3DEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No 3D view selected.");
            return Result.Cancelled;
        }

        // Get selected 3D view
        ElementId selected3DViewId = new ElementId(int.Parse(selectedView3DEntries[0]["View Id"].ToString()));
        View3D selected3DView = doc.GetElement(selected3DViewId) as View3D;

        // Step 2: Get all RevitLinkInstances in the project
        List<RevitLinkInstance> linkInstances = new FilteredElementCollector(doc)
            .OfClass(typeof(RevitLinkInstance))
            .Cast<RevitLinkInstance>()
            .Where(link => link.GetLinkDocument() != null)
            .ToList();

        if (linkInstances.Count == 0)
        {
            TaskDialog.Show("Info", "No loaded linked models found in the project.");
            return Result.Cancelled;
        }

        // Prepare data for the second DataGrid (linked models selection)
        var linkEntries = linkInstances.Select(link =>
        {
            Document linkedDocument = link.GetLinkDocument();
            RevitLinkType linkType = doc.GetElement(link.GetTypeId()) as RevitLinkType;
            string linkPath = "Unknown";
            string typeName = linkType?.Name ?? "Unknown Type";
            
            // Try to get external file reference path
            if (linkType != null)
            {
                try
                {
                    ExternalFileReference fileRef = linkType.GetExternalFileReference();
                    if (fileRef != null)
                    {
                        ModelPath modelPath = fileRef.GetPath();
                        if (modelPath != null)
                        {
                            linkPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                        }
                    }
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    // This link type doesn't have an external file reference (might be embedded)
                    linkPath = "Embedded/No External Path";
                }
            }
            
            return new Dictionary<string, object>
            {
                { "Link Name", link.Name },
                { "Type", typeName },
                { "Document Title", linkedDocument.Title },
                { "Path", linkPath },
                { "Link Id", link.Id.IntegerValue.ToString() }
            };
        }).ToList();

        // Define column headers for linked models
        var linkPropertyNames = new List<string> { "Type" };

        // Show second DataGrid for linked model selection (allow multiple selection)
        List<Dictionary<string, object>> selectedLinkEntries = CustomGUIs.DataGrid(linkEntries, linkPropertyNames, false);
        if (selectedLinkEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No linked models selected.");
            return Result.Cancelled;
        }

        // Get selected RevitLinkInstances
        List<RevitLinkInstance> selectedLinks = selectedLinkEntries
            .Select(entry => doc.GetElement(new ElementId(int.Parse(entry["Link Id"].ToString()))) as RevitLinkInstance)
            .Where(link => link != null)
            .ToList();

        // Step 3: Collect views from selected linked models
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
                    { "Link Name", linkInstance.Name },
                    { "Link Type", (doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType)?.Name ?? "Unknown Type" }
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
            "Link Type"
        };

        // Show third DataGrid for view selection (multiple selection allowed)
        List<Dictionary<string, object>> selectedViewEntries = CustomGUIs.DataGrid(allViewEntries, viewPropertyNames, false);
        if (selectedViewEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No views selected.");
            return Result.Cancelled;
        }

        // Step 4: Duplicate the 3D view for each selected linked view and set section boxes
        using (Transaction trans = new Transaction(doc, "Duplicate 3D View and Set Section Boxes"))
        {
            trans.Start();
            try
            {
                foreach (var viewEntry in selectedViewEntries)
                {
                    // Get the linked view information
                    ElementId linkedViewId = new ElementId(int.Parse(viewEntry["View Id"].ToString()));
                    ElementId linkInstanceId = new ElementId(int.Parse(viewEntry["Link Instance Id"].ToString()));
                    
                    RevitLinkInstance linkInstance = doc.GetElement(linkInstanceId) as RevitLinkInstance;
                    Document linkedDoc = linkInstance.GetLinkDocument();
                    View linkedView = linkedDoc.GetElement(linkedViewId) as View;

                    // Duplicate the 3D view (without detailing)
                    ElementId duplicatedViewId = selected3DView.Duplicate(ViewDuplicateOption.Duplicate);
                    View3D duplicatedView = doc.GetElement(duplicatedViewId) as View3D;

                    if (duplicatedView != null)
                    {
                        // Create a meaningful name that indicates the source
                        string linkTypeName = viewEntry["Link Type"].ToString();
                        string linkedViewName = viewEntry["View Name"].ToString();
                        string sheetNumber = viewEntry["Sheet Number"].ToString();
                        
                        // Build the name with sheet number if available
                        string nameFormat;
                        if (!string.IsNullOrWhiteSpace(sheetNumber))
                        {
                            nameFormat = $"{selected3DView.Name} - {linkTypeName} - {sheetNumber} - {linkedViewName}";
                        }
                        else
                        {
                            nameFormat = $"{selected3DView.Name} - {linkTypeName} - {linkedViewName}";
                        }
                        
                        // Sanitize names to remove prohibited characters
                        string sanitizedName = SanitizeViewName(nameFormat);
                        duplicatedView.Name = sanitizedName;

                        // Calculate and set the section box
                        BoundingBoxXYZ sectionBox = CalculateSectionBoxFromLinkedView(linkedView, linkInstance);
                        if (sectionBox != null)
                        {
                            duplicatedView.IsSectionBoxActive = true;
                            duplicatedView.SetSectionBox(sectionBox);
                        }
                    }
                }

                trans.Commit();
                
                TaskDialog.Show("Success", $"Successfully duplicated {selectedViewEntries.Count} 3D view(s) with section boxes from linked views.");
            }
            catch (Exception ex)
            {
                trans.RollBack();
                message = ex.Message;
                return Result.Failed;
            }
        }

        return Result.Succeeded;
    }

    private BoundingBoxXYZ CalculateSectionBoxFromLinkedView(View linkedView, RevitLinkInstance linkInstance)
    {
        // Get the transform from the link instance
        Transform linkTransform = linkInstance.GetTotalTransform();

        // Get the crop region of the linked view
        BoundingBoxXYZ cropBox = linkedView.CropBox;

        // Initialize min/max values
        double minX, minY, minZ, maxX, maxY, maxZ;

        // Check if the linked view is a plan view
        if (linkedView.ViewType == ViewType.FloorPlan ||
            linkedView.ViewType == ViewType.CeilingPlan ||
            linkedView.ViewType == ViewType.EngineeringPlan ||
            linkedView.ViewType == ViewType.AreaPlan)
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
            ViewPlan viewPlan = linkedView as ViewPlan;
            PlanViewRange viewRange = viewPlan.GetViewRange();

            // Get the levels associated with the view range
            Document linkedDoc = linkInstance.GetLinkDocument();
            ElementId topLevelId = viewRange.GetLevelId(PlanViewPlane.TopClipPlane);
            ElementId bottomLevelId = viewRange.GetLevelId(PlanViewPlane.BottomClipPlane);
            ElementId viewDepthLevelId = viewRange.GetLevelId(PlanViewPlane.ViewDepthPlane);

            Level topLevel = linkedDoc.GetElement(topLevelId) as Level;
            Level bottomLevel = linkedDoc.GetElement(bottomLevelId) as Level;
            Level viewDepthLevel = linkedDoc.GetElement(viewDepthLevelId) as Level;

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

        return sectionBox;
    }

    private string SanitizeViewName(string name)
    {
        // Remove prohibited characters from view names
        // Common prohibited characters include: \ / : * ? " < > | { } [ ] ` ~
        char[] prohibitedChars = { '\\', '/', ':', '*', '?', '"', '<', '>', '|', '{', '}', '[', ']', '`', '~' };
        
        string sanitized = name;
        foreach (char c in prohibitedChars)
        {
            sanitized = sanitized.Replace(c.ToString(), "");
        }
        
        // Also remove any leading or trailing whitespace
        sanitized = sanitized.Trim();
        
        // Ensure the name is not empty after sanitization
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            sanitized = "Duplicated View";
        }
        
        return sanitized;
    }
}
