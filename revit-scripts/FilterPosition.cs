using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.ReadOnly)]
public class FilterPosition : IExternalCommand
{
    // Helper class to store element data along with its reference
    private class ElementDataWithReference
    {
        public Dictionary<string, object> Data { get; set; }
        public Reference Reference { get; set; }
        public Element Element { get; set; }
        public bool IsLinked { get; set; }
        public string DocumentName { get; set; }
        public RevitLinkInstance LinkInstance { get; set; }
        public Transform LinkTransform { get; set; }
        public List<string> ContainingGroups { get; set; }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the active UIDocument and Document.
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Retrieve the current selection using the SelectionModeManager which supports linked references
        IList<Reference> selectedRefs = uidoc.GetReferences();
        if (selectedRefs == null || !selectedRefs.Any())
        {
            // Try to get regular selection IDs as fallback
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            if (selectedIds == null || !selectedIds.Any())
            {
                TaskDialog.Show("Warning", "Please select elements before running the command.");
                return Result.Failed;
            }
            
            // Convert ElementIds to References for consistency
            selectedRefs = selectedIds.Select(id => new Reference(doc.GetElement(id))).ToList();
        }

        // Get all grid lines in the host document for grid proximity calculations
        List<Grid> allGrids = new FilteredElementCollector(doc)
            .OfClass(typeof(Grid))
            .Cast<Grid>()
            .ToList();

        // Get all scope boxes in the host document
        List<Element> allScopeBoxes = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
            .ToList();

        // Get all model groups in the host document
        List<Group> allModelGroups = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_IOSModelGroups)
            .WhereElementIsNotElementType()
            .Cast<Group>()
            .ToList();

        // Process references to get elements (both from current doc and linked docs)
        List<ElementDataWithReference> elementDataList = new List<ElementDataWithReference>();

        foreach (Reference reference in selectedRefs)
        {
            Element elem = null;
            bool isLinked = false;
            string documentName = doc.Title;
            RevitLinkInstance linkInstance = null;
            Transform linkTransform = Transform.Identity;

            try
            {
                if (reference.LinkedElementId != ElementId.InvalidElementId)
                {
                    // This is a linked element
                    isLinked = true;
                    linkInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkInstance != null)
                    {
                        Document linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            elem = linkedDoc.GetElement(reference.LinkedElementId);
                            documentName = linkedDoc.Title;
                            linkTransform = linkInstance.GetTotalTransform();
                        }
                    }
                }
                else
                {
                    // Regular element in current document
                    elem = doc.GetElement(reference.ElementId);
                }

                if (elem != null)
                {
                    elementDataList.Add(new ElementDataWithReference
                    {
                        Element = elem,
                        Reference = reference,
                        IsLinked = isLinked,
                        DocumentName = documentName,
                        LinkInstance = linkInstance,
                        LinkTransform = linkTransform,
                        ContainingGroups = new List<string>()
                    });
                }
            }
            catch
            {
                // Skip problematic references
                continue;
            }
        }

        if (!elementDataList.Any())
        {
            TaskDialog.Show("Warning", "No valid elements found in selection.");
            return Result.Failed;
        }

        // Define the property names (columns) for the data grid.
        List<string> propertyNames = new List<string>
        {
            "Element Id",
            "Document",
            "Category",
            "Name",
            "LocationType",
            "X",
            "Y",
            "Z",
            "Rotation",
            "FacingFlipped",
            "HandFlipped",
            "Curve Start",
            "Curve End",
            "ScopeBox",
            "ScopeBox Center",
            "Grid X Up",
            "Grid X Down",
            "Grid Y Up",
            "Grid Y Down"
        };

        // Prepare the list of dictionaries that will hold each element's orientation data.
        List<Dictionary<string, object>> gridData = new List<Dictionary<string, object>>();

        foreach (var elemData in elementDataList)
        {
            Element elem = elemData.Element;
            Dictionary<string, object> properties = new Dictionary<string, object>();

            // Basic element properties.
            properties["Element Id"] = elem.Id.IntegerValue;
            properties["Document"] = elemData.DocumentName + (elemData.IsLinked ? " (Linked)" : "");
            properties["Category"] = elem.Category != null ? elem.Category.Name : "";
            properties["Name"] = elem.Name;

            // Initialize orientation fields.
            properties["LocationType"] = "";
            properties["X"] = "";
            properties["Y"] = "";
            properties["Z"] = "";
            properties["Rotation"] = "";
            properties["FacingFlipped"] = "";
            properties["HandFlipped"] = "";
            properties["Curve Start"] = "";
            properties["Curve End"] = "";
            properties["ScopeBox"] = "";
            properties["ScopeBox Center"] = "";
            properties["Grid X Up"] = "";
            properties["Grid X Down"] = "";
            properties["Grid Y Up"] = "";
            properties["Grid Y Down"] = "";

            XYZ elementLocation = null;

            // Check if the element has a location.
            Location loc = elem.Location;
            if (loc != null)
            {
                // For elements with a point location.
                if (loc is LocationPoint locPoint)
                {
                    properties["LocationType"] = "Point";
                    XYZ point = locPoint.Point;
                    
                    // Transform the point if it's from a linked model
                    if (elemData.IsLinked)
                    {
                        point = elemData.LinkTransform.OfPoint(point);
                    }
                    
                    elementLocation = point;
                    properties["X"] = point.X;
                    properties["Y"] = point.Y;
                    properties["Z"] = point.Z;
                    properties["Rotation"] = locPoint.Rotation;
                }
                // For elements with a curve location.
                else if (loc is LocationCurve locCurve)
                {
                    properties["LocationType"] = "Curve";
                    Curve curve = locCurve.Curve;
                    if (curve != null)
                    {
                        XYZ start = curve.GetEndPoint(0);
                        XYZ end = curve.GetEndPoint(1);
                        
                        // Transform the points if from a linked model
                        if (elemData.IsLinked)
                        {
                            start = elemData.LinkTransform.OfPoint(start);
                            end = elemData.LinkTransform.OfPoint(end);
                        }
                        
                        properties["Curve Start"] = $"({start.X:F2}, {start.Y:F2}, {start.Z:F2})";
                        properties["Curve End"] = $"({end.X:F2}, {end.Y:F2}, {end.Z:F2})";
                        
                        // Use midpoint as element location for grid calculations
                        elementLocation = (start + end) / 2;
                    }
                }
            }

            // If the element is a FamilyInstance, record its flip properties.
            if (elem is FamilyInstance familyInstance)
            {
                properties["FacingFlipped"] = familyInstance.FacingFlipped;
                properties["HandFlipped"] = familyInstance.HandFlipped;
            }

            // If the element is a View, and it has an active CropBox,
            // use the CropBox center as the location.
            if (elem is View view)
            {
                if (view.CropBoxActive && view.CropBox != null)
                {
                    XYZ center = (view.CropBox.Min + view.CropBox.Max) / 2;
                    
                    // Transform if from linked model
                    if (elemData.IsLinked)
                    {
                        center = elemData.LinkTransform.OfPoint(center);
                    }
                    
                    elementLocation = center;
                    properties["LocationType"] = "View CropBox Center";
                    properties["X"] = center.X;
                    properties["Y"] = center.Y;
                    properties["Z"] = center.Z;
                }
            }

            // Compute element bounding box in host coordinates
            BoundingBoxXYZ elemBB = GetTransformedBoundingBox(elem.get_BoundingBox(null), elemData.LinkTransform);
            if (elemBB == null && elem is View view2 && view2.CropBoxActive)
            {
                elemBB = GetTransformedBoundingBox(view2.CropBox, elemData.LinkTransform);
            }

            // Check for model groups whose bounding boxes contain the element's bounding box
            if (elemBB != null)
            {
                List<string> containingGroups = allModelGroups
                    .Select(g => new { Group = g, BB = g.get_BoundingBox(null) })
                    .Where(x => x.BB != null && Contains(x.BB, elemBB))
                    .Select(x => x.Group.Name)
                    .ToList();

                elemData.ContainingGroups = containingGroups;
            }

            // Check for scope box containment
            if (elementLocation != null)
            {
                List<string> scopeNames = new List<string>();

                foreach (Element scopeBox in allScopeBoxes)
                {
                    BoundingBoxXYZ bb = scopeBox.get_BoundingBox(null);
                    if (bb != null &&
                        elementLocation.X >= bb.Min.X && elementLocation.X <= bb.Max.X &&
                        elementLocation.Y >= bb.Min.Y && elementLocation.Y <= bb.Max.Y &&
                        elementLocation.Z >= bb.Min.Z && elementLocation.Z <= bb.Max.Z)
                    {
                        scopeNames.Add(scopeBox.Name);
                    }
                }

                properties["ScopeBox"] = string.Join(", ", scopeNames);
                properties["ScopeBox Center"] = string.Join(", ", scopeNames);

                // Find nearest grid lines
                FindNearestGrids(elementLocation, allGrids, properties);
            }

            // Store the data along with the index for later reference
            elemData.Data = properties;
            gridData.Add(properties);
        }

        // Determine the maximum number of containing groups
        int maxGroups = elementDataList.Any() ? elementDataList.Max(ed => ed.ContainingGroups?.Count ?? 0) : 0;

        // Insert group columns before ScopeBox
        int scopeBoxIndex = propertyNames.IndexOf("ScopeBox");

        for (int i = 1; i <= maxGroups; i++)
        {
            propertyNames.Insert(scopeBoxIndex + (i - 1), $"Group {i}");
        }

        // Populate the group columns in each properties dictionary
        foreach (var elemData in elementDataList)
        {
            var properties = elemData.Data;
            var groups = elemData.ContainingGroups ?? new List<string>();
            for (int i = 1; i <= maxGroups; i++)
            {
                properties[$"Group {i}"] = (i - 1 < groups.Count) ? groups[i - 1] : "";
            }
        }

        // Display the data grid.
        List<Dictionary<string, object>> selectedFromGrid = CustomGUIs.DataGrid(gridData, propertyNames, false);

        // If the user made a selection in the grid, update the active selection.
        if (selectedFromGrid?.Any() == true)
        {
            // Build a new list of references based on the selected items
            List<Reference> finalReferences = new List<Reference>();

            foreach (var selectedData in selectedFromGrid)
            {
                // Find the matching element data by comparing element ID and document name
                var matchingElemData = elementDataList.FirstOrDefault(ed =>
                    (int)ed.Data["Element Id"] == (int)selectedData["Element Id"] &&
                    ed.Data["Document"].ToString() == selectedData["Document"].ToString());

                if (matchingElemData != null)
                {
                    finalReferences.Add(matchingElemData.Reference);
                }
            }

            // Update the selection using SetReferences to maintain linked element references
            uidoc.SetReferences(finalReferences);
        }

        return Result.Succeeded;
    }

    private void FindNearestGrids(XYZ location, List<Grid> allGrids, Dictionary<string, object> properties)
    {
        if (location == null || !allGrids.Any())
            return;

        // Separate grids by direction (roughly X and Y)
        List<Grid> xGrids = new List<Grid>();
        List<Grid> yGrids = new List<Grid>();

        foreach (Grid grid in allGrids)
        {
            Curve gridCurve = grid.Curve;
            if (gridCurve == null) continue;

            XYZ start = gridCurve.GetEndPoint(0);
            XYZ end = gridCurve.GetEndPoint(1);
            XYZ direction = (end - start).Normalize();

            // Determine if grid is more X-aligned or Y-aligned
            if (Math.Abs(direction.X) > Math.Abs(direction.Y))
            {
                xGrids.Add(grid);
            }
            else
            {
                yGrids.Add(grid);
            }
        }

        // Find nearest X grids (up and down in Y direction)
        Grid xGridUp = null;
        Grid xGridDown = null;
        double minDistUp = double.MaxValue;
        double minDistDown = double.MaxValue;

        foreach (Grid grid in xGrids)
        {
            // Project the location onto the grid line to find the closest point
            Curve gridCurve = grid.Curve;
            IntersectionResult result = gridCurve.Project(location);
            if (result != null)
            {
                XYZ closestPoint = result.XYZPoint;
                double yDiff = closestPoint.Y - location.Y;

                if (yDiff > 0 && Math.Abs(yDiff) < minDistUp)
                {
                    minDistUp = Math.Abs(yDiff);
                    xGridUp = grid;
                }
                else if (yDiff < 0 && Math.Abs(yDiff) < minDistDown)
                {
                    minDistDown = Math.Abs(yDiff);
                    xGridDown = grid;
                }
            }
        }

        // Find nearest Y grids (up and down in X direction)
        Grid yGridUp = null;
        Grid yGridDown = null;
        minDistUp = double.MaxValue;
        minDistDown = double.MaxValue;

        foreach (Grid grid in yGrids)
        {
            // Project the location onto the grid line
            Curve gridCurve = grid.Curve;
            IntersectionResult result = gridCurve.Project(location);
            if (result != null)
            {
                XYZ closestPoint = result.XYZPoint;
                double xDiff = closestPoint.X - location.X;

                if (xDiff > 0 && Math.Abs(xDiff) < minDistUp)
                {
                    minDistUp = Math.Abs(xDiff);
                    yGridUp = grid;
                }
                else if (xDiff < 0 && Math.Abs(xDiff) < minDistDown)
                {
                    minDistDown = Math.Abs(xDiff);
                    yGridDown = grid;
                }
            }
        }

        // Set the grid properties
        properties["Grid X Up"] = xGridUp != null ? xGridUp.Name : "";
        properties["Grid X Down"] = xGridDown != null ? xGridDown.Name : "";
        properties["Grid Y Up"] = yGridUp != null ? yGridUp.Name : "";
        properties["Grid Y Down"] = yGridDown != null ? yGridDown.Name : "";
    }

    private BoundingBoxXYZ GetTransformedBoundingBox(BoundingBoxXYZ bb, Transform transform)
    {
        if (bb == null) return null;

        XYZ[] corners = new XYZ[8];
        corners[0] = bb.Min;
        corners[1] = new XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z);
        corners[2] = new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z);
        corners[3] = new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z);
        corners[4] = new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z);
        corners[5] = new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z);
        corners[6] = new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z);
        corners[7] = bb.Max;

        double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

        foreach (XYZ p in corners)
        {
            XYZ tp = transform.OfPoint(p);
            if (tp.X < minX) minX = tp.X;
            if (tp.Y < minY) minY = tp.Y;
            if (tp.Z < minZ) minZ = tp.Z;
            if (tp.X > maxX) maxX = tp.X;
            if (tp.Y > maxY) maxY = tp.Y;
            if (tp.Z > maxZ) maxZ = tp.Z;
        }

        return new BoundingBoxXYZ { Min = new XYZ(minX, minY, minZ), Max = new XYZ(maxX, maxY, maxZ) };
    }

    private bool Contains(BoundingBoxXYZ outer, BoundingBoxXYZ inner)
    {
        const double epsilon = 1e-9;
        if (outer == null || inner == null) return false;

        return inner.Min.X >= outer.Min.X - epsilon && inner.Max.X <= outer.Max.X + epsilon &&
               inner.Min.Y >= outer.Min.Y - epsilon && inner.Max.Y <= outer.Max.Y + epsilon &&
               inner.Min.Z >= outer.Min.Z - epsilon && inner.Max.Z <= outer.Max.Z + epsilon;
    }
}
