using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.ReadOnly)]
public class FilterPositionOnAreas : IExternalCommand
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
        public XYZ LocationPoint { get; set; }
    }

    // Helper class to store area information
    private class AreaInfo
    {
        public Area Area { get; set; }
        public ViewPlan AreaPlan { get; set; }
        public Dictionary<string, string> Parameters { get; set; }
        public Level Level { get; set; }
        public double LevelElevation { get; set; }
        public List<List<XYZ>> BoundaryPolygons { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
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

        // Process references to get elements (both from current doc and linked docs)
        List<ElementDataWithReference> elementDataList = new List<ElementDataWithReference>();

        foreach (Reference reference in selectedRefs)
        {
            Element elem = null;
            bool isLinked = false;
            string documentName = doc.Title;
            RevitLinkInstance linkInstance = null;

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
                    // Get element location point
                    XYZ locationPoint = GetElementLocationPoint(elem, linkInstance);
                    
                    elementDataList.Add(new ElementDataWithReference
                    {
                        Element = elem,
                        Reference = reference,
                        IsLinked = isLinked,
                        DocumentName = documentName,
                        LinkInstance = linkInstance,
                        LocationPoint = locationPoint
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

        // Get all areas in the current document
        List<AreaInfo> allAreas = GetAllAreasWithParameters(doc);

        if (!allAreas.Any())
        {
            TaskDialog.Show("Warning", "No areas found in the current document.");
            return Result.Failed;
        }

        // Get all unique area parameter names for column headers
        HashSet<string> areaParameterNames = new HashSet<string>();
        foreach (var areaInfo in allAreas)
        {
            foreach (var paramName in areaInfo.Parameters.Keys)
            {
                areaParameterNames.Add($"Area_{paramName}");
            }
        }

        // Define the property names (columns) for the data grid.
        List<string> propertyNames = new List<string>
        {
            "Element Id",
            "Document",
            "Category",
            "Name",
            "Located In Areas",
            "Area Count"
        };

        // Add area parameter columns
        propertyNames.AddRange(areaParameterNames.OrderBy(n => n));

        // Add position columns
        propertyNames.AddRange(new[] { "X", "Y", "Z" });

        // Prepare the list of dictionaries that will hold each element's data.
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

            // Initialize area-related fields
            properties["Located In Areas"] = "";
            properties["Area Count"] = 0;

            // Initialize all area parameter columns
            foreach (var paramName in areaParameterNames)
            {
                properties[paramName] = "";
            }

            // Position data
            if (elemData.LocationPoint != null)
            {
                properties["X"] = elemData.LocationPoint.X;
                properties["Y"] = elemData.LocationPoint.Y;
                properties["Z"] = elemData.LocationPoint.Z;
            }
            else
            {
                properties["X"] = "";
                properties["Y"] = "";
                properties["Z"] = "";
            }

            // Find areas containing this element
            if (elemData.LocationPoint != null)
            {
                List<AreaInfo> containingAreas = FindContainingAreas(elemData.LocationPoint, allAreas);
                
                if (containingAreas.Any())
                {
                    // Set area names
                    properties["Located In Areas"] = string.Join(", ", containingAreas.Select(a => a.Area.Name ?? $"Area {a.Area.Number}"));
                    properties["Area Count"] = containingAreas.Count;

                    // Aggregate area parameters
                    // If element is in multiple areas, concatenate values with semicolon
                    Dictionary<string, List<string>> aggregatedParams = new Dictionary<string, List<string>>();
                    
                    foreach (var areaInfo in containingAreas)
                    {
                        foreach (var kvp in areaInfo.Parameters)
                        {
                            string columnName = $"Area_{kvp.Key}";
                            if (!aggregatedParams.ContainsKey(columnName))
                                aggregatedParams[columnName] = new List<string>();
                            
                            if (!string.IsNullOrEmpty(kvp.Value))
                                aggregatedParams[columnName].Add(kvp.Value);
                        }
                    }

                    // Set aggregated values
                    foreach (var kvp in aggregatedParams)
                    {
                        properties[kvp.Key] = string.Join("; ", kvp.Value.Distinct());
                    }
                }
            }

            // Store the data along with the index for later reference
            elemData.Data = properties;
            gridData.Add(properties);
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

    /// <summary>
    /// Get the location point of an element, considering linked instances
    /// </summary>
    private XYZ GetElementLocationPoint(Element elem, RevitLinkInstance linkInstance)
    {
        XYZ point = null;

        // Try to get location from Location property
        Location loc = elem.Location;
        if (loc is LocationPoint locPoint)
        {
            point = locPoint.Point;
        }
        else if (loc is LocationCurve locCurve)
        {
            // For curves, use midpoint
            Curve curve = locCurve.Curve;
            if (curve != null)
            {
                point = curve.Evaluate(0.5, true);
            }
        }
        
        // If no location, try bounding box center
        if (point == null)
        {
            BoundingBoxXYZ bb = elem.get_BoundingBox(null);
            if (bb != null)
            {
                point = (bb.Min + bb.Max) / 2;
            }
        }

        // Transform point if it's from a linked instance
        if (point != null && linkInstance != null)
        {
            Transform transform = linkInstance.GetTotalTransform();
            point = transform.OfPoint(point);
        }

        return point;
    }

    /// <summary>
    /// Get all areas in the document with their parameters
    /// </summary>
    private List<AreaInfo> GetAllAreasWithParameters(Document doc)
    {
        List<AreaInfo> areaInfos = new List<AreaInfo>();

        // Get all area elements
        var areas = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Areas)
            .WhereElementIsNotElementType()
            .Cast<Area>()
            .Where(a => a.Area > 0) // Only valid areas with positive area
            .ToList();

        // Pre-fetch all area plans in one go for better performance
        var allAreaPlans = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewPlan))
            .Cast<ViewPlan>()
            .Where(v => v.AreaScheme != null)
            .ToList();

        // Create a lookup dictionary for area plans
        var areaPlanLookup = new Dictionary<string, ViewPlan>();
        foreach (var plan in allAreaPlans)
        {
            if (plan.AreaScheme != null && plan.GenLevel != null)
            {
                string key = $"{plan.AreaScheme.Id.IntegerValue}_{plan.GenLevel.Id.IntegerValue}";
                areaPlanLookup[key] = plan;
            }
        }

        // Pre-process spatial boundary options once
        SpatialElementBoundaryOptions boundaryOptions = new SpatialElementBoundaryOptions();

        foreach (var area in areas)
        {
            var areaInfo = new AreaInfo
            {
                Area = area,
                Parameters = new Dictionary<string, string>(),
                Level = area.Level,
                LevelElevation = area.Level?.Elevation ?? 0,
                BoundaryPolygons = new List<List<XYZ>>()
            };

            // Get the area plan view using lookup
            if (area.AreaScheme != null && area.Level != null)
            {
                string key = $"{area.AreaScheme.Id.IntegerValue}_{area.Level.Id.IntegerValue}";
                if (areaPlanLookup.TryGetValue(key, out ViewPlan areaPlan))
                {
                    areaInfo.AreaPlan = areaPlan;
                }
            }

            // Pre-process area boundaries
            try
            {
                IList<IList<BoundarySegment>> boundaries = area.GetBoundarySegments(boundaryOptions);
                if (boundaries != null && boundaries.Any())
                {
                    double minX = double.MaxValue, minY = double.MaxValue;
                    double maxX = double.MinValue, maxY = double.MinValue;

                    foreach (var boundary in boundaries)
                    {
                        List<XYZ> polygonPoints = new List<XYZ>();
                        
                        foreach (var segment in boundary)
                        {
                            Curve curve = segment.GetCurve();
                            
                            // Tessellate curves for better accuracy
                            IList<XYZ> tessellatedPoints = curve.Tessellate();
                            foreach (var pt in tessellatedPoints)
                            {
                                XYZ pt2D = new XYZ(pt.X, pt.Y, 0);
                                polygonPoints.Add(pt2D);
                                
                                // Update bounding box
                                minX = Math.Min(minX, pt.X);
                                minY = Math.Min(minY, pt.Y);
                                maxX = Math.Max(maxX, pt.X);
                                maxY = Math.Max(maxY, pt.Y);
                            }
                        }
                        
                        // Remove duplicate points
                        if (polygonPoints.Count > 0)
                        {
                            var cleanedPoints = new List<XYZ> { polygonPoints[0] };
                            for (int i = 1; i < polygonPoints.Count; i++)
                            {
                                if (!polygonPoints[i].IsAlmostEqualTo(cleanedPoints.Last(), 0.001))
                                {
                                    cleanedPoints.Add(polygonPoints[i]);
                                }
                            }
                            areaInfo.BoundaryPolygons.Add(cleanedPoints);
                        }
                    }

                    // Set bounding box
                    if (minX < double.MaxValue)
                    {
                        areaInfo.BoundingBox = new BoundingBoxXYZ
                        {
                            Min = new XYZ(minX, minY, areaInfo.LevelElevation - 1),
                            Max = new XYZ(maxX, maxY, areaInfo.LevelElevation + 10)
                        };
                    }
                }
            }
            catch { /* Skip areas with problematic boundaries */ }

            // Collect area parameters
            foreach (Parameter param in area.Parameters)
            {
                try
                {
                    string paramName = param.Definition.Name;
                    string paramValue = GetParameterValue(param);
                    
                    // Skip certain system parameters that might not be useful
                    if (!string.IsNullOrEmpty(paramName) && !string.IsNullOrEmpty(paramValue))
                    {
                        areaInfo.Parameters[paramName] = paramValue;
                    }
                }
                catch { /* Skip problematic parameters */ }
            }

            // Only add areas that have valid boundaries
            if (areaInfo.BoundaryPolygons.Any())
            {
                areaInfos.Add(areaInfo);
            }
        }

        return areaInfos;
    }

    /// <summary>
    /// Get parameter value as string
    /// </summary>
    private string GetParameterValue(Parameter param)
    {
        if (!param.HasValue)
            return "";

        switch (param.StorageType)
        {
            case StorageType.Double:
                return param.AsValueString() ?? param.AsDouble().ToString();
            case StorageType.Integer:
                return param.AsInteger().ToString();
            case StorageType.String:
                return param.AsString() ?? "";
            case StorageType.ElementId:
                ElementId id = param.AsElementId();
                if (id.IntegerValue > 0)
                {
                    Element elem = param.Element.Document.GetElement(id);
                    return elem?.Name ?? id.ToString();
                }
                return "";
            default:
                return "";
        }
    }

    /// <summary>
    /// Find all areas that contain the given point
    /// </summary>
    private List<AreaInfo> FindContainingAreas(XYZ point, List<AreaInfo> allAreas)
    {
        List<AreaInfo> containingAreas = new List<AreaInfo>();

        foreach (var areaInfo in allAreas)
        {
            Area area = areaInfo.Area;
            
            // Check if point is at the same level (within tolerance)
            Level areaLevel = area.Level;
            if (areaLevel != null)
            {
                double levelElevation = areaLevel.Elevation;
                double tolerance = 10.0; // 10 feet tolerance for level matching
                
                if (Math.Abs(point.Z - levelElevation) > tolerance)
                    continue; // Point is not at this level
            }

            // Get area boundary
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
            IList<IList<BoundarySegment>> boundaries = area.GetBoundarySegments(options);
            
            if (boundaries == null || !boundaries.Any())
                continue;

            // Check if point is inside area boundary (using 2D check)
            if (IsPointInAreaBoundary(new XYZ(point.X, point.Y, 0), boundaries))
            {
                containingAreas.Add(areaInfo);
            }
        }

        return containingAreas;
    }

    /// <summary>
    /// Check if a point is inside area boundaries using ray casting algorithm
    /// </summary>
    private bool IsPointInAreaBoundary(XYZ point, IList<IList<BoundarySegment>> boundaries)
    {
        // Use ray casting algorithm - count how many times a ray from the point crosses boundaries
        int crossings = 0;

        foreach (var boundary in boundaries)
        {
            List<XYZ> polygonPoints = new List<XYZ>();
            
            foreach (var segment in boundary)
            {
                Curve curve = segment.GetCurve();
                // For simplicity, we'll use start and end points of curves
                // For more accuracy, you could tessellate curves
                polygonPoints.Add(new XYZ(curve.GetEndPoint(0).X, curve.GetEndPoint(0).Y, 0));
            }

            // Ray casting algorithm
            for (int i = 0; i < polygonPoints.Count; i++)
            {
                XYZ p1 = polygonPoints[i];
                XYZ p2 = polygonPoints[(i + 1) % polygonPoints.Count];

                if ((p1.Y <= point.Y && point.Y < p2.Y) || (p2.Y <= point.Y && point.Y < p1.Y))
                {
                    double x = p1.X + (point.Y - p1.Y) * (p2.X - p1.X) / (p2.Y - p1.Y);
                    if (point.X < x)
                        crossings++;
                }
            }
        }

        // If odd number of crossings, point is inside
        return (crossings % 2) == 1;
    }
}
