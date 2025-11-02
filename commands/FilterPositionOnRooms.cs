using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.ReadOnly)]
public class FilterPositionOnRooms : IExternalCommand
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

    // Helper class to store room information
    private class RoomInfo
    {
        public Room Room { get; set; }
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

        // Get all rooms in the current document
        List<RoomInfo> allRooms = GetAllRoomsWithParameters(doc);

        if (!allRooms.Any())
        {
            TaskDialog.Show("Warning", "No rooms found in the current document.");
            return Result.Failed;
        }

        // Get all unique room parameter names for column headers
        HashSet<string> roomParameterNames = new HashSet<string>();
        foreach (var roomInfo in allRooms)
        {
            foreach (var paramName in roomInfo.Parameters.Keys)
            {
                roomParameterNames.Add($"Room_{paramName}");
            }
        }

        // Define the property names (columns) for the data grid.
        List<string> propertyNames = new List<string>
        {
            "Element Id",
            "Document",
            "Category",
            "Name",
            "Located In Rooms",
            "Room Count"
        };

        // Add room parameter columns
        propertyNames.AddRange(roomParameterNames.OrderBy(n => n));

        // Add position columns
        propertyNames.AddRange(new[] { "X", "Y", "Z" });

        // Prepare the list of dictionaries that will hold each element's data.
        List<Dictionary<string, object>> gridData = new List<Dictionary<string, object>>();

        // Process elements - batch them for better performance
        int totalElements = elementDataList.Count;
        int processedElements = 0;
        
        foreach (var elemData in elementDataList)
        {
            Element elem = elemData.Element;
            Dictionary<string, object> properties = new Dictionary<string, object>();

            // Basic element properties.
            properties["Element Id"] = elem.Id.IntegerValue;
            properties["Document"] = elemData.DocumentName + (elemData.IsLinked ? " (Linked)" : "");
            properties["Category"] = elem.Category != null ? elem.Category.Name : "";
            properties["Name"] = elem.Name;

            // Initialize room-related fields
            properties["Located In Rooms"] = "";
            properties["Room Count"] = 0;

            // Initialize all room parameter columns
            foreach (var paramName in roomParameterNames)
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

            // Find rooms containing this element
            if (elemData.LocationPoint != null)
            {
                List<RoomInfo> containingRooms = FindContainingRooms(elemData.LocationPoint, allRooms);
                
                if (containingRooms.Any())
                {
                    // Set room names
                    properties["Located In Rooms"] = string.Join(", ", containingRooms.Select(r => r.Room.Name ?? $"Room {r.Room.Number}"));
                    properties["Room Count"] = containingRooms.Count;

                    // Aggregate room parameters
                    // If element is in multiple rooms, concatenate values with semicolon
                    Dictionary<string, List<string>> aggregatedParams = new Dictionary<string, List<string>>();
                    
                    foreach (var roomInfo in containingRooms)
                    {
                        foreach (var kvp in roomInfo.Parameters)
                        {
                            string columnName = $"Room_{kvp.Key}";
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

            processedElements++;
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
    /// Get all rooms in the document with their parameters
    /// </summary>
    private List<RoomInfo> GetAllRoomsWithParameters(Document doc)
    {
        List<RoomInfo> roomInfos = new List<RoomInfo>();

        // Get all room elements
        var rooms = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType()
            .Cast<Room>()
            .Where(r => r.Area > 0) // Only valid rooms with positive area
            .ToList();

        // Pre-process spatial boundary options once
        SpatialElementBoundaryOptions boundaryOptions = new SpatialElementBoundaryOptions();
        boundaryOptions.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;

        foreach (var room in rooms)
        {
            var roomInfo = new RoomInfo
            {
                Room = room,
                Parameters = new Dictionary<string, string>(),
                Level = room.Level,
                LevelElevation = room.Level?.Elevation ?? 0,
                BoundaryPolygons = new List<List<XYZ>>()
            };

            // Pre-process room boundaries
            try
            {
                IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(boundaryOptions);
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
                            roomInfo.BoundaryPolygons.Add(cleanedPoints);
                        }
                    }

                    // Set bounding box
                    if (minX < double.MaxValue)
                    {
                        roomInfo.BoundingBox = new BoundingBoxXYZ
                        {
                            Min = new XYZ(minX, minY, roomInfo.LevelElevation - 1),
                            Max = new XYZ(maxX, maxY, roomInfo.LevelElevation + 10)
                        };
                    }
                }
            }
            catch { /* Skip rooms with problematic boundaries */ }

            // Collect room parameters
            foreach (Parameter param in room.Parameters)
            {
                try
                {
                    string paramName = param.Definition.Name;
                    string paramValue = GetParameterValue(param);
                    
                    // Skip certain system parameters that might not be useful
                    if (!string.IsNullOrEmpty(paramName) && !string.IsNullOrEmpty(paramValue))
                    {
                        roomInfo.Parameters[paramName] = paramValue;
                    }
                }
                catch { /* Skip problematic parameters */ }
            }

            // Only add rooms that have valid boundaries
            if (roomInfo.BoundaryPolygons.Any())
            {
                roomInfos.Add(roomInfo);
            }
        }

        return roomInfos;
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
    /// Find all rooms that contain the given point
    /// </summary>
    private List<RoomInfo> FindContainingRooms(XYZ point, List<RoomInfo> allRooms)
    {
        List<RoomInfo> containingRooms = new List<RoomInfo>();

        // Group rooms by level for faster lookup
        var roomsByLevel = allRooms.GroupBy(r => r.LevelElevation).ToList();

        foreach (var levelGroup in roomsByLevel)
        {
            double levelElevation = levelGroup.Key;
            double tolerance = 10.0; // 10 feet tolerance for level matching
            
            // Skip this level if point is not within tolerance
            if (Math.Abs(point.Z - levelElevation) > tolerance)
                continue;

            // Check rooms at this level
            foreach (var roomInfo in levelGroup)
            {
                // Quick bounding box check first
                if (roomInfo.BoundingBox != null)
                {
                    if (point.X < roomInfo.BoundingBox.Min.X || point.X > roomInfo.BoundingBox.Max.X ||
                        point.Y < roomInfo.BoundingBox.Min.Y || point.Y > roomInfo.BoundingBox.Max.Y)
                    {
                        continue; // Point is outside bounding box, skip detailed check
                    }
                }

                // Check if point is inside room boundary using pre-computed polygons
                if (IsPointInPolygons(new XYZ(point.X, point.Y, 0), roomInfo.BoundaryPolygons))
                {
                    containingRooms.Add(roomInfo);
                }
            }
        }

        return containingRooms;
    }

    /// <summary>
    /// Check if a point is inside pre-computed boundary polygons using ray casting algorithm
    /// </summary>
    private bool IsPointInPolygons(XYZ point, List<List<XYZ>> polygons)
    {
        // Use ray casting algorithm - count how many times a ray from the point crosses boundaries
        int crossings = 0;

        foreach (var polygon in polygons)
        {
            if (polygon.Count < 3)
                continue; // Not a valid polygon

            // Ray casting algorithm
            for (int i = 0; i < polygon.Count; i++)
            {
                XYZ p1 = polygon[i];
                XYZ p2 = polygon[(i + 1) % polygon.Count];

                // Check if the ray crosses this edge
                if ((p1.Y <= point.Y && point.Y < p2.Y) || (p2.Y <= point.Y && point.Y < p1.Y))
                {
                    // Calculate X coordinate of intersection
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
