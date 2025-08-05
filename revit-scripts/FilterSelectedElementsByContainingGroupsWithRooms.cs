using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;

[Transaction(TransactionMode.Manual)]
public class FilterSelectedElementsByContainingGroupsWithRooms : IExternalCommand
{
    // Cache for element test points to avoid recalculation
    private Dictionary<ElementId, List<XYZ>> _elementTestPointsCache = new Dictionary<ElementId, List<XYZ>>();
    
    // Cache for room data
    private Dictionary<ElementId, RoomData> _roomDataCache = new Dictionary<ElementId, RoomData>();
    
    private class RoomData
    {
        public double LevelZ { get; set; }
        public bool IsUnbound { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
    }
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        try
        {
            // Get selected elements
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            
            if (selectedIds.Count == 0)
            {
                message = "Please select elements first";
                return Result.Failed;
            }
            
            // Get selected elements (exclude groups and element types)
            List<Element> selectedElements = new List<Element>();
            
            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem != null && !(elem is Group) && !(elem is ElementType))
                {
                    selectedElements.Add(elem);
                }
            }
            
            if (selectedElements.Count == 0)
            {
                message = "No valid elements found in selection. Please select elements (not groups).";
                return Result.Failed;
            }
            
            // Pre-calculate test points for all selected elements (OPTIMIZATION)
            foreach (Element elem in selectedElements)
            {
                _elementTestPointsCache[elem.Id] = GetElementTestPoints(elem);
            }
            
            // Create bounding box for all selected elements for quick filtering
            BoundingBoxXYZ overallBB = GetOverallBoundingBox(selectedElements);
            if (overallBB == null)
            {
                message = "Could not determine bounding box of selected elements";
                return Result.Failed;
            }
            
            // Find all groups in the document
            FilteredElementCollector groupCollector = new FilteredElementCollector(doc);
            IList<Group> allGroups = groupCollector
                .OfClass(typeof(Group))
                .WhereElementIsNotElementType()
                .Cast<Group>()
                .ToList();
            
            // Dictionary to store which groups contain which elements
            Dictionary<ElementId, List<Group>> elementToGroupsMap = new Dictionary<ElementId, List<Group>>();
            
            // Initialize the map
            foreach (Element elem in selectedElements)
            {
                elementToGroupsMap[elem.Id] = new List<Group>();
            }
            
            // Process each group with spatial filtering
            foreach (Group group in allGroups)
            {
                // Quick bounding box check first
                BoundingBoxXYZ groupBB = GetGroupBoundingBox(group, doc);
                if (groupBB == null || !BoundingBoxesIntersect(overallBB, groupBB))
                {
                    continue; // Skip this group - bounding boxes don't intersect
                }
                
                // Check which selected elements are contained in this group's rooms
                List<Element> containedElements = GetElementsContainedInGroupRoomsOptimized(group, selectedElements, doc);
                
                // Update the map
                foreach (Element elem in containedElements)
                {
                    elementToGroupsMap[elem.Id].Add(group);
                }
            }
            
            // Find maximum number of groups any single element is contained in
            int maxGroupsPerElement = 0;
            foreach (var groups in elementToGroupsMap.Values)
            {
                maxGroupsPerElement = Math.Max(maxGroupsPerElement, groups.Count);
            }
            
            if (maxGroupsPerElement == 0)
            {
                TaskDialog.Show("No Groups Found", 
                    "No groups were found containing the selected elements within their rooms.");
                return Result.Succeeded;
            }
            
            // Prepare data for DataGrid
            List<Dictionary<string, object>> gridEntries = new List<Dictionary<string, object>>();
            List<string> propertyNames = new List<string> { "Type", "Family", "Id", "Comments" };
            
            // Add columns for groups (Group 1, Group 2, etc.)
            for (int i = 1; i <= maxGroupsPerElement; i++)
            {
                propertyNames.Add($"Group {i}");
            }
            
            // Create entries for each element
            foreach (Element elem in selectedElements)
            {
                Dictionary<string, object> entry = new Dictionary<string, object>();
                
                // Type
                ElementType elemType = doc.GetElement(elem.GetTypeId()) as ElementType;
                entry["Type"] = elemType?.Name ?? "Unknown";
                
                // Family
                string familyName = "Unknown";
                if (elemType != null)
                {
                    FamilySymbol famSymbol = elemType as FamilySymbol;
                    if (famSymbol != null && famSymbol.Family != null)
                    {
                        familyName = famSymbol.Family.Name;
                    }
                    else
                    {
                        // For system families, use category name
                        familyName = elem.Category?.Name ?? "System Family";
                    }
                }
                entry["Family"] = familyName;
                
                // Id
                entry["Id"] = elem.Id.IntegerValue.ToString();
                
                // Comments
                Parameter commentsParam = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                string comments = commentsParam?.AsString() ?? "";
                entry["Comments"] = comments;
                
                // Store the actual element object for later retrieval
                entry["_ElementObject"] = elem;
                
                // Get groups containing this element
                List<Group> containingGroups = elementToGroupsMap[elem.Id];
                
                // Sort groups by name for consistent ordering
                containingGroups = containingGroups.OrderBy(g => 
                {
                    GroupType gt = doc.GetElement(g.GetTypeId()) as GroupType;
                    return gt?.Name ?? "Unknown";
                }).ToList();
                
                // Fill in group columns
                for (int i = 0; i < maxGroupsPerElement; i++)
                {
                    string columnName = $"Group {i + 1}";
                    if (i < containingGroups.Count)
                    {
                        Group group = containingGroups[i];
                        GroupType groupType = doc.GetElement(group.GetTypeId()) as GroupType;
                        string groupName = groupType?.Name ?? "Unknown";
                        entry[columnName] = groupName;
                    }
                    else
                    {
                        entry[columnName] = "";
                    }
                }
                
                gridEntries.Add(entry);
            }
            
            // Show DataGrid to user
            List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(
                gridEntries,
                propertyNames,
                false,
                null
            );
            
            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                // User cancelled or selected nothing
                return Result.Cancelled;
            }
            
            // Determine which elements to select based on selected entries
            HashSet<Element> elementsToSelect = new HashSet<Element>();
            
            foreach (var entry in selectedEntries)
            {
                // Get the element object from the hidden key
                if (entry.ContainsKey("_ElementObject") && entry["_ElementObject"] is Element)
                {
                    Element elem = entry["_ElementObject"] as Element;
                    elementsToSelect.Add(elem);
                }
            }
            
            // Set selection to chosen elements
            if (elementsToSelect.Count > 0)
            {
                List<ElementId> elementIds = elementsToSelect.Select(e => e.Id).ToList();
                uidoc.Selection.SetElementIds(elementIds);
            }
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
    
    // Pre-calculate test points for an element (OPTIMIZATION)
    private List<XYZ> GetElementTestPoints(Element element)
    {
        List<XYZ> testPoints = new List<XYZ>();
        
        LocationPoint locPoint = element.Location as LocationPoint;
        LocationCurve locCurve = element.Location as LocationCurve;
        
        if (locPoint != null)
        {
            testPoints.Add(locPoint.Point);
        }
        else if (locCurve != null)
        {
            // For curves, check endpoints
            Curve curve = locCurve.Curve;
            testPoints.Add(curve.GetEndPoint(0));
            testPoints.Add(curve.GetEndPoint(1));
        }
        else
        {
            // For other elements, use bounding box center
            BoundingBoxXYZ bb = element.get_BoundingBox(null);
            if (bb != null)
            {
                testPoints.Add((bb.Min + bb.Max) * 0.5);
            }
        }
        
        return testPoints;
    }
    
    // Get overall bounding box for a list of elements
    private BoundingBoxXYZ GetOverallBoundingBox(List<Element> elements)
    {
        if (elements.Count == 0) return null;
        
        double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
        
        bool hasValidBB = false;
        
        foreach (Element elem in elements)
        {
            BoundingBoxXYZ bb = elem.get_BoundingBox(null);
            if (bb != null)
            {
                hasValidBB = true;
                minX = Math.Min(minX, bb.Min.X);
                minY = Math.Min(minY, bb.Min.Y);
                minZ = Math.Min(minZ, bb.Min.Z);
                maxX = Math.Max(maxX, bb.Max.X);
                maxY = Math.Max(maxY, bb.Max.Y);
                maxZ = Math.Max(maxZ, bb.Max.Z);
            }
        }
        
        if (!hasValidBB) return null;
        
        BoundingBoxXYZ overallBB = new BoundingBoxXYZ();
        overallBB.Min = new XYZ(minX, minY, minZ);
        overallBB.Max = new XYZ(maxX, maxY, maxZ);
        
        return overallBB;
    }
    
    // Get bounding box for a group
    private BoundingBoxXYZ GetGroupBoundingBox(Group group, Document doc)
    {
        ICollection<ElementId> memberIds = group.GetMemberIds();
        if (memberIds.Count == 0) return null;
        
        double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
        
        bool hasValidBB = false;
        
        foreach (ElementId id in memberIds)
        {
            Element member = doc.GetElement(id);
            if (member != null)
            {
                BoundingBoxXYZ bb = member.get_BoundingBox(null);
                if (bb != null)
                {
                    hasValidBB = true;
                    minX = Math.Min(minX, bb.Min.X);
                    minY = Math.Min(minY, bb.Min.Y);
                    minZ = Math.Min(minZ, bb.Min.Z);
                    maxX = Math.Max(maxX, bb.Max.X);
                    maxY = Math.Max(maxY, bb.Max.Y);
                    maxZ = Math.Max(maxZ, bb.Max.Z);
                }
            }
        }
        
        if (!hasValidBB) return null;
        
        BoundingBoxXYZ groupBB = new BoundingBoxXYZ();
        groupBB.Min = new XYZ(minX, minY, minZ);
        groupBB.Max = new XYZ(maxX, maxY, maxZ);
        
        return groupBB;
    }
    
    // Check if two bounding boxes intersect
    private bool BoundingBoxesIntersect(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        return !(bb1.Max.X < bb2.Min.X || bb2.Max.X < bb1.Min.X ||
                 bb1.Max.Y < bb2.Min.Y || bb2.Max.Y < bb1.Min.Y ||
                 bb1.Max.Z < bb2.Min.Z || bb2.Max.Z < bb1.Min.Z);
    }
    
    // Optimized version that uses cached data
    private List<Element> GetElementsContainedInGroupRoomsOptimized(Group group, List<Element> candidateElements, Document doc)
    {
        List<Element> containedElements = new List<Element>();
        
        // Get all rooms that are members of this group
        List<Room> groupRooms = new List<Room>();
        ICollection<ElementId> memberIds = group.GetMemberIds();
        
        // Also get the overall height range of the group
        double groupMinZ = double.MaxValue;
        double groupMaxZ = double.MinValue;
        
        foreach (ElementId id in memberIds)
        {
            Element member = doc.GetElement(id);
            if (member is Room)
            {
                Room room = member as Room;
                groupRooms.Add(room);
                
                // Cache room data if not already cached
                if (!_roomDataCache.ContainsKey(room.Id))
                {
                    RoomData roomData = new RoomData();
                    roomData.IsUnbound = room.Volume <= 0.001;
                    roomData.BoundingBox = room.get_BoundingBox(null);
                    
                    if (room.LevelId != null && room.LevelId != ElementId.InvalidElementId)
                    {
                        Level roomLevel = doc.GetElement(room.LevelId) as Level;
                        if (roomLevel != null)
                        {
                            roomData.LevelZ = roomLevel.Elevation;
                        }
                    }
                    
                    _roomDataCache[room.Id] = roomData;
                }
            }
            
            // Update group height bounds
            BoundingBoxXYZ bb = member.get_BoundingBox(null);
            if (bb != null)
            {
                groupMinZ = Math.Min(groupMinZ, bb.Min.Z);
                groupMaxZ = Math.Max(groupMaxZ, bb.Max.Z);
            }
        }
        
        if (groupRooms.Count == 0)
        {
            return containedElements;
        }
        
        // Add some tolerance to the height bounds
        groupMinZ -= 1.0; // 1 foot below
        groupMaxZ += 1.0; // 1 foot above
        
        // Check each candidate element against all rooms
        foreach (Element elem in candidateElements)
        {
            // Get cached test points
            List<XYZ> testPoints = _elementTestPointsCache[elem.Id];
            
            foreach (Room room in groupRooms)
            {
                if (IsElementInRoomOptimized(elem, room, testPoints, groupMinZ, groupMaxZ))
                {
                    containedElements.Add(elem);
                    break; // Element is in at least one room, no need to check others
                }
            }
        }
        
        return containedElements;
    }
    
    // Optimized version using cached test points and room data
    private bool IsElementInRoomOptimized(Element element, Room room, List<XYZ> testPoints, double groupMinZ, double groupMaxZ)
    {
        // Skip unplaced or unenclosed rooms
        if (room.Area <= 0)
            return false;
        
        // Get cached room data
        RoomData roomData = _roomDataCache[room.Id];
        
        // Check if any test point is within the room
        foreach (XYZ point in testPoints)
        {
            // First check Z coordinate against group bounds
            if (point.Z < groupMinZ || point.Z > groupMaxZ)
                continue;
            
            // For ALL rooms (bounded or unbound), check XY containment at room level
            double testZ = roomData.LevelZ + 1.0; // 1 foot above room level
            if (!roomData.IsUnbound && roomData.BoundingBox != null)
            {
                // For bounded rooms, use a point within the room's height range
                testZ = (roomData.BoundingBox.Min.Z + roomData.BoundingBox.Max.Z) * 0.5; // Middle of room height
            }
            
            // Create test point at appropriate Z for XY containment check
            XYZ testPointForXY = new XYZ(point.X, point.Y, testZ);
            bool inRoomXY = room.IsPointInRoom(testPointForXY);
            
            if (inRoomXY)
            {
                return true;
            }
        }
        
        return false;
    }
}
