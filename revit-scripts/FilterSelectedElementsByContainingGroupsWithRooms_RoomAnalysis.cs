using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

public partial class FilterSelectedElementsByContainingGroupsWithRooms
{
    // Cache for room data
    private Dictionary<ElementId, RoomData> _roomDataCache = new Dictionary<ElementId, RoomData>();
    
    // Cache for group bounding boxes
    private Dictionary<ElementId, BoundingBoxXYZ> _groupBoundingBoxCache = new Dictionary<ElementId, BoundingBoxXYZ>();
    
    // Spatial index for rooms
    private class SpatialIndex
    {
        private Dictionary<int, List<Room>> _grid = new Dictionary<int, List<Room>>();
        private double _cellSize = 10.0; // 10 feet grid cells
        
        public void AddRoom(Room room)
        {
            BoundingBoxXYZ bb = room.get_BoundingBox(null);
            if (bb == null) return;
            
            int minX = GetGridIndex(bb.Min.X);
            int maxX = GetGridIndex(bb.Max.X);
            int minY = GetGridIndex(bb.Min.Y);
            int maxY = GetGridIndex(bb.Max.Y);
            
            for (int x = minX; x <= maxX; x++)
            {
                for (int y = minY; y <= maxY; y++)
                {
                    int key = GetCellKey(x, y);
                    if (!_grid.ContainsKey(key))
                        _grid[key] = new List<Room>();
                    _grid[key].Add(room);
                }
            }
        }
        
        public List<Room> GetPotentialRooms(XYZ point)
        {
            int x = GetGridIndex(point.X);
            int y = GetGridIndex(point.Y);
            int key = GetCellKey(x, y);
            
            return _grid.ContainsKey(key) ? _grid[key] : new List<Room>();
        }
        
        public List<Room> GetPotentialRooms(BoundingBoxXYZ bb)
        {
            if (bb == null) return new List<Room>();
            
            HashSet<Room> rooms = new HashSet<Room>();
            int minX = GetGridIndex(bb.Min.X);
            int maxX = GetGridIndex(bb.Max.X);
            int minY = GetGridIndex(bb.Min.Y);
            int maxY = GetGridIndex(bb.Max.Y);
            
            for (int x = minX; x <= maxX; x++)
            {
                for (int y = minY; y <= maxY; y++)
                {
                    int key = GetCellKey(x, y);
                    if (_grid.ContainsKey(key))
                    {
                        foreach (Room room in _grid[key])
                            rooms.Add(room);
                    }
                }
            }
            
            return rooms.ToList();
        }
        
        private int GetGridIndex(double coord)
        {
            return (int)Math.Floor(coord / _cellSize);
        }
        
        private int GetCellKey(int x, int y)
        {
            return (x * 100000) + y; // Simple hash for 2D grid
        }
    }
    
    private class RoomData
    {
        public double LevelZ { get; set; }
        public double Height { get; set; }
        public double MinZ { get; set; }
        public double MaxZ { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
        public Room Room { get; set; }
        public Level Level { get; set; }
    }
    
    private class ElementData
    {
        public List<XYZ> TestPoints { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
        public Level Level { get; set; }
    }
    
    // Build room cache and spatial index
    private List<Room> BuildRoomCacheAndIndex(Document doc, SpatialIndex index)
    {
        FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
        List<Room> allRooms = roomCollector
            .OfClass(typeof(SpatialElement))
            .WhereElementIsNotElementType()
            .OfType<Room>()
            .Where(r => r != null && r.Area > 0) // Only placed rooms
            .ToList();
        
        // Build cache and index in one pass
        foreach (Room room in allRooms)
        {
            Level roomLevel = doc.GetElement(room.LevelId) as Level;
            if (roomLevel == null) continue;
            
            double roomLevelElevation = roomLevel.Elevation;
            double roomHeight = room.LookupParameter("Height")?.AsDouble() 
                ?? room.get_Parameter(BuiltInParameter.ROOM_HEIGHT)?.AsDouble()
                ?? UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Feet);
            
            RoomData data = new RoomData
            {
                Room = room,
                Level = roomLevel,
                LevelZ = roomLevelElevation,
                Height = roomHeight,
                MinZ = roomLevelElevation,
                MaxZ = roomLevelElevation + roomHeight,
                BoundingBox = room.get_BoundingBox(null)
            };
            
            _roomDataCache[room.Id] = data;
            index.AddRoom(room);
        }
        
        return allRooms;
    }
    
    // Pre-calculate group bounding boxes
    private void PreCalculateGroupBoundingBoxes(IList<Group> allGroups, Document doc)
    {
        _groupBoundingBoxCache.Clear();
        
        foreach (Group group in allGroups)
        {
            BoundingBoxXYZ bb = CalculateGroupBoundingBox(group, doc);
            if (bb != null)
            {
                _groupBoundingBoxCache[group.Id] = bb;
            }
        }
    }
    
    // Calculate group bounding box
    private BoundingBoxXYZ CalculateGroupBoundingBox(Group group, Document doc)
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
    
    // Build comprehensive room-to-group mapping with improved priority handling (from Copy command)
    private Dictionary<ElementId, Group> BuildRoomToGroupMapping(IList<Group> allGroups, Document doc)
    {
        Dictionary<ElementId, Group> roomToGroupMap = new Dictionary<ElementId, Group>();
        Dictionary<ElementId, List<Group>> roomDirectMembership = new Dictionary<ElementId, List<Group>>();
        Dictionary<ElementId, List<Group>> roomSpatialContainment = new Dictionary<ElementId, List<Group>>();
        
        // First pass: identify direct memberships
        foreach (Group group in allGroups)
        {
            foreach (ElementId memberId in group.GetMemberIds())
            {
                Element member = doc.GetElement(memberId);
                if (member is Room room && room.Area > 0)
                {
                    if (!roomDirectMembership.ContainsKey(memberId))
                    {
                        roomDirectMembership[memberId] = new List<Group>();
                    }
                    roomDirectMembership[memberId].Add(group);
                }
            }
        }
        
        // Second pass: identify spatial containments
        foreach (Group group in allGroups)
        {
            BoundingBoxXYZ groupBB = _groupBoundingBoxCache.ContainsKey(group.Id)
                ? _groupBoundingBoxCache[group.Id]
                : group.get_BoundingBox(null);
            
            if (groupBB == null) continue;
            
            // Expand bounding box slightly for tolerance
            double tolerance = 0.1;
            XYZ minExpanded = new XYZ(
                groupBB.Min.X - tolerance,
                groupBB.Min.Y - tolerance,
                groupBB.Min.Z - tolerance
            );
            XYZ maxExpanded = new XYZ(
                groupBB.Max.X + tolerance,
                groupBB.Max.Y + tolerance,
                groupBB.Max.Z + tolerance
            );
            
            // Check all rooms for spatial containment
            foreach (var kvp in _roomDataCache)
            {
                ElementId roomId = kvp.Key;
                RoomData roomData = kvp.Value;
                
                // Skip if this room has direct membership in ANY group
                if (roomDirectMembership.ContainsKey(roomId))
                    continue;
                
                // Check spatial containment
                if (IsRoomWithinBoundingBox(roomData, minExpanded, maxExpanded))
                {
                    if (!roomSpatialContainment.ContainsKey(roomId))
                    {
                        roomSpatialContainment[roomId] = new List<Group>();
                    }
                    roomSpatialContainment[roomId].Add(group);
                }
            }
        }
        
        // Build final mapping with improved priority
        // Direct memberships first (these are definitive)
        foreach (var kvp in roomDirectMembership)
        {
            ElementId roomId = kvp.Key;
            Room room = doc.GetElement(roomId) as Room;
            
            if (kvp.Value.Count == 1)
            {
                // Only one group claims direct membership - use it
                roomToGroupMap[roomId] = kvp.Value.First();
            }
            else
            {
                // Multiple groups claim direct membership - use boundary analysis
                Group bestGroup = SelectBestGroupForRoom(room, kvp.Value, doc);
                roomToGroupMap[roomId] = bestGroup;
                
                _diagnostics.AppendLine($"Room {room.Name ?? roomId.ToString()} has multiple direct memberships - selected {GetGroupTypeName(bestGroup, doc)} based on boundary analysis");
            }
        }
        
        // Add spatial containments (only for rooms without direct membership)
        foreach (var kvp in roomSpatialContainment)
        {
            if (!roomToGroupMap.ContainsKey(kvp.Key))
            {
                ElementId roomId = kvp.Key;
                Room room = doc.GetElement(roomId) as Room;
                
                if (kvp.Value.Count == 1)
                {
                    // Only one group spatially contains this room
                    roomToGroupMap[roomId] = kvp.Value.First();
                }
                else
                {
                    // Multiple groups spatially contain this room - use boundary analysis
                    Group bestGroup = SelectBestGroupForRoom(room, kvp.Value, doc);
                    roomToGroupMap[roomId] = bestGroup;
                    
                    _diagnostics.AppendLine($"Room {room.Name ?? roomId.ToString()} is spatially contained by multiple groups - selected {GetGroupTypeName(bestGroup, doc)} based on boundary analysis");
                }
            }
        }
        
        return roomToGroupMap;
    }
    
    // Select the best group for a room based on which group has more walls forming the room boundary
    private Group SelectBestGroupForRoom(Room room, List<Group> candidateGroups, Document doc)
    {
        if (candidateGroups.Count == 1)
            return candidateGroups.First();
        
        // Get room boundary segments
        SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
        options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
        
        IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);
        
        if (boundaries == null || boundaries.Count == 0)
        {
            // No boundaries - fall back to smallest bounding box
            return SelectGroupBySmallestBoundingBox(candidateGroups);
        }
        
        // Count how many boundary elements each group contains
        Dictionary<Group, int> groupBoundaryCount = new Dictionary<Group, int>();
        Dictionary<Group, double> groupBoundaryLength = new Dictionary<Group, double>();
        
        foreach (Group group in candidateGroups)
        {
            int count = 0;
            double totalLength = 0.0;
            HashSet<ElementId> memberIds = new HashSet<ElementId>(group.GetMemberIds());
            
            // Check each boundary segment
            foreach (IList<BoundarySegment> loop in boundaries)
            {
                foreach (BoundarySegment segment in loop)
                {
                    if (memberIds.Contains(segment.ElementId))
                    {
                        count++;
                        
                        // Get segment length
                        Curve segmentCurve = segment.GetCurve();
                        if (segmentCurve != null)
                        {
                            totalLength += segmentCurve.Length;
                        }
                    }
                }
            }
            
            groupBoundaryCount[group] = count;
            groupBoundaryLength[group] = totalLength;
        }
        
        // Find the group with the most boundary elements
        int maxCount = groupBoundaryCount.Values.Max();
        List<Group> groupsWithMaxCount = groupBoundaryCount
            .Where(kvp => kvp.Value == maxCount)
            .Select(kvp => kvp.Key)
            .ToList();
        
        if (groupsWithMaxCount.Count == 1)
        {
            // Clear winner
            Group winner = groupsWithMaxCount.First();
            _diagnostics.AppendLine($"  Selected {GetGroupTypeName(winner, doc)} with {maxCount} boundary segments");
            return winner;
        }
        
        // If there's still a tie, use total boundary length
        if (maxCount > 0)
        {
            Group bestByLength = groupsWithMaxCount
                .OrderByDescending(g => groupBoundaryLength[g])
                .First();
            
            _diagnostics.AppendLine($"  Selected {GetGroupTypeName(bestByLength, doc)} with {maxCount} boundary segments and {groupBoundaryLength[bestByLength]:F2} ft total length");
            return bestByLength;
        }
        
        // If no groups have boundary elements, check for spatial relationship
        // Check which group has more walls near the room
        Dictionary<Group, int> groupWallCount = new Dictionary<Group, int>();
        
        foreach (Group group in candidateGroups)
        {
            int wallCount = 0;
            ICollection<ElementId> memberIds = group.GetMemberIds();
            
            foreach (ElementId memberId in memberIds)
            {
                Element member = doc.GetElement(memberId);
                if (member is Wall wall)
                {
                    // Check if wall is near the room
                    LocationCurve wallLocation = wall.Location as LocationCurve;
                    if (wallLocation != null)
                    {
                        Curve wallCurve = wallLocation.Curve;
                        XYZ wallMidpoint = wallCurve.Evaluate(0.5, true);
                        
                        // Check if wall midpoint is near room boundary
                        if (IsPointNearRoomBoundary(wallMidpoint, room, 2.0)) // 2 feet tolerance
                        {
                            wallCount++;
                        }
                    }
                }
            }
            
            groupWallCount[group] = wallCount;
        }
        
        // Select group with most walls near the room
        Group bestByWalls = groupWallCount
            .OrderByDescending(kvp => kvp.Value)
            .ThenBy(kvp => GetGroupBoundingBoxVolume(kvp.Key)) // Smaller volume as tiebreaker
            .First().Key;
        
        _diagnostics.AppendLine($"  Selected {GetGroupTypeName(bestByWalls, doc)} with {groupWallCount[bestByWalls]} walls near room");
        
        return bestByWalls;
    }
    
    // Check if a point is near a room boundary
    private bool IsPointNearRoomBoundary(XYZ point, Room room, double tolerance)
    {
        SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
        IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);
        
        if (boundaries == null) return false;
        
        foreach (IList<BoundarySegment> loop in boundaries)
        {
            foreach (BoundarySegment segment in loop)
            {
                Curve segmentCurve = segment.GetCurve();
                if (segmentCurve != null)
                {
                    // Find closest point on curve to the given point
                    IntersectionResult result = segmentCurve.Project(point);
                    if (result != null)
                    {
                        double distance = result.Distance;
                        if (distance <= tolerance)
                        {
                            return true;
                        }
                    }
                }
            }
        }
        
        return false;
    }
    
    // Fallback: Select group by smallest bounding box
    private Group SelectGroupBySmallestBoundingBox(List<Group> groups)
    {
        Group bestGroup = null;
        double smallestVolume = double.MaxValue;
        
        foreach (Group group in groups)
        {
            double volume = GetGroupBoundingBoxVolume(group);
            if (volume < smallestVolume)
            {
                smallestVolume = volume;
                bestGroup = group;
            }
        }
        
        return bestGroup ?? groups.First();
    }
    
    // Get group bounding box volume
    private double GetGroupBoundingBoxVolume(Group group)
    {
        BoundingBoxXYZ bb = _groupBoundingBoxCache.ContainsKey(group.Id)
            ? _groupBoundingBoxCache[group.Id]
            : group.get_BoundingBox(null);
        
        if (bb == null) return double.MaxValue;
        
        XYZ size = bb.Max - bb.Min;
        return size.X * size.Y * size.Z;
    }
    
    // Get group type name for diagnostics
    private string GetGroupTypeName(Group group, Document doc)
    {
        if (group == null) return "null";
        GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
        return gt?.Name ?? "Unknown";
    }
    
    // Check if room is within bounding box
    private bool IsRoomWithinBoundingBox(RoomData roomData, XYZ bbMin, XYZ bbMax)
    {
        BoundingBoxXYZ roomBB = roomData.BoundingBox;
        if (roomBB == null) return false;
        
        return roomBB.Min.X >= bbMin.X && roomBB.Min.Y >= bbMin.Y && roomBB.Min.Z >= bbMin.Z &&
               roomBB.Max.X <= bbMax.X && roomBB.Max.Y <= bbMax.Y && roomBB.Max.Z <= bbMax.Z;
    }
    
    // Find which rooms contain the selected elements - OPTIMIZED with spatial index
    private Dictionary<ElementId, List<Room>> FindRoomsContainingElementsOptimized(
        List<Element> elements, 
        Dictionary<ElementId, ElementData> elementDataCache,
        SpatialIndex roomIndex)
    {
        Dictionary<ElementId, List<Room>> result = new Dictionary<ElementId, List<Room>>();
        
        // Initialize map
        foreach (Element elem in elements)
        {
            result[elem.Id] = new List<Room>();
        }
        
        // Check each element against potentially overlapping rooms only
        foreach (Element elem in elements)
        {
            ElementData elemData = elementDataCache[elem.Id];
            
            // Get potential rooms from spatial index
            List<Room> potentialRooms;
            if (elemData.BoundingBox != null)
            {
                potentialRooms = roomIndex.GetPotentialRooms(elemData.BoundingBox);
            }
            else if (elemData.TestPoints.Count > 0)
            {
                // For point-based elements
                HashSet<Room> roomSet = new HashSet<Room>();
                foreach (XYZ point in elemData.TestPoints)
                {
                    foreach (Room room in roomIndex.GetPotentialRooms(point))
                    {
                        roomSet.Add(room);
                    }
                }
                potentialRooms = roomSet.ToList();
            }
            else
            {
                continue;
            }
            
            // Only check against potential rooms
            foreach (Room room in potentialRooms)
            {
                if (IsElementInRoomOptimized(elem, room, elemData, _roomDataCache[room.Id]))
                {
                    result[elem.Id].Add(room);
                }
            }
            
            // Diagnostic: Report if element is in multiple rooms
            if (result[elem.Id].Count > 1)
            {
                string elemLevelName = elemData.Level?.Name ?? "No Level";
                _diagnostics.AppendLine($"\nElement {elem.Id.IntegerValue} (Level: {elemLevelName}) is in {result[elem.Id].Count} rooms:");
                foreach (Room room in result[elem.Id])
                {
                    RoomData roomData = _roomDataCache[room.Id];
                    string roomLevelName = roomData.Level?.Name ?? "No Level";
                    _diagnostics.AppendLine($"  - {room.Name ?? "Unnamed"} (Level: {roomLevelName})");
                }
            }
        }
        
        return result;
    }
    
    // Optimized room containment check
    private bool IsElementInRoomOptimized(Element element, Room room, ElementData elemData, RoomData roomData)
    {
        // Quick bounding box check first
        if (elemData.BoundingBox != null && roomData.BoundingBox != null)
        {
            // If bounding boxes don't overlap in Z, element can't be in room
            if (elemData.BoundingBox.Max.Z < roomData.MinZ - 0.5 || 
                elemData.BoundingBox.Min.Z > roomData.MaxZ + 0.5)
            {
                return false;
            }
            
            // If bounding boxes don't overlap in XY, element can't be in room
            if (elemData.BoundingBox.Max.X < roomData.BoundingBox.Min.X || 
                elemData.BoundingBox.Min.X > roomData.BoundingBox.Max.X ||
                elemData.BoundingBox.Max.Y < roomData.BoundingBox.Min.Y || 
                elemData.BoundingBox.Min.Y > roomData.BoundingBox.Max.Y)
            {
                return false;
            }
        }
        
        // For point-based elements (like sprinklers), use only the location point
        LocationPoint locPoint = element.Location as LocationPoint;
        if (locPoint != null)
        {
            XYZ point = locPoint.Point;
            
            // Check if point is within room's vertical bounds
            double tolerance = 0.5; // 0.5 feet tolerance for ceiling-mounted elements
            if (point.Z < roomData.MinZ - tolerance || point.Z > roomData.MaxZ + tolerance)
            {
                return false;
            }
            
            // Check XY containment at optimal test heights
            XYZ testPoint1 = new XYZ(point.X, point.Y, roomData.LevelZ + 1.0);
            XYZ testPoint2 = new XYZ(point.X, point.Y, roomData.LevelZ + roomData.Height * 0.5);
            
            try
            {
                if (room.IsPointInRoom(testPoint1) || room.IsPointInRoom(testPoint2))
                {
                    return true;
                }
            }
            catch
            {
                try
                {
                    return room.IsPointInRoom(point);
                }
                catch
                {
                    return false;
                }
            }
            
            return false;
        }
        
        // For linear and other elements, check fewer test points
        // Only check points that are within room's Z bounds
        foreach (XYZ point in elemData.TestPoints)
        {
            if (point.Z < roomData.MinZ - 0.1 || point.Z > roomData.MaxZ + 0.1)
            {
                continue;
            }
            
            XYZ testPoint = new XYZ(point.X, point.Y, roomData.LevelZ + roomData.Height * 0.5);
            
            try
            {
                if (room.IsPointInRoom(testPoint))
                {
                    return true;
                }
            }
            catch
            {
                continue;
            }
        }
        
        return false;
    }
    
    // Get test points for an element - OPTIMIZED to generate fewer points
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
            Curve curve = locCurve.Curve;
            testPoints.Add(curve.GetEndPoint(0));
            testPoints.Add(curve.GetEndPoint(1));
            testPoints.Add(curve.Evaluate(0.5, true)); // Midpoint
            
            // For longer curves, add a couple more points
            double length = curve.Length;
            if (length > 10.0) // Only for curves longer than 10 feet
            {
                testPoints.Add(curve.Evaluate(0.25, true));
                testPoints.Add(curve.Evaluate(0.75, true));
            }
        }
        else
        {
            // For other elements, use bounding box center and fewer corners
            BoundingBoxXYZ bb = element.get_BoundingBox(null);
            if (bb != null)
            {
                // Center
                testPoints.Add((bb.Min + bb.Max) * 0.5);
                
                // Just 4 corners (alternating high/low) instead of all 8
                testPoints.Add(new XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z));
                testPoints.Add(new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z));
                testPoints.Add(new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z));
                testPoints.Add(new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z));
                
                // Add 2 face centers for better coverage on large elements
                double sizeX = bb.Max.X - bb.Min.X;
                double sizeY = bb.Max.Y - bb.Min.Y;
                double sizeZ = bb.Max.Z - bb.Min.Z;
                
                // Only add face centers for larger elements
                if (Math.Max(Math.Max(sizeX, sizeY), sizeZ) > 5.0)
                {
                    testPoints.Add(new XYZ((bb.Min.X + bb.Max.X) * 0.5, bb.Min.Y, (bb.Min.Z + bb.Max.Z) * 0.5));
                    testPoints.Add(new XYZ(bb.Min.X, (bb.Min.Y + bb.Max.Y) * 0.5, (bb.Min.Z + bb.Max.Z) * 0.5));
                }
            }
        }
        
        return testPoints;
    }
}
