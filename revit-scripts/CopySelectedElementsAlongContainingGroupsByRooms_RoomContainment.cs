using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
    // Build comprehensive room-to-group mapping with improved priority handling
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
                
                diagnosticLog.AppendLine($"Room {room.Name ?? roomId.ToString()} has multiple direct memberships - selected {GetGroupTypeName(bestGroup, doc)} based on boundary analysis");
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
                    
                    diagnosticLog.AppendLine($"Room {room.Name ?? roomId.ToString()} is spatially contained by multiple groups - selected {GetGroupTypeName(bestGroup, doc)} based on boundary analysis");
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

        // Collect all boundary element IDs
        HashSet<ElementId> boundaryElementIds = new HashSet<ElementId>();
        foreach (IList<BoundarySegment> loop in boundaries)
        {
            foreach (BoundarySegment segment in loop)
            {
                Element boundaryElement = doc.GetElement(segment.ElementId);
                if (boundaryElement != null)
                {
                    boundaryElementIds.Add(segment.ElementId);
                    
                    // If it's a wall, also check if it's a curtain wall with panels
                    if (boundaryElement is Wall wall)
                    {
                        // Get curtain grid if this is a curtain wall
                        CurtainGrid grid = wall.CurtainGrid;
                        if (grid != null)
                        {
                            // Add panels as boundary elements
                            ICollection<ElementId> panelIds = grid.GetPanelIds();
                            foreach (ElementId panelId in panelIds)
                            {
                                boundaryElementIds.Add(panelId);
                            }
                        }
                    }
                }
            }
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
            diagnosticLog.AppendLine($"  Selected {GetGroupTypeName(winner, doc)} with {maxCount} boundary segments");
            return winner;
        }

        // If there's still a tie, use total boundary length
        if (maxCount > 0)
        {
            Group bestByLength = groupsWithMaxCount
                .OrderByDescending(g => groupBoundaryLength[g])
                .First();
            
            diagnosticLog.AppendLine($"  Selected {GetGroupTypeName(bestByLength, doc)} with {maxCount} boundary segments and {groupBoundaryLength[bestByLength]:F2} ft total length");
            return bestByLength;
        }

        // If no groups have boundary elements (rare), check for spatial relationship
        // This can happen with rooms that are defined by room separation lines or other elements
        
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

        diagnosticLog.AppendLine($"  Selected {GetGroupTypeName(bestByWalls, doc)} with {groupWallCount[bestByWalls]} walls near room");
        
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

    // Get elements contained in group's rooms using filtered mapping
    private List<Element> GetElementsContainedInGroupRoomsFiltered(
        Group group,
        List<Element> candidateElements,
        Document doc,
        Dictionary<ElementId, Group> roomToGroupMap)
    {
        List<Element> containedElements = new List<Element>();

        // Get all rooms that map to this group
        List<Room> groupRooms = new List<Room>();
        foreach (var kvp in roomToGroupMap)
        {
            if (kvp.Value.Id == group.Id)
            {
                Room room = doc.GetElement(kvp.Key) as Room;
                if (room != null && room.Area > 0)
                {
                    groupRooms.Add(room);
                }
            }
        }

        if (groupRooms.Count == 0)
        {
            return containedElements;
        }

        // Get the overall height range of the group
        ICollection<ElementId> memberIds = group.GetMemberIds();
        double groupMinZ = double.MaxValue;
        double groupMaxZ = double.MinValue;

        foreach (ElementId id in memberIds)
        {
            Element member = doc.GetElement(id);
            BoundingBoxXYZ bb = member.get_BoundingBox(null);
            if (bb != null)
            {
                groupMinZ = Math.Min(groupMinZ, bb.Min.Z);
                groupMaxZ = Math.Max(groupMaxZ, bb.Max.Z);
            }
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

    // Check if room is within bounding box
    private bool IsRoomWithinBoundingBox(RoomData roomData, XYZ bbMin, XYZ bbMax)
    {
        BoundingBoxXYZ roomBB = roomData.BoundingBox;
        if (roomBB == null) return false;

        return roomBB.Min.X >= bbMin.X && roomBB.Min.Y >= bbMin.Y && roomBB.Min.Z >= bbMin.Z &&
               roomBB.Max.X <= bbMax.X && roomBB.Max.Y <= bbMax.Y && roomBB.Max.Z <= bbMax.Z;
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
