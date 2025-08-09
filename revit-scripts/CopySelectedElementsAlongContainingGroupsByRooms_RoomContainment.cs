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
                if (bestGroup != null)
                {
                    roomToGroupMap[roomId] = bestGroup;
                    diagnosticLog.AppendLine($"Room {room.Name ?? roomId.ToString()} has multiple direct memberships - selected {GetGroupTypeName(bestGroup, doc)} based on boundary analysis");
                }
                else
                {
                    diagnosticLog.AppendLine($"Room {room.Name ?? roomId.ToString()} has multiple direct memberships but no valid group assignment");
                }
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
                    // Only one group spatially contains this room - but verify it actually has walls near it
                    Group group = kvp.Value.First();
                    if (DoesGroupHaveWallsNearRoom(group, room, doc))
                    {
                        roomToGroupMap[roomId] = group;
                    }
                    else
                    {
                        diagnosticLog.AppendLine($"Room {room.Name ?? roomId.ToString()} is spatially contained by {GetGroupTypeName(group, doc)} but has no nearby walls - NOT assigned");
                    }
                }
                else
                {
                    // Multiple groups spatially contain this room - use boundary analysis
                    Group bestGroup = SelectBestGroupForRoom(room, kvp.Value, doc);
                    if (bestGroup != null)
                    {
                        roomToGroupMap[roomId] = bestGroup;
                        diagnosticLog.AppendLine($"Room {room.Name ?? roomId.ToString()} is spatially contained by multiple groups - selected {GetGroupTypeName(bestGroup, doc)} based on boundary analysis");
                    }
                    else
                    {
                        diagnosticLog.AppendLine($"Room {room.Name ?? roomId.ToString()} is spatially contained by multiple groups but no valid group assignment");
                    }
                }
            }
        }

        return roomToGroupMap;
    }

    // Check if a group has walls near a room (more strict version)
    private bool DoesGroupHaveWallsNearRoom(Group group, Room room, Document doc)
    {
        // Get room level for proper Z-coordinate checking
        Level roomLevel = doc.GetElement(room.LevelId) as Level;
        if (roomLevel == null) return false;
        
        double roomLevelZ = roomLevel.Elevation;
        double tolerance = 2.0; // 2 feet horizontal tolerance
        double verticalTolerance = 5.0; // 5 feet vertical tolerance
        
        // Get room bounding box
        BoundingBoxXYZ roomBB = room.get_BoundingBox(null);
        if (roomBB == null) return false;

        // Count walls from group that are near the room
        int nearbyWallCount = 0;
        ICollection<ElementId> memberIds = group.GetMemberIds();
        
        foreach (ElementId memberId in memberIds)
        {
            Element member = doc.GetElement(memberId);
            if (member is Wall wall)
            {
                LocationCurve wallLocation = wall.Location as LocationCurve;
                if (wallLocation != null)
                {
                    Curve wallCurve = wallLocation.Curve;
                    XYZ wallStart = wallCurve.GetEndPoint(0);
                    XYZ wallEnd = wallCurve.GetEndPoint(1);
                    XYZ wallMid = wallCurve.Evaluate(0.5, true);
                    
                    // Check if wall is at approximately the same level as the room
                    double wallZ = (wallStart.Z + wallEnd.Z) / 2.0;
                    if (Math.Abs(wallZ - roomLevelZ) > verticalTolerance)
                        continue;
                    
                    // Check horizontal proximity
                    bool nearRoom = false;
                    
                    // Check if wall endpoints or midpoint are near room bounding box
                    XYZ[] checkPoints = { wallStart, wallEnd, wallMid };
                    foreach (XYZ point in checkPoints)
                    {
                        if (point.X >= roomBB.Min.X - tolerance && point.X <= roomBB.Max.X + tolerance &&
                            point.Y >= roomBB.Min.Y - tolerance && point.Y <= roomBB.Max.Y + tolerance)
                        {
                            nearRoom = true;
                            break;
                        }
                    }
                    
                    if (nearRoom)
                        nearbyWallCount++;
                }
            }
        }
        
        // Require at least 2 walls to consider the group as defining the room
        return nearbyWallCount >= 2;
    }

    // Select the best group for a room based on which group has more walls forming the room boundary
    private Group SelectBestGroupForRoom(Room room, List<Group> candidateGroups, Document doc)
    {
        if (candidateGroups.Count == 1)
        {
            Group singleGroup = candidateGroups.First();
            // Even for single candidate, verify it has walls near the room
            if (!DoesGroupHaveWallsNearRoom(singleGroup, room, doc))
            {
                diagnosticLog.AppendLine($"  Single candidate group {GetGroupTypeName(singleGroup, doc)} has no walls near room {room.Name ?? room.Id.ToString()} - excluding");
                return null;
            }
            return singleGroup;
        }

        // Get room boundary segments
        SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
        options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;

        IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);

        if (boundaries == null || boundaries.Count == 0)
        {
            // No boundaries - this room might be unbound or defined by room separation lines
            diagnosticLog.AppendLine($"  Room {room.Name ?? room.Id.ToString()} has no boundaries - checking for nearby walls");
            
            // Filter to only groups that actually have walls near this room
            List<Group> validGroups = new List<Group>();
            foreach (Group group in candidateGroups)
            {
                if (DoesGroupHaveWallsNearRoom(group, room, doc))
                {
                    validGroups.Add(group);
                }
            }
            
            if (validGroups.Count == 0)
            {
                diagnosticLog.AppendLine($"    No groups have walls near room {room.Name ?? room.Id.ToString()} - excluding from group assignment");
                return null;
            }
            else if (validGroups.Count == 1)
            {
                Group selected = validGroups.First();
                diagnosticLog.AppendLine($"    Selected {GetGroupTypeName(selected, doc)} as only group with walls near room");
                return selected;
            }
            else
            {
                // Multiple groups with walls - use the one with smallest bounding box
                Group best = SelectGroupBySmallestBoundingBox(validGroups);
                diagnosticLog.AppendLine($"    Selected {GetGroupTypeName(best, doc)} based on smallest bounding box");
                return best;
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
        
        // If no group has any boundary segments, check for nearby walls
        if (maxCount == 0)
        {
            diagnosticLog.AppendLine($"    No groups have boundary segments for room {room.Name ?? room.Id.ToString()}");
            
            // Filter to only groups with walls near the room
            List<Group> validGroups = new List<Group>();
            foreach (Group group in candidateGroups)
            {
                if (DoesGroupHaveWallsNearRoom(group, room, doc))
                {
                    validGroups.Add(group);
                }
            }
            
            if (validGroups.Count == 0)
            {
                diagnosticLog.AppendLine($"    No groups have walls near room - excluding from assignment");
                return null;
            }
            else
            {
                Group best = SelectGroupBySmallestBoundingBox(validGroups);
                diagnosticLog.AppendLine($"    Selected {GetGroupTypeName(best, doc)} based on smallest bounding box among groups with nearby walls");
                return best;
            }
        }
        
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
        Group bestByLength = groupsWithMaxCount
            .OrderByDescending(g => groupBoundaryLength[g])
            .First();

        diagnosticLog.AppendLine($"  Selected {GetGroupTypeName(bestByLength, doc)} with {maxCount} boundary segments and {groupBoundaryLength[bestByLength]:F2} ft total length");
        return bestByLength;
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

        // Get the group's location for better vertical filtering
        LocationPoint groupLocation = group.Location as LocationPoint;
        double groupZ = groupLocation?.Point.Z ?? 0;
        
        // Use tighter vertical bounds based on typical floor height
        double typicalFloorHeight = 15.0; // 15 feet typical floor-to-floor height
        double groupMinZ = groupZ - 2.0; // 2 feet below group origin
        double groupMaxZ = groupZ + typicalFloorHeight; // One floor height above

        // Check each candidate element against all rooms
        foreach (Element elem in candidateElements)
        {
            // Get cached test points
            List<XYZ> testPoints = _elementTestPointsCache[elem.Id];

            foreach (Room room in groupRooms)
            {
                if (IsElementInRoomWithStrictZCheck(elem, room, testPoints, groupMinZ, groupMaxZ, groupZ))
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

    // More strict version of element-in-room check with better Z filtering
    private bool IsElementInRoomWithStrictZCheck(Element element, Room room, List<XYZ> testPoints, 
                                                 double groupMinZ, double groupMaxZ, double groupOriginZ)
    {
        // Skip unplaced or unenclosed rooms
        if (room.Area <= 0)
            return false;

        // Get cached room data
        RoomData roomData = _roomDataCache[room.Id];
        
        // Get room level
        Level roomLevel = room.Level;
        if (roomLevel == null) return false;
        
        double roomLevelZ = roomLevel.Elevation;
        
        // Check if room is at approximately the same level as the group
        // Allow for some vertical tolerance but not entire floors
        double levelDifference = Math.Abs(roomLevelZ - groupOriginZ);
        if (levelDifference > 5.0) // More than 5 feet difference in levels
        {
            // This room is on a different floor than the group
            return false;
        }

        // Check if any test point is within the room
        foreach (XYZ point in testPoints)
        {
            // Strict Z check - element must be within the group's vertical range
            if (point.Z < groupMinZ || point.Z > groupMaxZ)
                continue;
            
            // Additional check: element should be close to room level
            double elementToRoomLevelDist = Math.Abs(point.Z - roomLevelZ);
            if (elementToRoomLevelDist > 10.0) // Element more than 10 feet from room level
                continue;

            // For ALL rooms (bounded or unbound), check XY containment at room level
            double testZ = roomLevelZ + 1.0; // 1 foot above room level
            if (!roomData.IsUnbound && roomData.BoundingBox != null)
            {
                // For bounded rooms, use a point within the room's height range
                testZ = (roomData.BoundingBox.Min.Z + roomData.BoundingBox.Max.Z) * 0.5;
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
