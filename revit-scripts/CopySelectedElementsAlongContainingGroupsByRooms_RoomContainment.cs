using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
    // Configurable validation thresholds
    private double ROOM_POSITION_TOLERANCE = 5.0; // feet
    private double ROOM_POSITION_TOLERANCE_RELAXED = 10.0; // feet for relaxed mode
    private double ROOM_AREA_TOLERANCE = 0.3; // 30%
    private double ROOM_AREA_TOLERANCE_RELAXED = 0.5; // 50% for relaxed mode
    private double ROOM_VALIDATION_THRESHOLD = 0.3; // Require 30% of instances instead of 50%
    private int ROOM_VALIDATION_MIN_MATCHES = 1; // At least 1 other instance must have the room

    // Structure to track room mapping with metadata
    private class RoomMappingInfo
    {
        public Group MappedGroup { get; set; }
        public bool IsDirectMember { get; set; }
        public XYZ RelativePosition { get; set; }
        public double Area { get; set; }
        public string RoomName { get; set; }
        public string RoomNumber { get; set; }
        public ElementId RoomId { get; set; }
        public XYZ RoomLocation { get; set; }
    }

    // Cache for room locations indexed by spatial grid
    private Dictionary<int, List<RoomMappingInfo>> _roomSpatialIndex = new Dictionary<int, List<RoomMappingInfo>>();
    private const double ROOM_SPATIAL_GRID_SIZE = 10.0; // 10 feet grid for room lookups

    // Build comprehensive room-to-group mapping with validation for spatial containment
    private Dictionary<ElementId, Group> BuildRoomToGroupMapping(IList<Group> relevantGroups, Document doc)
    {
        Dictionary<ElementId, RoomMappingInfo> roomMappingInfo = new Dictionary<ElementId, RoomMappingInfo>();
        Dictionary<ElementId, List<Group>> roomDirectMembership = new Dictionary<ElementId, List<Group>>();
        Dictionary<ElementId, List<Group>> roomSpatialContainment = new Dictionary<ElementId, List<Group>>();

        // Build spatial index for all rooms first
        BuildRoomSpatialIndex(doc);

        // First pass: identify direct memberships
        foreach (Group group in relevantGroups)
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
                    roomsAsDirectMembers++;
                }
            }
        }

        // Second pass: identify spatial containments
        foreach (Group group in relevantGroups)
        {
            if (!_groupBoundingBoxCache.ContainsKey(group.Id)) continue;

            BoundingBoxXYZ groupBB = _groupBoundingBoxCache[group.Id];
            if (groupBB == null) continue;

            // Use spatial index to find potentially contained rooms
            var potentialRooms = GetRoomsInBoundingBox(groupBB);

            foreach (var roomInfo in potentialRooms)
            {
                ElementId roomId = roomInfo.RoomId;

                // Skip if this room has direct membership in ANY group
                if (roomDirectMembership.ContainsKey(roomId))
                    continue;

                // Check if room is actually within the group's bounding box
                RoomData roomData = _roomDataCache[roomId];
                if (roomData != null && IsRoomWithinBoundingBox(roomData, groupBB.Min, groupBB.Max))
                {
                    // Additional check: does the room have significant boundary overlap with group members?
                    Room room = doc.GetElement(roomId) as Room;
                    if (room != null && ShouldForceIncludeRoom(room, group, doc))
                    {
                        if (!roomSpatialContainment.ContainsKey(roomId))
                        {
                            roomSpatialContainment[roomId] = new List<Group>();
                        }
                        roomSpatialContainment[roomId].Add(group);
                    }
                }
            }
        }

        // Build mapping with direct memberships
        Dictionary<ElementId, Group> roomToGroupMap = new Dictionary<ElementId, Group>();

        foreach (var kvp in roomDirectMembership)
        {
            ElementId roomId = kvp.Key;
            Room room = doc.GetElement(roomId) as Room;

            if (kvp.Value.Count == 1)
            {
                Group group = kvp.Value.First();
                roomToGroupMap[roomId] = group;
            }
            else
            {
                Group bestGroup = SelectBestGroupForRoom(room, kvp.Value, doc);
                roomToGroupMap[roomId] = bestGroup;
            }
        }

        // Validate spatially contained rooms
        FilteredElementCollector allGroupsCollector = new FilteredElementCollector(doc);
        IList<Group> allGroups = allGroupsCollector
            .OfClass(typeof(Group))
            .WhereElementIsNotElementType()
            .Cast<Group>()
            .ToList();

        Dictionary<ElementId, List<Group>> validatedSpatialContainments = ValidateSpatialContainmentsEnhanced(
            roomSpatialContainment, allGroups, doc);

        // Add validated spatial containments
        foreach (var kvp in validatedSpatialContainments)
        {
            if (!roomToGroupMap.ContainsKey(kvp.Key))
            {
                ElementId roomId = kvp.Key;
                Room room = doc.GetElement(roomId) as Room;

                if (kvp.Value.Count == 1)
                {
                    Group group = kvp.Value.First();
                    roomToGroupMap[roomId] = group;
                }
                else
                {
                    Group bestGroup = SelectBestGroupForRoom(room, kvp.Value, doc);
                    roomToGroupMap[roomId] = bestGroup;
                }
            }
        }

        return roomToGroupMap;
    }

    // Build spatial index for all rooms
    private void BuildRoomSpatialIndex(Document doc)
    {
        _roomSpatialIndex.Clear();

        foreach (var kvp in _roomDataCache)
        {
            ElementId roomId = kvp.Key;
            RoomData roomData = kvp.Value;
            Room room = doc.GetElement(roomId) as Room;

            if (room == null || room.Area <= 0) continue;

            XYZ roomLocation = (room.Location as LocationPoint)?.Point ?? GetRoomCentroid(room);

            var roomInfo = new RoomMappingInfo
            {
                RoomId = roomId,
                RoomLocation = roomLocation,
                Area = room.Area,
                RoomName = room.Name,
                RoomNumber = room.Number
            };

            // Add to spatial index
            int gridX = (int)Math.Floor(roomLocation.X / ROOM_SPATIAL_GRID_SIZE);
            int gridY = (int)Math.Floor(roomLocation.Y / ROOM_SPATIAL_GRID_SIZE);
            int cellKey = GetRoomSpatialCellKey(gridX, gridY);

            if (!_roomSpatialIndex.ContainsKey(cellKey))
            {
                _roomSpatialIndex[cellKey] = new List<RoomMappingInfo>();
            }
            _roomSpatialIndex[cellKey].Add(roomInfo);
        }
    }

    // Get rooms in bounding box using spatial index
    private List<RoomMappingInfo> GetRoomsInBoundingBox(BoundingBoxXYZ bb)
    {
        List<RoomMappingInfo> rooms = new List<RoomMappingInfo>();

        int minGridX = (int)Math.Floor(bb.Min.X / ROOM_SPATIAL_GRID_SIZE);
        int maxGridX = (int)Math.Floor(bb.Max.X / ROOM_SPATIAL_GRID_SIZE);
        int minGridY = (int)Math.Floor(bb.Min.Y / ROOM_SPATIAL_GRID_SIZE);
        int maxGridY = (int)Math.Floor(bb.Max.Y / ROOM_SPATIAL_GRID_SIZE);

        for (int x = minGridX; x <= maxGridX; x++)
        {
            for (int y = minGridY; y <= maxGridY; y++)
            {
                int cellKey = GetRoomSpatialCellKey(x, y);
                if (_roomSpatialIndex.ContainsKey(cellKey))
                {
                    rooms.AddRange(_roomSpatialIndex[cellKey]);
                }
            }
        }

        return rooms;
    }

    // Get spatial cell key for room grid
    private int GetRoomSpatialCellKey(int gridX, int gridY)
    {
        return (gridX + 10000) * 20000 + (gridY + 10000);
    }

    // Enhanced validation result with detailed information
    private class DetailedValidationResult
    {
        public bool IsValid { get; set; }
        public int InstancesChecked { get; set; }
        public int InstancesWithRoom { get; set; }
        public double PercentageMatch => InstancesChecked > 0 ? (double)InstancesWithRoom / InstancesChecked * 100 : 0;
        public string PrimaryFailureReason { get; set; }
        public List<string> MatchDetails { get; set; } = new List<string>();
        public List<string> MismatchDetails { get; set; } = new List<string>();
    }

    // Enhanced validation with better diagnostics and relaxed thresholds
    private Dictionary<ElementId, List<Group>> ValidateSpatialContainmentsEnhanced(
        Dictionary<ElementId, List<Group>> roomSpatialContainment,
        IList<Group> allGroups,
        Document doc)
    {
        Dictionary<ElementId, List<Group>> validatedContainments = new Dictionary<ElementId, List<Group>>();

        // Group all groups by their type
        Dictionary<ElementId, List<Group>> groupsByType = new Dictionary<ElementId, List<Group>>();
        foreach (Group group in allGroups)
        {
            ElementId typeId = group.GetTypeId();
            if (!groupsByType.ContainsKey(typeId))
            {
                groupsByType[typeId] = new List<Group>();
            }
            groupsByType[typeId].Add(group);
        }

        foreach (var kvp in roomSpatialContainment)
        {
            ElementId roomId = kvp.Key;
            List<Group> candidateGroups = kvp.Value;
            Room room = doc.GetElement(roomId) as Room;

            if (room == null) continue;

            List<Group> validatedGroups = new List<Group>();

            foreach (Group candidateGroup in candidateGroups)
            {
                ElementId groupTypeId = candidateGroup.GetTypeId();
                List<Group> sameTypeGroups = groupsByType[groupTypeId];

                // Skip if there's only one instance of this group type
                if (sameTypeGroups.Count < 2)
                {
                    TrackStat("Single Instance Group Type");
                    continue;
                }

                // Check if similar rooms exist in other instances
                var validationResult = ValidateRoomInOtherInstancesWithDetails(
                    room, candidateGroup, sameTypeGroups, doc);

                if (validationResult.IsValid)
                {
                    validatedGroups.Add(candidateGroup);
                    roomsValidatedBySimilarity++;
                    LogRoomValidationSuccess(room, candidateGroup, validationResult, doc);
                }
                else
                {
                    roomsInvalidatedByDissimilarity++;
                    LogRoomValidationFailure(room, candidateGroup, sameTypeGroups, validationResult, doc);
                }
            }

            if (validatedGroups.Count > 0)
            {
                validatedContainments[roomId] = validatedGroups;
            }
        }

        return validatedContainments;
    }

    // Enhanced room validation with relaxed thresholds
    private DetailedValidationResult ValidateRoomInOtherInstancesWithDetails(
        Room testRoom, Group candidateGroup, List<Group> sameTypeGroups, Document doc)
    {
        var result = new DetailedValidationResult();

        // Get test room properties
        XYZ candidateGroupOrigin = (candidateGroup.Location as LocationPoint)?.Point ?? XYZ.Zero;
        XYZ testRoomLocation = (testRoom.Location as LocationPoint)?.Point ?? GetRoomCentroid(testRoom);
        XYZ relativePosition = testRoomLocation - candidateGroupOrigin;
        double testRoomArea = testRoom.Area;
        string testRoomName = testRoom.Name;

        foreach (Group otherGroup in sameTypeGroups)
        {
            if (otherGroup.Id == candidateGroup.Id) continue;

            result.InstancesChecked++;
            XYZ otherGroupOrigin = (otherGroup.Location as LocationPoint)?.Point ?? XYZ.Zero;
            XYZ expectedRoomLocation = otherGroupOrigin + relativePosition;

            bool foundMatch = false;

            // Use spatial index to find rooms near expected location
            int gridX = (int)Math.Floor(expectedRoomLocation.X / ROOM_SPATIAL_GRID_SIZE);
            int gridY = (int)Math.Floor(expectedRoomLocation.Y / ROOM_SPATIAL_GRID_SIZE);

            // Check a wider area (3x3 grid)
            for (int dx = -1; dx <= 1; dx++)
            {
                for (int dy = -1; dy <= 1; dy++)
                {
                    int cellKey = GetRoomSpatialCellKey(gridX + dx, gridY + dy);
                    if (!_roomSpatialIndex.ContainsKey(cellKey)) continue;

                    foreach (var roomInfo in _roomSpatialIndex[cellKey])
                    {
                        double distance = roomInfo.RoomLocation.DistanceTo(expectedRoomLocation);

                        // Try strict tolerance first
                        if (distance <= ROOM_POSITION_TOLERANCE)
                        {
                            double areaDifference = Math.Abs(roomInfo.Area - testRoomArea) / testRoomArea;

                            if (areaDifference <= ROOM_AREA_TOLERANCE)
                            {
                                // Check name/number if available
                                bool nameMatches = string.IsNullOrEmpty(testRoomName) ||
                                                 string.IsNullOrEmpty(roomInfo.RoomName) ||
                                                 roomInfo.RoomName.Equals(testRoomName, StringComparison.OrdinalIgnoreCase);

                                if (nameMatches)
                                {
                                    foundMatch = true;
                                    result.MatchDetails.Add($"Group {otherGroup.Id}: Room at {distance:F2}ft");
                                    break;
                                }
                                else
                                {
                                    TrackStat("Name Mismatch");
                                }
                            }
                            else if (areaDifference <= ROOM_AREA_TOLERANCE_RELAXED)
                            {
                                // Try with relaxed area tolerance
                                foundMatch = true;
                                result.MatchDetails.Add($"Group {otherGroup.Id}: Room with relaxed area tolerance");
                                break;
                            }
                            else
                            {
                                result.MismatchDetails.Add(GetRoomMismatchReason(testRoom, expectedRoomLocation, 
                                    roomInfo, distance, areaDifference));
                                TrackStat("Area Mismatch");
                            }
                        }
                        else if (distance <= ROOM_POSITION_TOLERANCE_RELAXED)
                        {
                            // Try with relaxed position tolerance
                            double areaDifference = Math.Abs(roomInfo.Area - testRoomArea) / testRoomArea;
                            if (areaDifference <= ROOM_AREA_TOLERANCE_RELAXED)
                            {
                                foundMatch = true;
                                result.MatchDetails.Add($"Group {otherGroup.Id}: Room with relaxed tolerances");
                                break;
                            }
                        }
                    }
                    if (foundMatch) break;
                }
                if (foundMatch) break;
            }

            if (!foundMatch)
            {
                result.MismatchDetails.Add($"Group {otherGroup.Id}: No room found near expected location");
                TrackStat("Position Mismatch");
            }
            else
            {
                result.InstancesWithRoom++;
            }
        }

        // Determine if validation passes with relaxed thresholds
        double ratio = result.InstancesChecked > 0 ? (double)result.InstancesWithRoom / result.InstancesChecked : 0;

        if (ratio >= ROOM_VALIDATION_THRESHOLD && result.InstancesWithRoom >= ROOM_VALIDATION_MIN_MATCHES)
        {
            result.IsValid = true;
            result.PrimaryFailureReason = "N/A - Validated";
        }
        else
        {
            result.IsValid = false;
            if (result.InstancesWithRoom == 0)
            {
                result.PrimaryFailureReason = "No similar rooms found in any other instance";
                TrackStat("No Similar Rooms Found");
            }
            else
            {
                result.PrimaryFailureReason = $"Below threshold: only {ratio:P0} of instances have similar rooms (need {ROOM_VALIDATION_THRESHOLD:P0})";
                TrackStat("Below Threshold");
            }
        }

        return result;
    }

    // Check if room boundaries are mostly formed by group members
    private bool ShouldForceIncludeRoom(Room room, Group group, Document doc)
    {
        SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
        options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;

        IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);
        if (boundaries == null || boundaries.Count == 0) return false;

        ICollection<ElementId> groupMemberIds = group.GetMemberIds();
        int totalSegments = 0;
        int groupSegments = 0;

        foreach (IList<BoundarySegment> loop in boundaries)
        {
            foreach (BoundarySegment segment in loop)
            {
                totalSegments++;
                if (groupMemberIds.Contains(segment.ElementId))
                {
                    groupSegments++;
                }
            }
        }

        // If more than 50% of room boundaries are from group members, include it
        double ratio = totalSegments > 0 ? (double)groupSegments / totalSegments : 0;
        return ratio > 0.5;
    }

    // Get room centroid
    private XYZ GetRoomCentroid(Room room)
    {
        BoundingBoxXYZ bb = room.get_BoundingBox(null);
        if (bb != null)
        {
            return (bb.Min + bb.Max) * 0.5;
        }

        LocationPoint locPoint = room.Location as LocationPoint;
        if (locPoint != null)
        {
            return locPoint.Point;
        }

        return XYZ.Zero;
    }

    // Select the best group for a room based on boundary analysis
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
            return SelectGroupBySmallestBoundingBox(candidateGroups);
        }

        // Count boundary elements for each group
        Dictionary<Group, int> groupBoundaryCount = new Dictionary<Group, int>();
        Dictionary<Group, double> groupBoundaryLength = new Dictionary<Group, double>();

        foreach (Group group in candidateGroups)
        {
            int count = 0;
            double totalLength = 0.0;
            HashSet<ElementId> memberIds = new HashSet<ElementId>(group.GetMemberIds());

            foreach (IList<BoundarySegment> loop in boundaries)
            {
                foreach (BoundarySegment segment in loop)
                {
                    if (memberIds.Contains(segment.ElementId))
                    {
                        count++;
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
            return groupsWithMaxCount.First();
        }

        // If there's still a tie, use total boundary length
        if (maxCount > 0)
        {
            Group bestByLength = groupsWithMaxCount
                .OrderByDescending(g => groupBoundaryLength[g])
                .First();

            return bestByLength;
        }

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
                    LocationCurve wallLocation = wall.Location as LocationCurve;
                    if (wallLocation != null)
                    {
                        Curve wallCurve = wallLocation.Curve;
                        XYZ wallMidpoint = wallCurve.Evaluate(0.5, true);

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
            .ThenBy(kvp => GetGroupBoundingBoxVolume(kvp.Key))
            .First().Key;

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

    // Select group by smallest bounding box
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

        // Add small tolerance
        double tolerance = 0.1;

        return roomBB.Min.X >= bbMin.X - tolerance &&
               roomBB.Min.Y >= bbMin.Y - tolerance &&
               roomBB.Min.Z >= bbMin.Z - tolerance &&
               roomBB.Max.X <= bbMax.X + tolerance &&
               roomBB.Max.Y <= bbMax.Y + tolerance &&
               roomBB.Max.Z <= bbMax.Z + tolerance;
    }

    // Enhanced: Get elements contained in group's rooms with simplified diagnostics
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
            BoundingBoxXYZ bb = member?.get_BoundingBox(null);
            if (bb != null)
            {
                groupMinZ = Math.Min(groupMinZ, bb.Min.Z);
                groupMaxZ = Math.Max(groupMaxZ, bb.Max.Z);
            }
        }

        // Add tolerance to height bounds
        groupMinZ -= 1.0; // 1 foot below
        groupMaxZ += 1.0; // 1 foot above

        // Check each candidate element against all rooms
        foreach (Element elem in candidateElements)
        {
            List<XYZ> testPoints = _elementTestPointsCache[elem.Id];
            bool foundInRoom = false;
            string containingRoomName = "";

            foreach (Room room in groupRooms)
            {
                if (IsElementInRoomOptimized(elem, room, testPoints, groupMinZ, groupMaxZ))
                {
                    containedElements.Add(elem);
                    foundInRoom = true;
                    containingRoomName = room.Name ?? room.Number ?? "Unnamed";
                    break;
                }
            }

            // Log only if element not found (using simplified diagnostic)
            if (!foundInRoom)
            {
                LogElementContainmentCheck(elem, group, false, "", doc);
            }
        }

        return containedElements;
    }

    // Optimized element-in-room check
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

            // For ALL rooms, check XY containment at room level
            double testZ = roomData.LevelZ + 1.0; // 1 foot above room level
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
