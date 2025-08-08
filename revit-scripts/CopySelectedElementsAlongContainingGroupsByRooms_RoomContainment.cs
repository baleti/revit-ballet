using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
    // Build comprehensive room-to-group mapping with priority handling
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
        
        // Build final mapping with priority
        foreach (var kvp in roomDirectMembership)
        {
            ElementId roomId = kvp.Key;
            RoomData roomData = _roomDataCache[roomId];
            
            // Sort groups by specificity (smaller bounding box = more specific)
            List<Group> sortedGroups = SortGroupsBySpecificity(kvp.Value, roomData.BoundingBox);
            roomToGroupMap[roomId] = sortedGroups.First();
        }
        
        // Add spatial containments (sorted by specificity, only for rooms without direct membership)
        foreach (var kvp in roomSpatialContainment)
        {
            if (!roomToGroupMap.ContainsKey(kvp.Key))
            {
                ElementId roomId = kvp.Key;
                RoomData roomData = _roomDataCache[roomId];
                
                // Sort groups by specificity (smaller bounding box = more specific)
                List<Group> sortedGroups = SortGroupsBySpecificity(kvp.Value, roomData.BoundingBox);
                roomToGroupMap[roomId] = sortedGroups.First();
            }
        }
        
        return roomToGroupMap;
    }
    
    // Sort groups by specificity relative to a reference bounding box
    private List<Group> SortGroupsBySpecificity(List<Group> groups, BoundingBoxXYZ referenceBB)
    {
        if (groups.Count <= 1) return groups;
        
        // Calculate reference volume (if available)
        double refVolume = 1.0; // Default to 1 to avoid division by zero
        if (referenceBB != null)
        {
            XYZ refSize = referenceBB.Max - referenceBB.Min;
            refVolume = refSize.X * refSize.Y * refSize.Z;
            if (refVolume <= 0) refVolume = 1.0;
        }
        
        // Create list of groups with their specificity scores
        List<(Group group, double score)> groupsWithScores = new List<(Group, double)>();
        
        foreach (Group group in groups)
        {
            BoundingBoxXYZ groupBB = _groupBoundingBoxCache.ContainsKey(group.Id) 
                ? _groupBoundingBoxCache[group.Id] 
                : null;
            
            if (groupBB != null)
            {
                XYZ groupSize = groupBB.Max - groupBB.Min;
                double groupVolume = groupSize.X * groupSize.Y * groupSize.Z;
                
                // Specificity score: ratio of group volume to reference volume
                // Smaller ratio = more specific (group is closer in size to reference)
                double score = groupVolume / refVolume;
                groupsWithScores.Add((group, score));
            }
            else
            {
                // If no bounding box, give it a high score (low priority)
                groupsWithScores.Add((group, double.MaxValue));
            }
        }
        
        // Sort by score (ascending - smaller scores are more specific)
        groupsWithScores.Sort((a, b) => a.score.CompareTo(b.score));
        
        return groupsWithScores.Select(gs => gs.group).ToList();
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
