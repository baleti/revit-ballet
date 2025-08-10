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
