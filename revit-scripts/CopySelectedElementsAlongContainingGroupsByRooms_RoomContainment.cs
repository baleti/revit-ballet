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
    private Dictionary<ElementId, Group> BuildRoomToGroupMapping(IList<Group> relevantGroups, Document doc, BoundingBoxXYZ selectedElementsBounds = null)
    {
        StartTiming("BuildRoomToGroupMapping");
        
        Dictionary<ElementId, RoomMappingInfo> roomMappingInfo = new Dictionary<ElementId, RoomMappingInfo>();
        Dictionary<ElementId, List<Group>> roomDirectMembership = new Dictionary<ElementId, List<Group>>();
        Dictionary<ElementId, List<Group>> roomSpatialContainment = new Dictionary<ElementId, List<Group>>();

        // Build spatial index for all rooms first
        BuildRoomSpatialIndex(doc);

        // First pass: identify direct memberships
        StartTiming("DirectMembershipPass");
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
        EndTiming("DirectMembershipPass");

        // Second pass: identify spatial containments
        StartTiming("SpatialContainmentPass");
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
       EndTiming("SpatialContainmentPass");

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
       StartTiming("ValidateSpatialContainments");
       FilteredElementCollector allGroupsCollector = new FilteredElementCollector(doc);
       IList<Group> allGroups = allGroupsCollector
           .OfClass(typeof(Group))
           .WhereElementIsNotElementType()
           .Cast<Group>()
           .ToList();

       Dictionary<ElementId, List<Group>> validatedSpatialContainments = ValidateSpatialContainmentsEnhanced(
           roomSpatialContainment, allGroups, doc);
       EndTiming("ValidateSpatialContainments");

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
       
       // THIRD PASS: Map rooms based on boundary walls
       Dictionary<ElementId, Group> finalMapping = MapRoomsByBoundaryWalls(
           roomToGroupMap, 
           relevantGroups, 
           doc,
           selectedElementsBounds);  // Pass bounds for filtering

       EndTiming("BuildRoomToGroupMapping");
       return finalMapping;
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

   // Map rooms to groups based on boundary walls
   private Dictionary<ElementId, Group> MapRoomsByBoundaryWalls(
       Dictionary<ElementId, Group> existingMapping,
       IList<Group> allGroups,
       Document doc,
       BoundingBoxXYZ selectedElementsBounds = null)  // ADD PARAMETER
   {
       StartTiming("MapRoomsByBoundaryWalls");
       
       // Start with existing mappings
       Dictionary<ElementId, Group> enhancedMapping = new Dictionary<ElementId, Group>(existingMapping);
       
       // OPTIMIZATION: Pre-filter rooms to only those near selected elements
       List<ElementId> roomsToCheck = new List<ElementId>();
       // Increased buffer from 20 to 50 feet to ensure we don't miss rooms
       double buffer = 50.0; // Increased from 20 feet
       if (selectedElementsBounds != null)
       {
           foreach (var kvp in _roomDataCache)
           {
               RoomData roomData = kvp.Value;
               if (roomData.BoundingBox != null)
               {
                   // Quick bounds check
                   if (!(roomData.BoundingBox.Max.X < selectedElementsBounds.Min.X - buffer ||
                         roomData.BoundingBox.Min.X > selectedElementsBounds.Max.X + buffer ||
                         roomData.BoundingBox.Max.Y < selectedElementsBounds.Min.Y - buffer ||
                         roomData.BoundingBox.Min.Y > selectedElementsBounds.Max.Y + buffer))
                   {
                       roomsToCheck.Add(kvp.Key);
                   }
               }
           }
       }
       else
       {
           roomsToCheck = _roomDataCache.Keys.ToList();
       }
       
       // Build a map of wall IDs to their containing groups
       Dictionary<ElementId, List<Group>> wallToGroups = new Dictionary<ElementId, List<Group>>();
       foreach (Group group in allGroups)
       {
           foreach (ElementId memberId in group.GetMemberIds())
           {
               Element member = doc.GetElement(memberId);
               if (member is Wall)
               {
                   if (!wallToGroups.ContainsKey(memberId))
                   {
                       wallToGroups[memberId] = new List<Group>();
                   }
                   wallToGroups[memberId].Add(group);
               }
           }
       }
       
       // Check unmapped rooms
       // OPTIMIZATION: Only check filtered rooms
       foreach (ElementId roomId in roomsToCheck)
       {
           
           // Skip if already mapped
           if (enhancedMapping.ContainsKey(roomId))
               continue;
               
           Room room = doc.GetElement(roomId) as Room;
           if (room == null || room.Area <= 0)
               continue;
           
           // Get room boundaries
           // OPTIMIZATION: Cache boundary segments
           IList<IList<BoundarySegment>> boundaries;
           if (!_roomBoundaryCache.ContainsKey(roomId))
           {
               StartTiming("GetBoundarySegments");
               SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
               options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
               
               boundaries = room.GetBoundarySegments(options);
               _roomBoundaryCache[roomId] = boundaries;
               EndTiming("GetBoundarySegments");
           }
           else
           {
               boundaries = _roomBoundaryCache[roomId];
           }
           
           if (boundaries == null || boundaries.Count == 0)
               continue;
           
           // Track which groups contain boundary walls
           Dictionary<Group, int> groupBoundaryCount = new Dictionary<Group, int>();
           int totalWallSegments = 0;
           
           foreach (IList<BoundarySegment> loop in boundaries)
           {
               foreach (BoundarySegment segment in loop)
               {
                   ElementId boundaryId = segment.ElementId;
                   Element boundaryElem = doc.GetElement(boundaryId);
                   
                   // Only count walls (not room separation lines)
                   if (boundaryElem is Wall)
                   {
                       totalWallSegments++;
                       
                       // Check which groups contain this wall
                       if (wallToGroups.ContainsKey(boundaryId))
                       {
                           foreach (Group group in wallToGroups[boundaryId])
                           {
                               if (!groupBoundaryCount.ContainsKey(group))
                               {
                                   groupBoundaryCount[group] = 0;
                               }
                               groupBoundaryCount[group]++;
                           }
                       }
                   }
               }
           }
           
           // If a significant portion of walls belong to one group, map the room to it
           if (totalWallSegments > 0)
           {
               foreach (var groupKvp in groupBoundaryCount)
               {
                   Group group = groupKvp.Key;
                   int wallCount = groupKvp.Value;
                   double ratio = (double)wallCount / totalWallSegments;
                   
                   // Reduced threshold from 0.75 to 0.50 to capture more room-group relationships
                   // This was causing rooms to not be mapped to their groups
                   if (ratio >= 0.50)
                   {
                       enhancedMapping[roomId] = group;
                       
                       if (enableDiagnostics)
                       {
                           GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
                           diagnosticLog.AppendLine($"  Room {room.Number} mapped to group {gt?.Name} via boundary walls ({wallCount}/{totalWallSegments} = {ratio:P0})");
                       }
                       break; // Use first group that meets threshold
                   }
               }
           }
       }
       
       EndTiming("MapRoomsByBoundaryWalls");
       return enhancedMapping;
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
       StartTiming("GetElementsContainedInGroupRoomsFiltered");
       List<Element> containedElements = new List<Element>();

       // OPTIMIZATION: Early exit if group has no rooms
       if (!roomToGroupMap.Any(kvp => kvp.Value.Id == group.Id))
       {
           EndTiming("GetElementsContainedInGroupRoomsFiltered");
           return containedElements;
       }
       
       // OPTIMIZATION: Pre-filter candidates by elevation
       BoundingBoxXYZ groupBB = _groupBoundingBoxCache.ContainsKey(group.Id) 
           ? _groupBoundingBoxCache[group.Id] 
           : null;
           
       if (groupBB == null)
       {
           EndTiming("GetElementsContainedInGroupRoomsFiltered");
           return containedElements;
       }
       
       List<Element> elevationFilteredElements = new List<Element>();
       
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
       
       // OPTIMIZATION: Filter elements by Z elevation FIRST
       foreach (Element elem in candidateElements)
       {
           List<XYZ> testPoints = _elementTestPointsCache[elem.Id];
           if (testPoints.Count > 0)
           {
               double elemZ = testPoints[0].Z;
               // Quick elevation check
               if (elemZ >= groupMinZ && elemZ <= groupMaxZ)
               {
                   elevationFilteredElements.Add(elem);
               }
           }
       }

       // Check only elevation-filtered elements
       foreach (Element elem in elevationFilteredElements)
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

       EndTiming("GetElementsContainedInGroupRoomsFiltered");
       return containedElements;
   }

   // Get actual floor-to-floor height based on floor slabs above and below the element
   private double GetActualFloorToFloorHeight(Level currentLevel, Document doc, XYZ testPoint)
   {
       StartTiming("GetActualFloorToFloorHeight");
       
       // OPTIMIZATION: Cache by grid location
       // Use coarser grid (50ft instead of 10ft) for better cache hits
       int gridX = (int)Math.Floor(testPoint.X / 50.0);
       int gridY = (int)Math.Floor(testPoint.Y / 50.0);
       string cacheKey = $"L{currentLevel.Id.IntegerValue}_X{gridX}_Y{gridY}";
       
       if (_floorToFloorHeightCache.ContainsKey(cacheKey))
       {
           EndTiming("GetActualFloorToFloorHeight");
           return _floorToFloorHeightCache[cacheKey];
       }
       
       double floorToFloorHeight = 10.0; // Default value
       
       if (currentLevel == null || testPoint == null) return 10.0; // Default

       // Find all floor elements in the document
       FilteredElementCollector floorCollector = new FilteredElementCollector(doc);
       List<Floor> allFloors = floorCollector
           .OfClass(typeof(Floor))
           .WhereElementIsNotElementType()
           .Cast<Floor>()
           .Where(f => f != null)
           .ToList();


       // Find floors at this XY location with their elevations
       List<(Floor floor, double elevation, ElementId id)> floorsAtLocation = new List<(Floor, double, ElementId)>();
       
       foreach (Floor floor in allFloors)
       {
           try
           {
               // Get floor bounding box for quick check
               BoundingBoxXYZ bb = floor.get_BoundingBox(null);
               if (bb == null) continue;
               
               // Check if test point is within floor's XY bounds (with some tolerance)
               double tolerance = 1.0; // 1 foot tolerance
               if (testPoint.X >= bb.Min.X - tolerance && testPoint.X <= bb.Max.X + tolerance &&
                   testPoint.Y >= bb.Min.Y - tolerance && testPoint.Y <= bb.Max.Y + tolerance)
               {
                   // This floor is at our XY location
                   // Use the top of the bounding box as the floor elevation
                   double floorTopZ = bb.Max.Z;
                   floorsAtLocation.Add((floor, floorTopZ, floor.Id));
               }
           }
           catch
           {
               // Skip floors that cause issues
               continue;
           }
       }

       // Sort floors by elevation
       floorsAtLocation = floorsAtLocation.OrderBy(f => f.elevation).ToList();
       
       // Find the floor at the room's level and the one above it
       // We want the floor-to-floor height for the ROOM, not for where the element is
       double roomLevelElevation = currentLevel.Elevation;
       
       var floorBelow = floorsAtLocation
           .Where(f => Math.Abs(f.elevation - roomLevelElevation) < 2.0) // Floor at room's level
           .OrderBy(f => Math.Abs(f.elevation - roomLevelElevation))
           .FirstOrDefault();
           
       if (floorBelow.floor != null)
       {
           // Find the next floor above the room's floor
           var floorAbove = floorsAtLocation
               .Where(f => f.elevation > floorBelow.elevation + 1.0) // At least 1 foot above
               .OrderBy(f => f.elevation)
               .FirstOrDefault();
       
           if (floorAbove.floor != null)
           {
               double floorToFloor = floorAbove.elevation - floorBelow.elevation;
               
               // Sanity check - floor-to-floor should be reasonable
               if (floorToFloor > 5.0 && floorToFloor < 20.0)
               {
                   if (enableDiagnostics)
                   {
                       diagnosticLog.AppendLine($"        Floor-to-floor from actual slabs at ROOM level:");
                       diagnosticLog.AppendLine($"        Room level: {currentLevel.Name} at {roomLevelElevation:F2}ft");
                       diagnosticLog.AppendLine($"        Floor at room level (ID: {floorBelow.id}): {floorBelow.elevation:F2}ft");
                       diagnosticLog.AppendLine($"        Floor above room (ID: {floorAbove.id}): {floorAbove.elevation:F2}ft");
                       diagnosticLog.AppendLine($"        Actual F2F height: {floorToFloor:F2}ft");
                   }
                   floorToFloorHeight = floorToFloor;
                   _floorToFloorHeightCache[cacheKey] = floorToFloorHeight;
                   EndTiming("GetActualFloorToFloorHeight");
                   return floorToFloorHeight;
               }
               floorToFloorHeight = floorToFloor;
               _floorToFloorHeightCache[cacheKey] = floorToFloorHeight;
               EndTiming("GetActualFloorToFloorHeight");
               return floorToFloorHeight;
           }
       }
       
       // If we only found floors but not bracketing the element, still try to find a reasonable F2F
       if (floorsAtLocation.Count >= 2)
       {
           // Find the floor closest to the current level
           double levelElevation = currentLevel.Elevation;
           var currentFloor = floorsAtLocation
               .Where(f => Math.Abs(f.elevation - levelElevation) < 2.0) // Within 2 feet
               .OrderBy(f => Math.Abs(f.elevation - levelElevation))
               .FirstOrDefault();
           
           if (currentFloor.floor != null)
           {
               // Find the next floor above
               var nextFloor = floorsAtLocation
                   .Where(f => f.elevation > currentFloor.elevation + 1.0) // At least 1 foot above
                   .OrderBy(f => f.elevation)
                   .FirstOrDefault();
               
               if (nextFloor.floor != null)
               {
                   double floorToFloor = nextFloor.elevation - currentFloor.elevation;
                   
                   // Sanity check - floor-to-floor should be reasonable
                   if (floorToFloor > 5.0 && floorToFloor < 20.0)
                   {
                       if (enableDiagnostics)
                       {
                           diagnosticLog.AppendLine($"        Found floor-to-floor from actual floors: {floorToFloor:F2}ft");
                           diagnosticLog.AppendLine($"        Current floor at: {currentFloor.elevation:F2}ft");
                           diagnosticLog.AppendLine($"        Next floor at: {nextFloor.elevation:F2}ft");
                       }
                       floorToFloorHeight = floorToFloor;
                       _floorToFloorHeightCache[cacheKey] = floorToFloorHeight;
                       EndTiming("GetActualFloorToFloorHeight");
                       return floorToFloorHeight;
                   }
               }
           }
       }

       // Fallback to level-based calculation if no floors found
       if (enableDiagnostics)
       {
           diagnosticLog.AppendLine($"        No floors found at location, using level-based calculation");
       }

       // Get all levels in the document
       FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
       List<Level> allLevels = levelCollector
           .OfClass(typeof(Level))
           .Cast<Level>()
           .OrderBy(l => l.Elevation)
           .ToList();

       // Find the next level above current level
       Level nextLevel = null;
       double currentElevation = currentLevel.Elevation;
       
       foreach (Level level in allLevels)
       {
           if (level.Elevation > currentElevation)
           {
               if (nextLevel == null || level.Elevation < nextLevel.Elevation)
               {
                   nextLevel = level;
               }
           }
       }

       // If we found a next level, return the difference
       if (nextLevel != null)
       {
           floorToFloorHeight = nextLevel.Elevation - currentElevation;
           _floorToFloorHeightCache[cacheKey] = floorToFloorHeight;
           EndTiming("GetActualFloorToFloorHeight");
           return floorToFloorHeight;
       }

       EndTiming("GetActualFloorToFloorHeight");
       _floorToFloorHeightCache[cacheKey] = floorToFloorHeight;  // Cache the default
       return floorToFloorHeight;
   }


   // Optimized element-in-room check
   private bool IsElementInRoomOptimized(Element element, Room room, List<XYZ> testPoints, double groupMinZ, double groupMaxZ)
   {
       StartTiming("IsElementInRoomOptimized");
       
       // Skip unplaced or unenclosed rooms
       if (room.Area <= 0)
       {
           EndTiming("IsElementInRoomOptimized");
           return false;
       }

       // Get cached room data
       RoomData roomData = _roomDataCache[room.Id];

       // OPTIMIZATION: Quick rejection based on Z bounds
       double elementZ = testPoints[0].Z;
       double maxPossibleHeight = roomData.LevelZ + 20.0; // Max reasonable floor height
       
       if (elementZ < roomData.LevelZ - 2.0 || elementZ > maxPossibleHeight)
       {
           // Element is clearly outside room's possible Z range
           EndTiming("IsElementInRoomOptimized");
           return false;
       }
       
       // OPTIMIZATION: Check if we've already calculated F2F for this room's level
       bool needsF2FCheck = Math.Abs(roomData.Height - 10.0) < 2.0; // Room height seems wrong
       
       foreach (XYZ point in testPoints)
       {
           // Check XY containment - try multiple Z heights to ensure we catch the room
           List<double> testZHeights = new List<double>();
           
           // Add various test heights
           testZHeights.Add(roomData.LevelZ + 1.0); // 1 foot above room level
           testZHeights.Add(roomData.LevelZ + roomData.Height * 0.5); // Middle of room height
           if (roomData.BoundingBox != null)
           {
               testZHeights.Add((roomData.BoundingBox.Min.Z + roomData.BoundingBox.Max.Z) * 0.5);
               testZHeights.Add(roomData.BoundingBox.Min.Z + 1.0);
           }

           bool inRoomXY = false;
           foreach (double testZ in testZHeights)
           {
               XYZ testPointForXY = new XYZ(point.X, point.Y, testZ);
               
               // OPTIMIZATION: Cache room containment checks
               string containmentKey = $"R{room.Id.IntegerValue}_X{(int)(point.X*10)}_Y{(int)(point.Y*10)}_Z{(int)(testZ*10)}";
               bool isInRoom;
               if (!_roomPointContainmentCache.ContainsKey(containmentKey))
               {
                   StartTiming("Room.IsPointInRoom");
                   isInRoom = room.IsPointInRoom(testPointForXY);
                   EndTiming("Room.IsPointInRoom");
                   _roomPointContainmentCache[containmentKey] = isInRoom;
               }
               else
               {
                   isInRoom = _roomPointContainmentCache[containmentKey];
               }
               
               if (isInRoom)
               {
                   inRoomXY = true;
                   break;
               }
           }

           if (inRoomXY)
           {
               // Element is within room's XY boundaries
               
               // First check if it's within the room's defined height
               if (point.Z >= roomData.MinZ && point.Z <= roomData.MaxZ)
               {
                   if (enableDiagnostics)
                   {
                       diagnosticLog.AppendLine($"  Element in room {room.Number} using room height bounds");
                   }
                   EndTiming("IsElementInRoomOptimized");
                   return true;
               }
               
               // If not within room's stated height, use actual floor-to-floor height
               // Get actual floor-to-floor height at this specific location
               double actualFloorHeight = GetActualFloorToFloorHeight(roomData.Level, element.Document, point);
               
               // Always prefer actual floor-to-floor height if it's different from room height
               if (Math.Abs(actualFloorHeight - roomData.Height) > 0.5) // More than 6 inches difference
               {
                   double actualMaxZ = roomData.LevelZ + actualFloorHeight;
                   
                   // Add tolerance for elements near boundaries
                   double tolerance = 1.0; // 1 foot tolerance
                   
                   if (point.Z >= roomData.LevelZ - tolerance && point.Z <= actualMaxZ + tolerance)
                   {
                       // Element is within actual floor-to-floor height
                       if (enableDiagnostics)
                       {
                           diagnosticLog.AppendLine($"  >>> Element INCLUDED in room {room.Number} using F2F height <<<");
                           diagnosticLog.AppendLine($"      Room height: {roomData.Height:F2}ft, Actual F2F: {actualFloorHeight:F2}ft");
                           diagnosticLog.AppendLine($"      Room Z range with F2F: {roomData.LevelZ:F2}ft to {actualMaxZ:F2}ft");
                           diagnosticLog.AppendLine($"      Element Z: {point.Z:F2}ft");
                       }
                       EndTiming("IsElementInRoomOptimized");
                       return true;
                   }
               }
           }
       }

       EndTiming("IsElementInRoomOptimized");
       return false;
   }
}
