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
    
    // Diagnostic data collection
    private StringBuilder _diagnostics = new StringBuilder();
    private Dictionary<string, List<string>> _groupDiagnostics = new Dictionary<string, List<string>>();
    
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
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        try
        {
            // Get selected elements
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            
            if (selectedIds.Count == 0)
            {
                message = "Please select elements first";
                return Result.Failed;
            }
            
            // Start diagnostics
            _diagnostics.AppendLine("=== DIAGNOSTIC REPORT ===");
            _diagnostics.AppendLine($"Total selected IDs: {selectedIds.Count}");
            _diagnostics.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _diagnostics.AppendLine();
            
            // Get selected elements - include everything except Groups themselves and ElementTypes
            List<Element> selectedElements = new List<Element>();
            _diagnostics.AppendLine("=== ELEMENT FILTERING ===");
            
            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem != null)
                {
                    string elemInfo = $"ID {id.IntegerValue}: {elem.GetType().Name}";
                    
                    if (elem is Group)
                    {
                        _diagnostics.AppendLine($"  {elemInfo} - EXCLUDED (is a Group instance)");
                    }
                    else if (elem is ElementType)
                    {
                        _diagnostics.AppendLine($"  {elemInfo} - EXCLUDED (is an ElementType)");
                    }
                    else if (elem is GroupType)
                    {
                        _diagnostics.AppendLine($"  {elemInfo} - EXCLUDED (is a GroupType)");
                    }
                    else
                    {
                        selectedElements.Add(elem);
                        _diagnostics.AppendLine($"  {elemInfo} - INCLUDED (Category: {elem.Category?.Name ?? "None"})");
                    }
                }
                else
                {
                    _diagnostics.AppendLine($"  ID {id.IntegerValue}: NULL element - EXCLUDED");
                }
            }
            
            _diagnostics.AppendLine($"\nFiltered elements count: {selectedElements.Count}");
            
            if (selectedElements.Count == 0)
            {
                _diagnostics.AppendLine("\nNO VALID ELEMENTS AFTER FILTERING!");
                _diagnostics.AppendLine("Make sure you're selecting actual elements, not group instances.");
                
                // Show diagnostics before failing
                ShowDiagnostics();
                
                message = "No valid elements found in selection. Please select elements (not group instances).\n\nCheck the diagnostic report for details.";
                return Result.Failed;
            }
            
            // Pre-build room cache and spatial index
            _diagnostics.AppendLine("\n=== BUILDING SPATIAL INDEX ===");
            SpatialIndex roomIndex = new SpatialIndex();
            List<Room> allRooms = BuildRoomCacheAndIndex(doc, roomIndex);
            _diagnostics.AppendLine($"Indexed {allRooms.Count} rooms");
            
            // Pre-calculate element test points
            _diagnostics.AppendLine("\n=== PRE-CALCULATING ELEMENT DATA ===");
            Dictionary<ElementId, ElementData> elementDataCache = new Dictionary<ElementId, ElementData>();
            foreach (Element elem in selectedElements)
            {
                elementDataCache[elem.Id] = new ElementData
                {
                    TestPoints = GetElementTestPoints(elem),
                    BoundingBox = elem.get_BoundingBox(null),
                    Level = doc.GetElement(elem.LevelId) as Level
                };
            }
            
            // Find which rooms contain the selected elements using spatial index
            _diagnostics.AppendLine("\n=== ROOM CONTAINMENT CHECK (WITH SPATIAL INDEX) ===");
            Dictionary<ElementId, List<Room>> elementToRoomsMap = FindRoomsContainingElementsOptimized(
                selectedElements, elementDataCache, roomIndex);
            
            // Report findings
            int elementsInRooms = elementToRoomsMap.Count(kvp => kvp.Value.Count > 0);
            _diagnostics.AppendLine($"Elements in at least one room: {elementsInRooms}/{selectedElements.Count}");
            
            if (elementsInRooms == 0)
            {
                _diagnostics.AppendLine("\nNO ELEMENTS ARE IN ANY ROOMS!");
                ShowDiagnostics();
                
                TaskDialog.Show("No Rooms Found", 
                    "None of the selected elements are contained in any rooms.\n\n" +
                    "Check the diagnostic report for details.");
                return Result.Succeeded;
            }
            
            // Now find which groups contain these rooms (either as members OR spatially)
            _diagnostics.AppendLine("\n=== GROUP ANALYSIS ===");
            Dictionary<ElementId, List<Group>> elementToGroupsMap = FindGroupsContainingRoomsOptimized(
                elementToRoomsMap, doc);
            
            // Find maximum number of groups any single element is contained in
            int maxGroupsPerElement = 0;
            foreach (var groups in elementToGroupsMap.Values)
            {
                maxGroupsPerElement = Math.Max(maxGroupsPerElement, groups.Count);
            }
            
            _diagnostics.AppendLine($"\n=== SUMMARY ===");
            _diagnostics.AppendLine($"Elements in rooms: {elementsInRooms}");
            _diagnostics.AppendLine($"Elements in groups (via rooms): {elementToGroupsMap.Count(kvp => kvp.Value.Count > 0)}");
            _diagnostics.AppendLine($"Maximum groups per element: {maxGroupsPerElement}");
            
            // Show diagnostics
            ShowDiagnostics();
            
            if (maxGroupsPerElement == 0)
            {
                TaskDialog.Show("No Groups Found", 
                    "The selected elements are in rooms, but those rooms are not part of any groups.\n\n" +
                    "Check the diagnostic report for details.");
                return Result.Succeeded;
            }
            
            // Prepare data for DataGrid
            List<Dictionary<string, object>> gridEntries = PrepareGridEntries(
                selectedElements, elementToRoomsMap, elementToGroupsMap, 
                maxGroupsPerElement, doc);
            
            // Prepare property names
            List<string> propertyNames = new List<string> { "Type", "Family", "Id", "Comments", "Room(s)" };
            for (int i = 1; i <= maxGroupsPerElement; i++)
            {
                propertyNames.Add($"Group {i}");
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
                // Get the element from the hidden key
                if (entry.ContainsKey("_Element") && entry["_Element"] is Element)
                {
                    Element elem = entry["_Element"] as Element;
                    elementsToSelect.Add(elem);
                }
            }
            
            // Set selection to chosen elements
            if (elementsToSelect.Count > 0)
            {
                List<ElementId> elementIds = elementsToSelect.Select(e => e.Id).ToList();
                uidoc.SetSelectionIds(elementIds);
                
                TaskDialog.Show("Success", 
                    $"Selected {elementIds.Count} element(s) based on your selection.");
            }
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            _diagnostics.AppendLine($"\n=== EXCEPTION ===");
            _diagnostics.AppendLine($"Message: {ex.Message}");
            _diagnostics.AppendLine($"Stack Trace: {ex.StackTrace}");
            ShowDiagnostics();
            return Result.Failed;
        }
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
    
    // Find which groups contain the rooms - Same as before but with minor optimizations
    private Dictionary<ElementId, List<Group>> FindGroupsContainingRoomsOptimized(
        Dictionary<ElementId, List<Room>> elementToRoomsMap, Document doc)
    {
        Dictionary<ElementId, List<Group>> result = new Dictionary<ElementId, List<Group>>();
        
        // Initialize map
        foreach (var kvp in elementToRoomsMap)
        {
            result[kvp.Key] = new List<Group>();
        }
        
        // Get all groups
        FilteredElementCollector groupCollector = new FilteredElementCollector(doc);
        List<Group> allGroups = groupCollector
            .OfClass(typeof(Group))
            .WhereElementIsNotElementType()
            .Cast<Group>()
            .ToList();
        
        _diagnostics.AppendLine($"Total groups in document: {allGroups.Count}");
        
        // Build a map of room ID to groups containing that room
        Dictionary<ElementId, List<Group>> roomToGroupsMap = new Dictionary<ElementId, List<Group>>();
        
        // Get all unique rooms from the element-to-rooms map
        HashSet<Room> allRelevantRooms = new HashSet<Room>();
        foreach (var rooms in elementToRoomsMap.Values)
        {
            foreach (Room room in rooms)
            {
                allRelevantRooms.Add(room);
            }
        }
        
        _diagnostics.AppendLine($"Checking {allRelevantRooms.Count} unique rooms for group containment");
        
        // FIRST PASS: Find all rooms that are direct members of groups
        HashSet<ElementId> roomsWithDirectMembership = new HashSet<ElementId>();
        
        // Pre-cache group member IDs for faster lookup
        Dictionary<Group, HashSet<ElementId>> groupMemberCache = new Dictionary<Group, HashSet<ElementId>>();
        foreach (Group group in allGroups)
        {
            groupMemberCache[group] = new HashSet<ElementId>(group.GetMemberIds());
        }
        
        foreach (Group group in allGroups)
        {
            HashSet<ElementId> memberIds = groupMemberCache[group];
            
            foreach (Room room in allRelevantRooms)
            {
                if (memberIds.Contains(room.Id))
                {
                    roomsWithDirectMembership.Add(room.Id);
                    
                    if (!roomToGroupsMap.ContainsKey(room.Id))
                    {
                        roomToGroupsMap[room.Id] = new List<Group>();
                    }
                    roomToGroupsMap[room.Id].Add(group);
                }
            }
        }
        
        _diagnostics.AppendLine($"Rooms with direct group membership: {roomsWithDirectMembership.Count}");
        
        // Track room-to-group relationships for debugging
        Dictionary<string, HashSet<string>> roomGroupDebug = new Dictionary<string, HashSet<string>>();
        
        // SECOND PASS: Check spatial containment ONLY for rooms WITHOUT direct membership
        foreach (Group group in allGroups)
        {
            GroupType groupType = doc.GetElement(group.GetTypeId()) as GroupType;
            string groupName = groupType?.Name ?? "Unknown";
            
            // Get group bounding box
            BoundingBoxXYZ groupBB = group.get_BoundingBox(null);
            if (groupBB == null) continue;
            
            // Expand bounding box slightly to account for tolerance
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
            
            int roomsInThisGroup = 0;
            
            // Check each relevant room
            foreach (Room room in allRelevantRooms)
            {
                string roomKey = $"{room.Name ?? "Unnamed"} ({room.Id.IntegerValue})";
                bool isInGroup = false;
                string containmentType = "";
                
                // Skip spatial check if room already has direct membership in ANY group
                if (roomsWithDirectMembership.Contains(room.Id))
                {
                    if (roomToGroupsMap.ContainsKey(room.Id) && roomToGroupsMap[room.Id].Contains(group))
                    {
                        isInGroup = true;
                        containmentType = "MEMBER";
                    }
                }
                else
                {
                    // Only check spatial containment for rooms WITHOUT any direct membership
                    RoomData roomData = _roomDataCache[room.Id];
                    if (roomData.BoundingBox != null && IsRoomWithinBoundingBoxOptimized(roomData, minExpanded, maxExpanded))
                    {
                        isInGroup = true;
                        containmentType = "SPATIAL";
                        
                        if (!roomToGroupsMap.ContainsKey(room.Id))
                        {
                            roomToGroupsMap[room.Id] = new List<Group>();
                        }
                        roomToGroupsMap[room.Id].Add(group);
                    }
                }
                
                if (isInGroup)
                {
                    roomsInThisGroup++;
                    
                    if (!roomGroupDebug.ContainsKey(roomKey))
                    {
                        roomGroupDebug[roomKey] = new HashSet<string>();
                    }
                    roomGroupDebug[roomKey].Add($"{groupName} ({containmentType})");
                }
            }
            
            if (roomsInThisGroup > 0)
            {
                _diagnostics.AppendLine($"  Group '{groupName}': contains {roomsInThisGroup} rooms");
            }
        }
        
        // Report room-group relationships
        _diagnostics.AppendLine($"\n=== ROOM-GROUP ANALYSIS ===");
        foreach (var kvp in roomGroupDebug)
        {
            if (kvp.Value.Count > 1)
            {
                _diagnostics.AppendLine($"  Room {kvp.Key} is in {kvp.Value.Count} groups:");
                foreach (string groupInfo in kvp.Value)
                {
                    _diagnostics.AppendLine($"    - {groupInfo}");
                }
            }
            else if (kvp.Value.Count == 1)
            {
                _diagnostics.AppendLine($"  Room {kvp.Key}: {kvp.Value.First()}");
            }
        }
        
        // Now map elements to groups via their containing rooms
        _diagnostics.AppendLine("\n=== ELEMENT-GROUP MAPPING ===");
        foreach (var kvp in elementToRoomsMap)
        {
            ElementId elemId = kvp.Key;
            List<Room> rooms = kvp.Value;
            
            HashSet<Group> uniqueGroups = new HashSet<Group>();
            Dictionary<string, List<string>> groupToRooms = new Dictionary<string, List<string>>();
            
            foreach (Room room in rooms)
            {
                if (roomToGroupsMap.ContainsKey(room.Id))
                {
                    foreach (Group group in roomToGroupsMap[room.Id])
                    {
                        uniqueGroups.Add(group);
                        
                        GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
                        string gName = gt?.Name ?? "Unknown";
                        if (!groupToRooms.ContainsKey(gName))
                        {
                            groupToRooms[gName] = new List<string>();
                        }
                        groupToRooms[gName].Add(room.Name ?? $"Room {room.Id.IntegerValue}");
                    }
                }
            }
            
            result[elemId] = uniqueGroups.ToList();
            
            if (result[elemId].Count > 0)
            {
                _diagnostics.AppendLine($"Element {elemId.IntegerValue}: in {result[elemId].Count} group(s)");
                foreach (var grp in groupToRooms)
                {
                    _diagnostics.AppendLine($"  - Group '{grp.Key}' via room(s): {string.Join(", ", grp.Value.Distinct())}");
                }
            }
        }
        
        // Check if elements are direct members of groups
        int directMemberships = 0;
        foreach (Group group in allGroups)
        {
            HashSet<ElementId> memberIds = groupMemberCache[group];
            
            foreach (var kvp in elementToRoomsMap)
            {
                ElementId elemId = kvp.Key;
                
                if (memberIds.Contains(elemId))
                {
                    directMemberships++;
                    if (!result[elemId].Contains(group))
                    {
                        result[elemId].Add(group);
                        
                        GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
                        string gName = gt?.Name ?? "Unknown";
                        _diagnostics.AppendLine($"  Element {elemId.IntegerValue} is DIRECT member of group '{gName}'");
                    }
                }
            }
        }
        
        if (directMemberships > 0)
        {
            _diagnostics.AppendLine($"\nFound {directMemberships} direct element-to-group memberships");
        }
        
        return result;
    }
    
    // Optimized room bounding box check
    private bool IsRoomWithinBoundingBoxOptimized(RoomData roomData, XYZ bbMin, XYZ bbMax)
    {
        BoundingBoxXYZ roomBB = roomData.BoundingBox;
        if (roomBB == null) return false;
        
        return roomBB.Min.X >= bbMin.X && roomBB.Min.Y >= bbMin.Y && roomBB.Min.Z >= bbMin.Z &&
               roomBB.Max.X <= bbMax.X && roomBB.Max.Y <= bbMax.Y && roomBB.Max.Z <= bbMax.Z;
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
    
    // Prepare grid entries - Extracted to separate method for clarity
    private List<Dictionary<string, object>> PrepareGridEntries(
        List<Element> selectedElements,
        Dictionary<ElementId, List<Room>> elementToRoomsMap,
        Dictionary<ElementId, List<Group>> elementToGroupsMap,
        int maxGroupsPerElement,
        Document doc)
    {
        List<Dictionary<string, object>> gridEntries = new List<Dictionary<string, object>>();
        
        foreach (Element elem in selectedElements)
        {
            Dictionary<string, object> entry = new Dictionary<string, object>();
            
            // Store the actual element object for later retrieval
            entry["_Element"] = elem;
            
            // Type (Category)
            entry["Type"] = elem.Category?.Name ?? "Unknown";
            
            // Family
            ElementType elemType = doc.GetElement(elem.GetTypeId()) as ElementType;
            FamilySymbol famSymbol = elemType as FamilySymbol;
            if (famSymbol != null)
            {
                entry["Family"] = famSymbol.FamilyName ?? "Unknown";
            }
            else
            {
                entry["Family"] = elemType?.Name ?? "Unknown";
            }
            
            // Id
            entry["Id"] = elem.Id.IntegerValue.ToString();
            
            // Comments
            Parameter commentsParam = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            entry["Comments"] = commentsParam?.AsString() ?? "";
            
            // Room(s) containing this element
            List<Room> containingRooms = elementToRoomsMap[elem.Id];
            if (containingRooms.Count > 0)
            {
                List<string> roomNames = containingRooms.Select(r => r.Name ?? $"Room {r.Id.IntegerValue}").ToList();
                entry["Room(s)"] = string.Join(", ", roomNames);
            }
            else
            {
                entry["Room(s)"] = "None";
            }
            
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
        
        return gridEntries;
    }
    
    // Show diagnostics dialog
    private void ShowDiagnostics()
    {
        TaskDialog diagnosticDialog = new TaskDialog("Diagnostics Report");
        diagnosticDialog.MainInstruction = "Group Detection Diagnostics";
        diagnosticDialog.MainContent = "The diagnostic report has been generated. Click 'Show Report' to view it.";
        diagnosticDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Show Report");
        diagnosticDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Continue");
        
        TaskDialogResult dialogResult = diagnosticDialog.Show();
        
        if (dialogResult == TaskDialogResult.CommandLink1)
        {
            // Show the diagnostics in a form that can be copied
            System.Windows.Forms.Form form = new System.Windows.Forms.Form();
            form.Text = "Diagnostic Report - Copy All Text";
            form.Size = new System.Drawing.Size(900, 700);
            
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            textBox.Multiline = true;
            textBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            textBox.ReadOnly = true;
            textBox.Font = new System.Drawing.Font("Consolas", 9);
            textBox.Text = _diagnostics.ToString();
            textBox.Dock = System.Windows.Forms.DockStyle.Fill;
            textBox.WordWrap = false;
            
            form.Controls.Add(textBox);
            form.ShowDialog();
        }
    }
}
