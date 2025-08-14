using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
    private bool enableDiagnostics = true;
    private bool saveDiagnosticsToFile = true;
    private StringBuilder diagnosticLog = new StringBuilder();

    private void InitializeDiagnostics()
    {
        if (!enableDiagnostics) return;
        diagnosticLog.Clear();
        diagnosticLog.AppendLine($"=== ELEMENT-TO-GROUP MAPPING TEST - {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
        diagnosticLog.AppendLine();
    }

    private void LogSelectedElements(List<Element> selectedElements)
    {
        if (!enableDiagnostics) return;

        foreach (Element elem in selectedElements)
        {
            string elemDesc = elem.Name ?? $"{elem.Category?.Name}";
            List<XYZ> testPoints = GetElementTestPoints(elem);
            if (testPoints.Count > 0)
            {
                XYZ point = testPoints[0];
                diagnosticLog.AppendLine($"TESTING ELEMENT: {elemDesc} (ID: {elem.Id})");
                diagnosticLog.AppendLine($"Position: ({point.X:F2}, {point.Y:F2}, {point.Z:F2})");
                diagnosticLog.AppendLine();
                
                TestRoomContainmentMapping(elem, point);
            }
        }
    }

    private void TestRoomContainmentMapping(Element element, XYZ point)
    {
    Document doc = element.Document;
    
    // ENSURE CACHE IS BUILT
    if (_roomDataCache == null || _roomDataCache.Count == 0)
    {
        diagnosticLog.AppendLine(">>> Building room cache...");
        BuildRoomCache(doc);
        diagnosticLog.AppendLine($">>> Cache built with {_roomDataCache.Count} rooms");
    }
        
        // Step 1: Find which rooms contain this element with F2F logic
        diagnosticLog.AppendLine("STEP 1: ROOMS CONTAINING ELEMENT (with F2F height):");
        List<Room> containingRooms = new List<Room>();
        
        foreach (var kvp in _roomDataCache)
        {
            Room room = doc.GetElement(kvp.Key) as Room;
            if (room == null || room.Area <= 0) continue;
            
            RoomData roomData = kvp.Value;
            
            // Check XY containment
            bool inRoomXY = false;
            double testZ = roomData.LevelZ + 1.0;
            XYZ testPointXY = new XYZ(point.X, point.Y, testZ);
            if (room.IsPointInRoom(testPointXY))
            {
                inRoomXY = true;
            }
            
            if (inRoomXY)
            {
                // Check with actual F2F height
                double actualF2F = GetActualFloorToFloorHeightAtLocation(roomData.Level, doc, point);
                double actualMaxZ = roomData.LevelZ + actualF2F;
                
                if (point.Z >= roomData.LevelZ - 1.0 && point.Z <= actualMaxZ + 1.0)
                {
                    containingRooms.Add(room);
                    diagnosticLog.AppendLine($"  ✓ Room {room.Number} - {room.Name} (ID: {room.Id})");
                    diagnosticLog.AppendLine($"    Level: {roomData.Level?.Name}, F2F: {actualF2F:F2}ft");
                }
            }
        }
        
        if (containingRooms.Count == 0)
        {
            diagnosticLog.AppendLine("  ✗ NO ROOMS FOUND");
            return;
        }
        
        // Step 2: Check room-to-group mapping
        diagnosticLog.AppendLine("\nSTEP 2: ROOM-TO-GROUP MAPPING:");
        
        // Build room-to-group map (simplified version)
        Dictionary<ElementId, Group> roomToGroupMap = BuildRoomToGroupMapping(
            GetSpatiallyRelevantGroups(GetOverallBoundingBox(new List<Element> { element })), 
            doc);
        
        foreach (Room room in containingRooms)
        {
            if (roomToGroupMap.ContainsKey(room.Id))
            {
                Group mappedGroup = roomToGroupMap[room.Id];
                GroupType gt = doc.GetElement(mappedGroup.GetTypeId()) as GroupType;
                diagnosticLog.AppendLine($"  ✓ Room {room.Number} → Group {gt?.Name} (ID: {mappedGroup.Id})");
                
                // Check if this is direct membership or spatial containment
                bool isDirectMember = mappedGroup.GetMemberIds().Contains(room.Id);
                diagnosticLog.AppendLine($"    Membership type: {(isDirectMember ? "DIRECT" : "SPATIAL")}");
                
                // Count instances of this group type
                FilteredElementCollector groupCollector = new FilteredElementCollector(doc);
                int instanceCount = groupCollector
                    .OfClass(typeof(Group))
                    .WhereElementIsNotElementType()
                    .Where(g => g.GetTypeId() == mappedGroup.GetTypeId())
                    .Count();
                diagnosticLog.AppendLine($"    Group type instances: {instanceCount}");
            }
            else
            {
                diagnosticLog.AppendLine($"  ✗ Room {room.Number} NOT MAPPED TO ANY GROUP!");
                
                // Investigate why
                diagnosticLog.AppendLine($"    Checking groups containing this room...");
                bool foundInAnyGroup = false;
                
                FilteredElementCollector groupCollector = new FilteredElementCollector(doc);
                foreach (Group group in groupCollector.OfClass(typeof(Group)).Cast<Group>())
                {
                    if (group.GetMemberIds().Contains(room.Id))
                    {
                        GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
                        diagnosticLog.AppendLine($"    Found as direct member in: {gt?.Name} (ID: {group.Id})");
                        foundInAnyGroup = true;
                    }
                }
                
                if (!foundInAnyGroup)
                {
                    diagnosticLog.AppendLine($"    Room is NOT a direct member of any group");
                    diagnosticLog.AppendLine($"    Spatial containment validation may have failed");
                }
            }
        }
        
        // Step 3: Test the actual element containment check
        diagnosticLog.AppendLine("\nSTEP 3: ELEMENT CONTAINMENT CHECK:");
        
        foreach (var kvp in roomToGroupMap)
        {
            Room room = doc.GetElement(kvp.Key) as Room;
            if (!containingRooms.Contains(room)) continue;
            
            Group group = kvp.Value;
            
            // Test with the actual IsElementInRoomOptimized method
            List<XYZ> testPoints = new List<XYZ> { point };
            bool isContained = IsElementInRoomOptimized(element, room, testPoints, point.Z - 5, point.Z + 5);
            
            GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
            diagnosticLog.AppendLine($"  Room {room.Number} in Group {gt?.Name}:");
            diagnosticLog.AppendLine($"    IsElementInRoomOptimized result: {(isContained ? "✓ YES" : "✗ NO")}");
            
            if (!isContained)
            {
                diagnosticLog.AppendLine($"    >>> MISMATCH: Element should be in room but check failed!");
            }
        }
    }

    // Simplified version that matches diagnostic logic exactly
    private double GetActualFloorToFloorHeightAtLocation(Level currentLevel, Document doc, XYZ testPoint)
    {
        if (currentLevel == null || testPoint == null) return 10.0;

        FilteredElementCollector floorCollector = new FilteredElementCollector(doc);
        List<Floor> allFloors = floorCollector
            .OfClass(typeof(Floor))
            .WhereElementIsNotElementType()
            .Cast<Floor>()
            .Where(f => f != null)
            .ToList();

        List<(Floor floor, double elevation)> floorsAtLocation = new List<(Floor, double)>();
        
        foreach (Floor floor in allFloors)
        {
            BoundingBoxXYZ bb = floor.get_BoundingBox(null);
            if (bb == null) continue;
            
            if (testPoint.X >= bb.Min.X - 1.0 && testPoint.X <= bb.Max.X + 1.0 &&
                testPoint.Y >= bb.Min.Y - 1.0 && testPoint.Y <= bb.Max.Y + 1.0)
            {
                floorsAtLocation.Add((floor, bb.Max.Z));
            }
        }
        
        floorsAtLocation = floorsAtLocation.OrderBy(f => f.elevation).ToList();
        double roomLevelElevation = currentLevel.Elevation;
        
        var floorBelow = floorsAtLocation
            .Where(f => Math.Abs(f.elevation - roomLevelElevation) < 2.0)
            .OrderBy(f => Math.Abs(f.elevation - roomLevelElevation))
            .FirstOrDefault();
            
        if (floorBelow.floor != null)
        {
            var floorAbove = floorsAtLocation
                .Where(f => f.elevation > floorBelow.elevation + 1.0)
                .OrderBy(f => f.elevation)
                .FirstOrDefault();
                
            if (floorAbove.floor != null)
            {
                return floorAbove.elevation - floorBelow.elevation;
            }
        }
        
        return 10.0; // Default
    }

    // Minimal stub implementations
    private void LogElementRoomCheck(Element element, Room room, bool isContained,
        List<XYZ> testPoints, double groupMinZ, double groupMaxZ, Document doc) { }

    private void LogFloorToFloorCheck(XYZ point, RoomData roomData, Document doc) { }

    private void LogRoomValidationFailure(Room room, Group candidateGroup,
        List<Group> sameTypeGroups, DetailedValidationResult validationResult, Document doc) { }

    private void LogRoomValidationSuccess(Room room, Group candidateGroup,
        DetailedValidationResult validationResult, Document doc) { }

    private void LogElementContainmentCheck(Element element, Group group, bool isContained, string roomName, Document doc) { }

    private void LogNearbyRoomsForElement(Element element, Document doc) { }

    private void DiagnoseElementsNotInGroups(List<Element> selectedElements,
        Dictionary<ElementId, List<Group>> elementsInGroups, Document doc) { }

    private void TrackStat(string statName) { }

    private void TestFloorDetection(Document doc, XYZ testPoint) { }

    private void TestRoomContainmentWithF2F(Element element, XYZ point) { }

    private void LogFinalSummary(int selectedCount, int foundInGroups, int totalCopied)
    {
        if (!enableDiagnostics) return;
        diagnosticLog.AppendLine($"\n=== SUMMARY ===");
        diagnosticLog.AppendLine($"Elements selected: {selectedCount}");
        diagnosticLog.AppendLine($"Elements in groups: {foundInGroups}");
        diagnosticLog.AppendLine($"Elements copied: {totalCopied}");
    }

    private void SaveDiagnostics(bool forceWrite = false)
    {
        if (!enableDiagnostics || !saveDiagnosticsToFile) return;

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string fileName = $"RevitMappingTest_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
        string filePath = System.IO.Path.Combine(desktopPath, fileName);

        try
        {
            System.IO.File.WriteAllText(filePath, diagnosticLog.ToString());
            if (forceWrite)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Diagnostics Saved", $"Saved to:\n{filePath}");
            }
        }
        catch { }
    }
}
