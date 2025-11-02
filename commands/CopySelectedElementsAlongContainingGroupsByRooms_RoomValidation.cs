using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
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

    // Simplified room similarity check diagnostic
    private string GetRoomMismatchReason(Room testRoom, XYZ expectedLocation,
        RoomMappingInfo foundRoom, double distance, double areaDifference)
    {
        if (distance > ROOM_POSITION_TOLERANCE_RELAXED)
        {
            return $"Too far: {distance:F1}ft (max: {ROOM_POSITION_TOLERANCE_RELAXED}ft)";
        }
        else if (distance > ROOM_POSITION_TOLERANCE)
        {
            if (areaDifference > ROOM_AREA_TOLERANCE_RELAXED)
            {
                return $"Position OK with relaxed tolerance, but area diff too large: {areaDifference:P0}";
            }
            return $"Distance {distance:F1}ft exceeds strict tolerance {ROOM_POSITION_TOLERANCE}ft";
        }
        else if (areaDifference > ROOM_AREA_TOLERANCE)
        {
            return $"Area mismatch: {areaDifference:P0} (max: {ROOM_AREA_TOLERANCE:P0})";
        }
        else if (!string.IsNullOrEmpty(testRoom.Name) && !string.IsNullOrEmpty(foundRoom.RoomName)
                 && !foundRoom.RoomName.Equals(testRoom.Name, StringComparison.OrdinalIgnoreCase))
        {
            return $"Name mismatch: '{foundRoom.RoomName}' vs '{testRoom.Name}'";
        }

        return "Unknown reason";
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

    // Show results dialog
    private void ShowResults(Document doc, List<Element> selectedElements,
                            Dictionary<ElementId, List<Group>> elementsInGroups,
                            Dictionary<ElementId, List<Group>> groupsByType,
                            Dictionary<ElementId, List<string>> elementToGroupTypeNames,
                            List<Group> spatiallyRelevantGroups,
                            IList<Group> allGroups,
                            CopyResult copyResult)
    {
        StringBuilder resultMessage = new StringBuilder();

        resultMessage.AppendLine($"=== EXECUTION COMPLETE ===");
        resultMessage.AppendLine();

        if (copyResult.TotalCopied == 0)
        {
            resultMessage.AppendLine($"No elements were copied.");
            resultMessage.AppendLine($"\nDiagnostic Information:");
            resultMessage.AppendLine($"- Selected elements: {selectedElements.Count}");
            resultMessage.AppendLine($"- Elements in groups: {elementsInGroups.Count}");
            resultMessage.AppendLine($"- Group types found: {groupsByType.Count}");
            resultMessage.AppendLine($"- Group types processed: {copyResult.GroupTypesProcessed}");
            resultMessage.AppendLine($"- Group types skipped (no reference elements): {copyResult.GroupTypesSkippedNoRefElements}");

            // Room validation info
            resultMessage.AppendLine($"\nRoom Validation:");
            resultMessage.AppendLine($"- Rooms validated by similarity: {roomsValidatedBySimilarity}");
            resultMessage.AppendLine($"- Rooms invalidated: {roomsInvalidatedByDissimilarity}");

            resultMessage.AppendLine($"\n[Check the diagnostic log file on your desktop for detailed information]");
        }
        else
        {
            resultMessage.AppendLine($"Copied {copyResult.TotalCopied} elements to group instances.");
            resultMessage.AppendLine($"\nGroup types processed: {copyResult.GroupTypesProcessed}");
            resultMessage.AppendLine($"Total copy operations: {copyResult.TotalCopyOperations}");
            resultMessage.AppendLine($"Groups checked: {spatiallyRelevantGroups.Count} of {allGroups.Count}");

            // Show room validation statistics
            resultMessage.AppendLine($"\n=== ROOM STATISTICS ===");
            resultMessage.AppendLine($"Rooms as direct group members: {roomsAsDirectMembers}");
            resultMessage.AppendLine($"Rooms validated by similarity: {roomsValidatedBySimilarity}");
            resultMessage.AppendLine($"Rooms invalidated: {roomsInvalidatedByDissimilarity}");

            // Show which elements were in multiple groups
            var multiGroupElements = elementToGroupTypeNames.Where(kvp => kvp.Value.Count > 1);
            if (multiGroupElements.Any())
            {
                resultMessage.AppendLine("\nElements in multiple groups:");
                foreach (var kvp in multiGroupElements)
                {
                    Element elem = doc.GetElement(kvp.Key);
                   string elemDesc = elem.Name ?? $"{elem.Category?.Name} {elem.Id}";
                   resultMessage.AppendLine($"- {elemDesc}: {string.Join(", ", kvp.Value)}");
               }
           }
       }

       TaskDialog.Show("Copy Complete", resultMessage.ToString());
   }
}
