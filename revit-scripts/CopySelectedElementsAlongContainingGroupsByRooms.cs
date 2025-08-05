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
public class CopySelectedElementsAlongContainingGroupsByRooms : IExternalCommand
{
    // ===== CONFIGURATION OPTIONS =====
    // Set to true to check for existing elements before copying (can cause false positives)
    private bool enableDuplicateCheck = false; 
    
    // Set to true to see detailed information about what the duplicate check finds
    private bool verboseDuplicateCheck = true; 
    
    // Note: The duplicate check is DISABLED by default because it can incorrectly
    // detect "duplicates" when elements of the same type are close together.
    // Only enable if you're getting actual duplicate elements after copying.
    // =================================
    
    // Cache for element test points to avoid recalculation
    private Dictionary<ElementId, List<XYZ>> _elementTestPointsCache = new Dictionary<ElementId, List<XYZ>>();
    
    // Cache for room data
    private Dictionary<ElementId, RoomData> _roomDataCache = new Dictionary<ElementId, RoomData>();
    
    private class RoomData
    {
        public double LevelZ { get; set; }
        public bool IsUnbound { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
    }
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        try
        {
            // Get selected elements
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            
            if (selectedIds.Count == 0)
            {
                message = "Please select elements first";
                return Result.Failed;
            }
            
            // Get selected elements (exclude groups and element types)
            List<Element> selectedElements = new List<Element>();
            
            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem != null && !(elem is Group) && !(elem is ElementType))
                {
                    selectedElements.Add(elem);
                }
            }
            
            if (selectedElements.Count == 0)
            {
                message = "No valid elements found in selection. Please select elements to copy.";
                return Result.Failed;
            }
            
            // Pre-calculate test points for all selected elements
            foreach (Element elem in selectedElements)
            {
                _elementTestPointsCache[elem.Id] = GetElementTestPoints(elem);
            }
            
            // Create bounding box for all selected elements for quick filtering
            BoundingBoxXYZ overallBB = GetOverallBoundingBox(selectedElements);
            if (overallBB == null)
            {
                message = "Could not determine bounding box of selected elements";
                return Result.Failed;
            }
            
            // Find all groups in the document
            FilteredElementCollector groupCollector = new FilteredElementCollector(doc);
            IList<Group> allGroups = groupCollector
                .OfClass(typeof(Group))
                .WhereElementIsNotElementType()
                .Cast<Group>()
                .ToList();
            
            // Dictionary to store which groups contain which elements
            Dictionary<ElementId, List<Group>> elementToGroupsMap = new Dictionary<ElementId, List<Group>>();
            
            // Initialize the map
            foreach (Element elem in selectedElements)
            {
                elementToGroupsMap[elem.Id] = new List<Group>();
            }
            
            // Process each group with spatial filtering
            foreach (Group group in allGroups)
            {
                // Quick bounding box check first
                BoundingBoxXYZ groupBB = GetGroupBoundingBox(group, doc);
                if (groupBB == null || !BoundingBoxesIntersect(overallBB, groupBB))
                {
                    continue; // Skip this group - bounding boxes don't intersect
                }
                
                // Check which selected elements are contained in this group's rooms
                List<Element> containedElements = GetElementsContainedInGroupRoomsOptimized(group, selectedElements, doc);
                
                // Update the map
                foreach (Element elem in containedElements)
                {
                    elementToGroupsMap[elem.Id].Add(group);
                }
            }
            
            // Filter out elements that aren't in any groups
            Dictionary<ElementId, List<Group>> elementsInGroups = elementToGroupsMap
                .Where(kvp => kvp.Value.Count > 0)
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            
            if (elementsInGroups.Count == 0)
            {
                message = "No selected elements are contained within rooms of any groups";
                return Result.Failed;
            }
            
            // Build element to group types mapping for Comments parameter
            Dictionary<ElementId, List<string>> elementToGroupTypeNames = new Dictionary<ElementId, List<string>>();
            foreach (var kvp in elementsInGroups)
            {
                ElementId elementId = kvp.Key;
                List<Group> containingGroups = kvp.Value;
                
                List<string> groupTypeNames = new List<string>();
                HashSet<string> uniqueNames = new HashSet<string>();
                
                foreach (Group group in containingGroups)
                {
                    GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
                    if (gt != null && uniqueNames.Add(gt.Name))
                    {
                        groupTypeNames.Add(gt.Name);
                    }
                }
                
                elementToGroupTypeNames[elementId] = groupTypeNames;
            }
            
            // Group the containing groups by their group type
            Dictionary<ElementId, List<Group>> groupsByType = new Dictionary<ElementId, List<Group>>();
            HashSet<Group> allContainingGroups = new HashSet<Group>();
            
            foreach (var kvp in elementsInGroups)
            {
                foreach (Group group in kvp.Value)
                {
                    allContainingGroups.Add(group);
                    ElementId typeId = group.GetTypeId();
                    if (!groupsByType.ContainsKey(typeId))
                    {
                        groupsByType[typeId] = new List<Group>();
                    }
                    if (!groupsByType[typeId].Contains(group))
                    {
                        groupsByType[typeId].Add(group);
                    }
                }
            }
            
            // Process each group type and copy elements
            int totalCopied = 0;
            int groupTypesProcessed = 0;
            int groupTypesSkippedNoRefElements = 0;
            Dictionary<string, List<string>> diagnosticInfo = new Dictionary<string, List<string>>();
            
            using (Transaction trans = new Transaction(doc, "Copy Elements Following Containing Groups By Rooms"))
            {
                trans.Start();
                
                // First, update Comments parameter for all selected elements that are in groups
                foreach (var elemKvp in elementToGroupTypeNames)
                {
                    Element elem = doc.GetElement(elemKvp.Key);
                    if (elem != null)
                    {
                        Parameter commentsParam = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                        if (commentsParam != null && !commentsParam.IsReadOnly)
                        {
                            // Set to all group types that contain this element
                            string groupNames = string.Join(", ", elemKvp.Value);
                            commentsParam.Set(groupNames);
                        }
                    }
                }
                
                // Now process each group type and copy elements
                foreach (var kvp in groupsByType)
                {
                    ElementId groupTypeId = kvp.Key;
                    List<Group> containingGroupsOfThisType = kvp.Value;
                    
                    GroupType groupType = doc.GetElement(groupTypeId) as GroupType;
                    if (groupType == null) continue;
                    
                    // Use the first containing group as reference
                    Group referenceGroup = containingGroupsOfThisType.First();
                    
                    // Get all instances of this group type
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    IList<Group> allGroupInstances = collector
                        .OfClass(typeof(Group))
                        .WhereElementIsNotElementType()
                        .Where(g => g.GetTypeId() == groupTypeId)
                        .Cast<Group>()
                        .ToList();
                    
                    if (allGroupInstances.Count < 2) continue; // Need at least 2 instances to copy
                    
                    // Initialize diagnostics for this group type
                    string groupTypeName = groupType.Name ?? "Unknown";
                    List<string> groupDiagnostics = new List<string>();
                    groupDiagnostics.Add($"Group type: {groupTypeName}");
                    groupDiagnostics.Add($"Total instances found: {allGroupInstances.Count}");
                    groupDiagnostics.Add($"Reference group ID: {referenceGroup.Id}");
                    
                    // Determine which elements to copy from this reference group
                    List<Element> elementsToCopy = new List<Element>();
                    foreach (var elemKvp in elementsInGroups)
                    {
                        Element elem = doc.GetElement(elemKvp.Key);
                        if (elemKvp.Value.Contains(referenceGroup))
                        {
                            elementsToCopy.Add(elem);
                        }
                    }
                    
                    if (elementsToCopy.Count == 0) continue;
                    
                    groupTypesProcessed++;
                    
                    // Get reference elements for transformation calculation
                    ReferenceElements refElements = GetReferenceElements(referenceGroup, doc);
                    
                    if (refElements == null || refElements.Elements.Count < 2)
                    {
                        groupTypesSkippedNoRefElements++;
                        continue;
                    }
                    
                    // Prepare element IDs for batch copying
                    List<ElementId> elementIdsToCopy = elementsToCopy.Select(e => e.Id).ToList();
                    
                    groupDiagnostics.Add($"Elements to copy: {elementIdsToCopy.Count}");
                    groupDiagnostics.Add($"Duplicate check enabled: {enableDuplicateCheck}");
                    
                    int instancesCopiedTo = 0;
                    int instancesSkippedSameAsRef = 0;
                    int instancesSkippedNoTransform = 0;
                    int instancesSkippedDuplicates = 0;
                    int instancesSkippedError = 0;
                    
                    // Copy elements to other instances
                    foreach (Group otherGroup in allGroupInstances)
                    {
                        if (otherGroup.Id == referenceGroup.Id) 
                        {
                            instancesSkippedSameAsRef++;
                            continue;
                        }
                        
                        XYZ otherOrigin = (otherGroup.Location as LocationPoint).Point;
                        
                        TransformResult transformResult = CalculateTransformation(refElements, otherGroup, doc);
                        
                        if (transformResult != null)
                        {
                            // Check if elements already exist at target location (if enabled)
                            bool skipDueToDuplicates = false;
                            string duplicateDetails = "";
                            
                            if (enableDuplicateCheck)
                            {
                                var duplicateCheckResult = CheckIfAnyElementExistsAtTarget(elementsToCopy[0], transformResult, 
                                    refElements.GroupOrigin, otherOrigin, doc);
                                skipDueToDuplicates = duplicateCheckResult.HasDuplicates;
                                duplicateDetails = duplicateCheckResult.Details;
                            }
                            
                            if (skipDueToDuplicates)
                            {
                                instancesSkippedDuplicates++;
                                groupDiagnostics.Add($"  - Instance {otherGroup.Id}: SKIPPED (duplicates detected{(verboseDuplicateCheck ? ": " + duplicateDetails : "")})");
                                continue; // Skip this group - elements likely already exist
                            }
                            
                            // Create the transformation
                            Transform transform = CreateTransform(transformResult, 
                                refElements.GroupOrigin, otherOrigin);
                            
                            try
                            {
                                // BATCH COPY - Copy all elements at once
                                ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                    doc, 
                                    elementIdsToCopy,
                                    doc,
                                    transform,
                                    null);
                                
                                // Get the level of the target group
                                Level targetGroupLevel = GetGroupLevel(otherGroup, doc);
                                
                                // Update properties for all copied elements in batch
                                if (copiedIds.Count > 0)
                                {
                                    foreach (ElementId copiedId in copiedIds)
                                    {
                                        Element copiedElem = doc.GetElement(copiedId);
                                        
                                        // Find the original element to get its group names
                                        ElementId originalId = ElementId.InvalidElementId;
                                        for (int i = 0; i < elementIdsToCopy.Count; i++)
                                        {
                                            if (i < copiedIds.Count && copiedIds.ElementAt(i) == copiedId)
                                            {
                                                originalId = elementIdsToCopy[i];
                                                break;
                                            }
                                        }
                                        
                                        // Update Comments parameter with all containing group types
                                        Parameter commentsParam = copiedElem.get_Parameter(
                                            BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                        if (commentsParam != null && !commentsParam.IsReadOnly && 
                                            originalId != ElementId.InvalidElementId && 
                                            elementToGroupTypeNames.ContainsKey(originalId))
                                        {
                                            string groupNames = string.Join(", ", elementToGroupTypeNames[originalId]);
                                            commentsParam.Set(groupNames);
                                        }
                                        
                                        // Update level if needed
                                        if (targetGroupLevel != null)
                                        {
                                            UpdateElementLevel(copiedElem, targetGroupLevel);
                                        }
                                    }
                                    
                                    totalCopied += copiedIds.Count;
                                    instancesCopiedTo++;
                                    groupDiagnostics.Add($"  - Instance {otherGroup.Id}: SUCCESS (copied {copiedIds.Count} elements)");
                                }
                            }
                            catch (Exception ex)
                            {
                                instancesSkippedError++;
                                groupDiagnostics.Add($"  - Instance {otherGroup.Id}: ERROR ({ex.Message})");
                                // Silent fail - continue with next group
                            }
                        }
                        else
                        {
                            instancesSkippedNoTransform++;
                            groupDiagnostics.Add($"  - Instance {otherGroup.Id}: SKIPPED (no transformation found)");
                        }
                    }
                    
                    // Add summary for this group type
                    groupDiagnostics.Add($"Summary for {groupTypeName}:");
                    groupDiagnostics.Add($"  - Instances processed: {allGroupInstances.Count}");
                    groupDiagnostics.Add($"  - Reference instance (skipped): {instancesSkippedSameAsRef}");
                    groupDiagnostics.Add($"  - Successfully copied to: {instancesCopiedTo}");
                    groupDiagnostics.Add($"  - Skipped (duplicates): {instancesSkippedDuplicates}");
                    groupDiagnostics.Add($"  - Skipped (no transform): {instancesSkippedNoTransform}");
                    groupDiagnostics.Add($"  - Skipped (errors): {instancesSkippedError}");
                    
                    diagnosticInfo[groupTypeName] = groupDiagnostics;
                }
                
                trans.Commit();
            }
            
            // Build informative message
            StringBuilder resultMessage = new StringBuilder();
            
            if (totalCopied == 0)
            {
                resultMessage.AppendLine($"No elements were copied.");
                resultMessage.AppendLine($"\nDiagnostic Information:");
                resultMessage.AppendLine($"- Selected elements: {selectedElements.Count}");
                resultMessage.AppendLine($"- Elements in groups: {elementsInGroups.Count}");
                resultMessage.AppendLine($"- Group types found: {groupsByType.Count}");
                resultMessage.AppendLine($"- Group types processed: {groupTypesProcessed}");
                resultMessage.AppendLine($"- Group types skipped (no reference elements): {groupTypesSkippedNoRefElements}");
                
                resultMessage.AppendLine($"\nConfiguration:");
                resultMessage.AppendLine($"- Duplicate check: {(enableDuplicateCheck ? "ENABLED" : "DISABLED")}");
                if (enableDuplicateCheck)
                {
                    resultMessage.AppendLine($"- Verbose duplicate check: {(verboseDuplicateCheck ? "YES" : "NO")}");
                }
                
                if (groupsByType.Count > 0)
                {
                    resultMessage.AppendLine("\nGroup types and instance counts:");
                    foreach (var typeKvp in groupsByType)
                    {
                        GroupType gt = doc.GetElement(typeKvp.Key) as GroupType;
                        FilteredElementCollector col = new FilteredElementCollector(doc);
                        int instanceCount = col
                            .OfClass(typeof(Group))
                            .WhereElementIsNotElementType()
                            .Where(g => g.GetTypeId() == typeKvp.Key)
                            .Count();
                        resultMessage.AppendLine($"- {gt?.Name ?? "Unknown"}: {instanceCount} instances" + 
                            (instanceCount < 2 ? " (SKIPPED - needs at least 2 instances to copy)" : ""));
                    }
                }
                
                // Show detailed diagnostics even when nothing was copied
                if (diagnosticInfo.Count > 0)
                {
                    resultMessage.AppendLine("\n=== DETAILED DIAGNOSTICS ===");
                    foreach (var kvp in diagnosticInfo)
                    {
                        resultMessage.AppendLine($"\n{kvp.Key}:");
                        foreach (string line in kvp.Value)
                        {
                            resultMessage.AppendLine(line);
                        }
                    }
                }
            }
            else
            {
                resultMessage.AppendLine($"Copied {totalCopied} elements to group instances.");
                resultMessage.AppendLine($"\nGroup types processed: {groupTypesProcessed}");
                
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
            
            // Note: If some instances were missed, you can run the command again
            // with elements from those instances selected.
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
    
    // Pre-calculate test points for an element
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
            // For curves, check endpoints
            Curve curve = locCurve.Curve;
            testPoints.Add(curve.GetEndPoint(0));
            testPoints.Add(curve.GetEndPoint(1));
        }
        else
        {
            // For other elements, use bounding box center
            BoundingBoxXYZ bb = element.get_BoundingBox(null);
            if (bb != null)
            {
                testPoints.Add((bb.Min + bb.Max) * 0.5);
            }
        }
        
        return testPoints;
    }
    
    // Get overall bounding box for a list of elements
    private BoundingBoxXYZ GetOverallBoundingBox(List<Element> elements)
    {
        if (elements.Count == 0) return null;
        
        double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
        
        bool hasValidBB = false;
        
        foreach (Element elem in elements)
        {
            BoundingBoxXYZ bb = elem.get_BoundingBox(null);
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
        
        if (!hasValidBB) return null;
        
        BoundingBoxXYZ overallBB = new BoundingBoxXYZ();
        overallBB.Min = new XYZ(minX, minY, minZ);
        overallBB.Max = new XYZ(maxX, maxY, maxZ);
        
        return overallBB;
    }
    
    // Get bounding box for a group
    private BoundingBoxXYZ GetGroupBoundingBox(Group group, Document doc)
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
    
    // Check if two bounding boxes intersect
    private bool BoundingBoxesIntersect(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        return !(bb1.Max.X < bb2.Min.X || bb2.Max.X < bb1.Min.X ||
                 bb1.Max.Y < bb2.Min.Y || bb2.Max.Y < bb1.Min.Y ||
                 bb1.Max.Z < bb2.Min.Z || bb2.Max.Z < bb1.Min.Z);
    }
    
    // Optimized version that uses cached data
    private List<Element> GetElementsContainedInGroupRoomsOptimized(Group group, List<Element> candidateElements, Document doc)
    {
        List<Element> containedElements = new List<Element>();
        
        // Get all rooms that are members of this group
        List<Room> groupRooms = new List<Room>();
        ICollection<ElementId> memberIds = group.GetMemberIds();
        
        // Also get the overall height range of the group
        double groupMinZ = double.MaxValue;
        double groupMaxZ = double.MinValue;
        
        foreach (ElementId id in memberIds)
        {
            Element member = doc.GetElement(id);
            if (member is Room)
            {
                Room room = member as Room;
                groupRooms.Add(room);
                
                // Cache room data if not already cached
                if (!_roomDataCache.ContainsKey(room.Id))
                {
                    RoomData roomData = new RoomData();
                    roomData.IsUnbound = room.Volume <= 0.001;
                    roomData.BoundingBox = room.get_BoundingBox(null);
                    
                    if (room.LevelId != null && room.LevelId != ElementId.InvalidElementId)
                    {
                        Level roomLevel = doc.GetElement(room.LevelId) as Level;
                        if (roomLevel != null)
                        {
                            roomData.LevelZ = roomLevel.Elevation;
                        }
                    }
                    
                    _roomDataCache[room.Id] = roomData;
                }
            }
            
            // Update group height bounds
            BoundingBoxXYZ bb = member.get_BoundingBox(null);
            if (bb != null)
            {
                groupMinZ = Math.Min(groupMinZ, bb.Min.Z);
                groupMaxZ = Math.Max(groupMaxZ, bb.Max.Z);
            }
        }
        
        if (groupRooms.Count == 0)
        {
            return containedElements;
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
    
    private class DuplicateCheckResult
    {
        public bool HasDuplicates { get; set; }
        public string Details { get; set; }
    }
    
    // More precise duplicate check with detailed results
    private DuplicateCheckResult CheckIfAnyElementExistsAtTarget(Element sampleElement, 
        TransformResult transformResult, XYZ refOrigin, XYZ targetOrigin, Document doc)
    {
        var result = new DuplicateCheckResult { HasDuplicates = false, Details = "" };
        
        Transform transform = CreateTransform(transformResult, refOrigin, targetOrigin);
        
        // Get element location
        XYZ testPoint = null;
        LocationPoint locPoint = sampleElement.Location as LocationPoint;
        LocationCurve locCurve = sampleElement.Location as LocationCurve;
        
        if (locPoint != null)
        {
            testPoint = locPoint.Point;
        }
        else if (locCurve != null)
        {
            testPoint = locCurve.Curve.GetEndPoint(0);
        }
        
        if (testPoint == null) 
        {
            result.Details = "no test point";
            return result;
        }
        
        // Transform the test point
        XYZ targetPoint = transform.OfPoint(testPoint);
        
        // Use very small search radius for precise check
        double searchRadius = 0.01; // 0.01 feet = ~3mm - very precise
        
        // Create a bounding box filter
        Outline outline = new Outline(
            targetPoint - new XYZ(searchRadius, searchRadius, searchRadius),
            targetPoint + new XYZ(searchRadius, searchRadius, searchRadius));
        BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);
        
        // Check for elements of same category and type
        FilteredElementCollector collector = new FilteredElementCollector(doc);
        var foundElements = collector
            .WhereElementIsNotElementType()
            .OfCategoryId(sampleElement.Category.Id)
            .WherePasses(bbFilter)
            .ToList();
        
        // Filter to elements of the same type, excluding the source element
        var sameTypeElements = foundElements
            .Where(e => e.GetTypeId() == sampleElement.GetTypeId() && e.Id != sampleElement.Id)
            .ToList();
        
        if (sameTypeElements.Count > 0)
        {
            result.HasDuplicates = true;
            result.Details = $"found {sameTypeElements.Count} elements of same type at target location";
            
            // For curves, do additional endpoint check
            if (locCurve != null)
            {
                bool exactMatch = false;
                foreach (var elem in sameTypeElements)
                {
                    if (elem.Location is LocationCurve existingCurve)
                    {
                        XYZ existingStart = existingCurve.Curve.GetEndPoint(0);
                        XYZ existingEnd = existingCurve.Curve.GetEndPoint(1);
                        
                        // Transform both endpoints of source curve
                        XYZ transformedStart = transform.OfPoint(locCurve.Curve.GetEndPoint(0));
                        XYZ transformedEnd = transform.OfPoint(locCurve.Curve.GetEndPoint(1));
                        
                        // Check if endpoints match (in either order)
                        bool endpointsMatch = 
                            (existingStart.IsAlmostEqualTo(transformedStart, 0.01) && 
                             existingEnd.IsAlmostEqualTo(transformedEnd, 0.01)) ||
                            (existingStart.IsAlmostEqualTo(transformedEnd, 0.01) && 
                             existingEnd.IsAlmostEqualTo(transformedStart, 0.01));
                        
                        if (endpointsMatch)
                        {
                            exactMatch = true;
                            break;
                        }
                    }
                }
                
                if (!exactMatch)
                {
                    result.HasDuplicates = false;
                    result.Details = "found elements nearby but not exact curve match";
                }
            }
        }
        else
        {
            result.Details = $"no elements of same type found within {searchRadius} feet";
        }
        
        return result;
    }
    
    // Update element level with all possible parameters
    private void UpdateElementLevel(Element element, Level targetLevel)
    {
        // Try various level parameters
        Parameter[] levelParams = new Parameter[]
        {
            element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM),
            element.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM),
            element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM),
            element.get_Parameter(BuiltInParameter.LEVEL_PARAM),
            element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
        };
        
        foreach (Parameter param in levelParams)
        {
            if (param != null && !param.IsReadOnly)
            {
                try
                {
                    param.Set(targetLevel.Id);
                }
                catch
                {
                    // Silent fail - some parameters might not accept the value
                }
            }
        }
    }
    
    private Transform CreateTransform(TransformResult transformResult, XYZ refOrigin, XYZ targetOrigin)
    {
        // For mirrored transformations with no rotation
        if (transformResult.IsMirrored && Math.Abs(transformResult.Rotation) < 0.01)
        {
            // Determine which axis is mirrored based on the translation
            XYZ midpoint = (refOrigin + targetOrigin) * 0.5;
            
            // Check which coordinates changed
            double xDiff = Math.Abs(refOrigin.X - targetOrigin.X);
            double yDiff = Math.Abs(refOrigin.Y - targetOrigin.Y);
            double zDiff = Math.Abs(refOrigin.Z - targetOrigin.Z);
            
            Transform mirror;
            
            if (xDiff > yDiff && xDiff > zDiff)
            {
                // X coordinates differ most - mirror about Y-Z plane
                Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, midpoint);
                mirror = Transform.CreateReflection(mirrorPlane);
            }
            else if (yDiff > xDiff && yDiff > zDiff)
            {
                // Y coordinates differ most - mirror about X-Z plane
                Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisY, midpoint);
                mirror = Transform.CreateReflection(mirrorPlane);
            }
            else
            {
                // Z coordinates differ most - mirror about X-Y plane
                Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, midpoint);
                mirror = Transform.CreateReflection(mirrorPlane);
            }
            
            return mirror;
        }
        
        // For complex transformations involving both rotation and mirroring
        if (Math.Abs(Math.Abs(transformResult.Rotation) - 180.0) < 1.0 && transformResult.IsMirrored)
        {
            // This is equivalent to a mirror about Y axis
            Transform t = Transform.Identity;
            
            // Set the basis vectors
            t.BasisX = -XYZ.BasisX;  // (-1, 0, 0)
            t.BasisY = XYZ.BasisY;   // (0, 1, 0) 
            t.BasisZ = XYZ.BasisZ;   // (0, 0, 1)
            
            // Set the translation
            t.Origin = new XYZ(
                2 * ((refOrigin.X + targetOrigin.X) / 2) - refOrigin.X,
                targetOrigin.Y - refOrigin.Y,
                targetOrigin.Z - refOrigin.Z
            );
            
            return t;
        }
        
        // Standard case: build transformation step by step
        Transform transform = Transform.Identity;
        
        // Move to origin
        transform = transform.Multiply(Transform.CreateTranslation(-refOrigin));
        
        // Apply rotation
        if (Math.Abs(transformResult.Rotation) > 0.01)
        {
            double radians = transformResult.Rotation * Math.PI / 180;
            transform = transform.Multiply(Transform.CreateRotation(XYZ.BasisZ, radians));
        }
        
        // Apply mirror if needed
        if (transformResult.IsMirrored)
        {
            Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, XYZ.Zero);
            transform = transform.Multiply(Transform.CreateReflection(mirrorPlane));
        }
        
        // Move to target
        transform = transform.Multiply(Transform.CreateTranslation(targetOrigin));
        
        return transform;
    }
    
    // Get the level of a group based on its members or location
    private Level GetGroupLevel(Group group, Document doc)
    {
        // First, try to get level from group members (rooms are good indicators)
        ICollection<ElementId> memberIds = group.GetMemberIds();
        foreach (ElementId id in memberIds)
        {
            Element member = doc.GetElement(id);
            if (member is Room)
            {
                Room room = member as Room;
                if (room.LevelId != null && room.LevelId != ElementId.InvalidElementId)
                {
                    return doc.GetElement(room.LevelId) as Level;
                }
            }
        }
        
        // If no rooms, try to get level from other elements
        foreach (ElementId id in memberIds)
        {
            Element member = doc.GetElement(id);
            if (member != null)
            {
                // Try various level parameters
                Parameter levelParam = member.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                if (levelParam == null)
                    levelParam = member.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                if (levelParam == null)
                    levelParam = member.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                if (levelParam == null)
                    levelParam = member.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                
                if (levelParam != null && levelParam.HasValue)
                {
                    ElementId levelId = levelParam.AsElementId();
                    if (levelId != null && levelId != ElementId.InvalidElementId)
                    {
                        return doc.GetElement(levelId) as Level;
                    }
                }
            }
        }
        
        // As a last resort, find the closest level based on group's elevation
        LocationPoint groupLoc = group.Location as LocationPoint;
        if (groupLoc != null)
        {
            double groupZ = groupLoc.Point.Z;
            
            // Get all levels in the document
            FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
            IList<Element> levels = levelCollector
                .OfClass(typeof(Level))
                .ToElements();
            
            Level closestLevel = null;
            double minDistance = double.MaxValue;
            
            foreach (Element elem in levels)
            {
                Level level = elem as Level;
                if (level != null)
                {
                    double distance = Math.Abs(level.Elevation - groupZ);
                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        closestLevel = level;
                    }
                }
            }
            
            return closestLevel;
        }
        
        return null;
    }
    
    private ReferenceElements GetReferenceElements(Group group, Document doc)
    {
        ReferenceElements refElements = new ReferenceElements();
        refElements.GroupOrigin = (group.Location as LocationPoint).Point;
        
        // Get group members
        ICollection<ElementId> memberIds = group.GetMemberIds();
        
        // Collect all elements by type
        List<Wall> walls = new List<Wall>();
        List<Element> curveElements = new List<Element>();
        List<Element> pointElements = new List<Element>();
        
        foreach (ElementId id in memberIds)
        {
            Element elem = doc.GetElement(id);
            if (elem == null) continue;
            
            if (elem is Wall)
            {
                walls.Add(elem as Wall);
            }
            else if (elem.Location is LocationCurve)
            {
                curveElements.Add(elem);
            }
            else if (elem.Location is LocationPoint)
            {
                pointElements.Add(elem);
            }
        }
        
        // Build unique elements list
        Dictionary<string, ElementInfo> uniqueElements = new Dictionary<string, ElementInfo>();
        
        // Process walls first (preferred because they have two points)
        foreach (Wall wall in walls)
        {
            string key = GetEnhancedUniqueKey(wall);
            
            if (!uniqueElements.ContainsKey(key))
            {
                ElementInfo info = new ElementInfo();
                info.Element = wall;
                info.UniqueKey = key;
                
                LocationCurve locCurve = wall.Location as LocationCurve;
                if (locCurve != null)
                {
                    info.Point1 = locCurve.Curve.GetEndPoint(0);
                    info.Point2 = locCurve.Curve.GetEndPoint(1);
                }
                
                uniqueElements[key] = info;
            }
        }
        
        // If we need more elements, add curve elements
        if (uniqueElements.Count < 2)
        {
            foreach (Element elem in curveElements)
            {
                string key = GetEnhancedUniqueKey(elem);
                
                if (!uniqueElements.ContainsKey(key))
                {
                    ElementInfo info = new ElementInfo();
                    info.Element = elem;
                    info.UniqueKey = key;
                    
                    LocationCurve locCurve = elem.Location as LocationCurve;
                    if (locCurve != null)
                    {
                        info.Point1 = locCurve.Curve.GetEndPoint(0);
                        info.Point2 = locCurve.Curve.GetEndPoint(1);
                    }
                    
                    uniqueElements[key] = info;
                }
            }
        }
        
        // If still need more, add point elements
        if (uniqueElements.Count < 2)
        {
            foreach (Element elem in pointElements)
            {
                string key = GetEnhancedUniqueKey(elem);
                
                if (!uniqueElements.ContainsKey(key))
                {
                    ElementInfo info = new ElementInfo();
                    info.Element = elem;
                    info.UniqueKey = key;
                    
                    LocationPoint locPoint = elem.Location as LocationPoint;
                    if (locPoint != null)
                    {
                        info.Point1 = locPoint.Point;
                        info.Point2 = locPoint.Point; // Same point for point elements
                    }
                    
                    uniqueElements[key] = info;
                }
            }
        }
        
        refElements.Elements = uniqueElements.Values.ToList();
        return refElements;
    }
    
    private string GetEnhancedUniqueKey(Element elem)
    {
        // Create unique key based on element type and parameters
        string key = elem.GetType().Name + "|";
        
        // Add category
        if (elem.Category != null)
            key += elem.Category.Id.IntegerValue + "|";
        
        // Add type id
        ElementId typeId = elem.GetTypeId();
        if (typeId != ElementId.InvalidElementId)
            key += typeId.IntegerValue + "|";
        
        // Add geometric properties for walls
        if (elem is Wall)
        {
            Wall wall = elem as Wall;
            
            // Add wall length
            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve != null)
            {
                double length = locCurve.Curve.Length;
                key += $"L:{length:F6}|";
            }
            
            // Add wall height
            Parameter heightParam = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
            if (heightParam != null && heightParam.HasValue)
            {
                key += $"H:{heightParam.AsDouble():F6}|";
            }
        }
        // Add properties for other elements
        else if (elem.Category != null)
        {
            // For curve-based elements, add length
            if (elem.Location is LocationCurve)
            {
                LocationCurve locCurve = elem.Location as LocationCurve;
                double length = locCurve.Curve.Length;
                key += $"L:{length:F6}|";
            }
        }
        
        // Add instance comments if present
        Parameter nameParam = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (nameParam != null && nameParam.HasValue && !string.IsNullOrEmpty(nameParam.AsString()))
            key += nameParam.AsString() + "|";
        
        return key;
    }
    
    private TransformResult CalculateTransformation(ReferenceElements refElements, Group otherGroup, Document doc)
    {
        try
        {
            // Get corresponding elements in other group
            List<ElementInfo> otherElements = GetCorrespondingElements(refElements, otherGroup, doc);
            
            // We need at least 2 matching elements to calculate transformation
            if (otherElements == null || otherElements.Count < 2) 
            {
                return null;
            }
            
            TransformResult result = new TransformResult();
            result.MatchingElements = otherElements.Count;
            
            // Calculate translation (difference in group origins)
            XYZ otherOrigin = (otherGroup.Location as LocationPoint).Point;
            result.Translation = otherOrigin - refElements.GroupOrigin;
            
            // Find two matching elements for transformation calculation
            ElementInfo ref1 = null;
            ElementInfo ref2 = null;
            ElementInfo other1 = null;
            ElementInfo other2 = null;
            
            // Find first matching pair
            foreach (var refElem in refElements.Elements)
            {
                var match = otherElements.FirstOrDefault(e => e.UniqueKey == refElem.UniqueKey);
                if (match != null)
                {
                    ref1 = refElem;
                    other1 = match;
                    break;
                }
            }
            
            // Find second matching pair (different from first)
            foreach (var refElem in refElements.Elements)
            {
                if (refElem == ref1) continue;
                var match = otherElements.FirstOrDefault(e => e.UniqueKey == refElem.UniqueKey);
                if (match != null)
                {
                    ref2 = refElem;
                    other2 = match;
                    break;
                }
            }
            
            if (ref1 == null || ref2 == null || other1 == null || other2 == null)
            {
                // If we can't find two different matching elements, try to work with just one
                if (ref1 != null && other1 != null && ref1.Point1 != null && ref1.Point2 != null)
                {
                    // Can still calculate basic rotation from one element
                    XYZ refVector = (ref1.Point2 - ref1.Point1).Normalize();
                    XYZ otherVector = (other1.Point2 - other1.Point1).Normalize();
                    
                    double singleDotProduct = refVector.DotProduct(otherVector);
                    double singleAngle = Math.Acos(Math.Max(-1.0, Math.Min(1.0, singleDotProduct)));
                    
                    XYZ singleCrossProduct = refVector.CrossProduct(otherVector);
                    if (singleCrossProduct.Z < 0) singleAngle = -singleAngle;
                    
                    result.Rotation = singleAngle * 180 / Math.PI;
                    result.IsMirrored = false; // Can't determine mirroring with one element
                    result.Scale = 1.0;
                    
                    return result;
                }
                return null;
            }
            
            // Calculate vectors in reference group
            XYZ refVector1 = (ref1.Point2 - ref1.Point1).Normalize();
            XYZ refVector2 = (ref2.Point2 - ref2.Point1).Normalize();
            
            // Calculate vectors in other group  
            XYZ otherVector1 = (other1.Point2 - other1.Point1).Normalize();
            XYZ otherVector2 = (other2.Point2 - other2.Point1).Normalize();
            
            // For mirrored groups, the matching elements might have reversed directions
            // Check if vectors are opposite
            double dotProduct = refVector1.DotProduct(otherVector1);
            bool vectorsReversed = dotProduct < -0.9; // Nearly opposite
            
            if (vectorsReversed)
            {
                // For mirrored walls, the direction might be flipped
                // This is common when groups are mirrored
                otherVector1 = -otherVector1;
                otherVector2 = -otherVector2;
                dotProduct = refVector1.DotProduct(otherVector1);
            }
            
            double angle = Math.Acos(Math.Max(-1.0, Math.Min(1.0, dotProduct)));
            
            // Determine rotation direction
            XYZ crossProduct = refVector1.CrossProduct(otherVector1);
            if (crossProduct.Z < 0) angle = -angle;
            
            result.Rotation = angle * 180 / Math.PI;
            
            // Simplified mirror detection for axis-aligned mirrors
            // Check the relative positions
            XYZ refMid1 = (ref1.Point1 + ref1.Point2) * 0.5;
            XYZ refMid2 = (ref2.Point1 + ref2.Point2) * 0.5;
            XYZ otherMid1 = (other1.Point1 + other1.Point2) * 0.5;
            XYZ otherMid2 = (other2.Point1 + other2.Point2) * 0.5;
            
            // Calculate relative positions from group origins
            XYZ refPos1 = refMid1 - refElements.GroupOrigin;
            XYZ refPos2 = refMid2 - refElements.GroupOrigin;
            XYZ otherPos1 = otherMid1 - otherOrigin;
            XYZ otherPos2 = otherMid2 - otherOrigin;
            
            // For simple mirror across Y-Z plane (X-axis mirror)
            // X coordinates should be negated, Y and Z should be same
            bool xMirrored = Math.Abs(refPos1.X + otherPos1.X) < 0.1 && 
                             Math.Abs(refPos1.Y - otherPos1.Y) < 0.1 &&
                             Math.Abs(refPos1.Z - otherPos1.Z) < 0.1;
            
            // For simple mirror across X-Z plane (Y-axis mirror)  
            bool yMirrored = Math.Abs(refPos1.X - otherPos1.X) < 0.1 && 
                             Math.Abs(refPos1.Y + otherPos1.Y) < 0.1 &&
                             Math.Abs(refPos1.Z - otherPos1.Z) < 0.1;
            
            result.IsMirrored = xMirrored || yMirrored || vectorsReversed;
            
            // Scale (typically 1.0 for groups)
            result.Scale = 1.0;
            
            return result;
        }
        catch (Exception ex)
        {
            return null;
        }
    }
    
    private List<ElementInfo> GetCorrespondingElements(ReferenceElements refElements, Group otherGroup, Document doc)
    {
        List<ElementInfo> otherElements = new List<ElementInfo>();
        
        // Get all members of the other group
        ICollection<ElementId> memberIds = otherGroup.GetMemberIds();
        
        // Create a dictionary of elements in the other group by their unique key
        Dictionary<string, Element> otherGroupElements = new Dictionary<string, Element>();
        foreach (ElementId id in memberIds)
        {
            Element elem = doc.GetElement(id);
            if (elem == null) continue;
            
            string key = GetEnhancedUniqueKey(elem);
            if (!otherGroupElements.ContainsKey(key))
            {
                otherGroupElements[key] = elem;
            }
        }
        
        // Find matching elements
        foreach (ElementInfo refInfo in refElements.Elements)
        {
            if (otherGroupElements.ContainsKey(refInfo.UniqueKey))
            {
                Element elem = otherGroupElements[refInfo.UniqueKey];
                
                ElementInfo info = new ElementInfo();
                info.Element = elem;
                info.UniqueKey = refInfo.UniqueKey;
                
                if (elem.Location is LocationCurve)
                {
                    LocationCurve locCurve = elem.Location as LocationCurve;
                    info.Point1 = locCurve.Curve.GetEndPoint(0);
                    info.Point2 = locCurve.Curve.GetEndPoint(1);
                }
                else if (elem.Location is LocationPoint)
                {
                    LocationPoint locPoint = elem.Location as LocationPoint;
                    info.Point1 = locPoint.Point;
                    info.Point2 = locPoint.Point;
                }
                
                otherElements.Add(info);
            }
        }
        
        return otherElements;
    }
    
    private class ReferenceElements
    {
        public XYZ GroupOrigin { get; set; }
        public List<ElementInfo> Elements { get; set; } = new List<ElementInfo>();
    }
    
    private class ElementInfo
    {
        public Element Element { get; set; }
        public string UniqueKey { get; set; }
        public XYZ Point1 { get; set; }
        public XYZ Point2 { get; set; }
    }
    
    private class TransformResult
    {
        public XYZ Translation { get; set; }
        public double Rotation { get; set; }
        public bool IsMirrored { get; set; }
        public double Scale { get; set; }
        public int MatchingElements { get; set; }
    }
}
