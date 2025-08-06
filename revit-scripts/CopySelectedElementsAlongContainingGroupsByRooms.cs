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
    private bool enableDiagnostics = true; // Enable diagnostic output
    // =================================
    
    // Diagnostic tracking
    private StringBuilder diagnosticLog = new StringBuilder();
    private Dictionary<string, int> transformFailureReasons = new Dictionary<string, int>
    {
        { "No Reference Elements", 0 },
        { "No Corresponding Elements", 0 },
        { "Insufficient Matching Elements", 0 },
        { "Transform Calculation Failed", 0 },
        { "Exception During Transform", 0 }
    };
    
    // Cache for element test points to avoid recalculation
    private Dictionary<ElementId, List<XYZ>> _elementTestPointsCache = new Dictionary<ElementId, List<XYZ>>();
    
    // Cache for room data
    private Dictionary<ElementId, RoomData> _roomDataCache = new Dictionary<ElementId, RoomData>();
    
    // Cache for group bounding boxes
    private Dictionary<ElementId, BoundingBoxXYZ> _groupBoundingBoxCache = new Dictionary<ElementId, BoundingBoxXYZ>();
    
    // Spatial index for groups
    private Dictionary<int, List<Group>> _spatialIndex = new Dictionary<int, List<Group>>();
    private const double SPATIAL_GRID_SIZE = 50.0; // 50 feet grid cells
    
    private class RoomData
    {
        public double LevelZ { get; set; }
        public bool IsUnbound { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
        public Level Level { get; set; }
        public double Height { get; set; }
        public double MinZ { get; set; }
        public double MaxZ { get; set; }
    }
    
    // Structure for true batch copying
    private class BatchCopyData
    {
        public ElementId SourceElementId { get; set; }
        public Transform Transform { get; set; }
        public Group TargetGroup { get; set; }
        public Level TargetLevel { get; set; }
        public List<string> GroupTypeNames { get; set; }
    }
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        try
        {
            diagnosticLog.Clear();
            diagnosticLog.AppendLine("=== DIAGNOSTIC LOG START ===");
            
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
            
            diagnosticLog.AppendLine($"Selected elements: {selectedElements.Count}");
            
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
            
            diagnosticLog.AppendLine($"Total groups in document: {allGroups.Count}");
            
            // Build room cache for ALL rooms in the document
            BuildRoomCache(doc);
            
            // PRE-CALCULATE ALL GROUP BOUNDING BOXES AND BUILD SPATIAL INDEX
            PreCalculateGroupDataAndSpatialIndex(allGroups, doc);
            
            // Get potentially relevant groups using spatial index
            List<Group> spatiallyRelevantGroups = GetSpatiallyRelevantGroups(overallBB);
            
            diagnosticLog.AppendLine($"Spatially relevant groups: {spatiallyRelevantGroups.Count}");
            
            // Dictionary to store which groups contain which elements
            Dictionary<ElementId, List<Group>> elementToGroupsMap = new Dictionary<ElementId, List<Group>>();
            
            // Initialize the map
            foreach (Element elem in selectedElements)
            {
                elementToGroupsMap[elem.Id] = new List<Group>();
            }
            
            // Build comprehensive room-to-group mapping with priority handling
            Dictionary<ElementId, Group> roomToSingleGroupMap = BuildRoomToGroupMapping(allGroups, doc);
            
            // Diagnostic tracking
            Dictionary<string, int> skipReasons = new Dictionary<string, int>
            {
                { "No Bounding Box", 0 },
                { "No BB Intersection", 0 },
                { "No Rooms", 0 },
                { "No Elements in Rooms", 0 },
                { "Contained Elements", 0 }
            };
            
            // Process only spatially relevant groups
            foreach (Group group in spatiallyRelevantGroups)
            {
                // Get cached bounding box
                BoundingBoxXYZ groupBB = _groupBoundingBoxCache[group.Id];
                
                if (groupBB == null)
                {
                    skipReasons["No Bounding Box"]++;
                    continue;
                }
                
                if (!BoundingBoxesIntersect(overallBB, groupBB))
                {
                    skipReasons["No BB Intersection"]++;
                    continue;
                }
                
                // Check which selected elements are contained in this group's rooms
                List<Element> containedElements = GetElementsContainedInGroupRoomsFiltered(
                    group, selectedElements, doc, roomToSingleGroupMap);
                
                if (containedElements.Count == 0)
                {
                    // Check if group has rooms at all
                    bool hasRooms = false;
                    foreach (ElementId id in group.GetMemberIds())
                    {
                        if (doc.GetElement(id) is Room room && room.Area > 0)
                        {
                            hasRooms = true;
                            break;
                        }
                    }
                    
                    if (!hasRooms)
                        skipReasons["No Rooms"]++;
                    else
                        skipReasons["No Elements in Rooms"]++;
                }
                else
                {
                    skipReasons["Contained Elements"]++;
                    
                    // Update the map
                    foreach (Element elem in containedElements)
                    {
                        elementToGroupsMap[elem.Id].Add(group);
                    }
                }
            }
            
            // Sort groups consistently by name when there are multiple groups per element
            foreach (var kvp in elementToGroupsMap)
            {
                kvp.Value.Sort((g1, g2) =>
                {
                    GroupType gt1 = doc.GetElement(g1.GetTypeId()) as GroupType;
                    GroupType gt2 = doc.GetElement(g2.GetTypeId()) as GroupType;
                    string name1 = gt1?.Name ?? "Unknown";
                    string name2 = gt2?.Name ?? "Unknown";
                    return string.Compare(name1, name2, StringComparison.Ordinal);
                });
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
            
            diagnosticLog.AppendLine($"Elements found in groups: {elementsInGroups.Count}");
            
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
            
            diagnosticLog.AppendLine($"Unique group types containing elements: {groupsByType.Count}");
            
            // Process each group type and copy elements
            int totalCopied = 0;
            int groupTypesProcessed = 0;
            int groupTypesSkippedNoRefElements = 0;
            int totalGroupInstancesProcessed = 0;
            int totalGroupInstancesSkipped = 0;
            int totalCopyOperations = 0;  // Track total copy operations
            Dictionary<string, List<string>> diagnosticInfo = new Dictionary<string, List<string>>();
            
            // COLLECT ALL BATCH COPY DATA FIRST
            List<BatchCopyData> allBatchCopyData = new List<BatchCopyData>();
            
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
                            string groupNames = string.Join(", ", elemKvp.Value);
                            commentsParam.Set(groupNames);
                        }
                    }
                }
                
                // PHASE 1: Collect all copy operations across all group types
                foreach (var kvp in groupsByType)
                {
                    ElementId groupTypeId = kvp.Key;
                    List<Group> containingGroupsOfThisType = kvp.Value;
                    
                    GroupType groupType = doc.GetElement(groupTypeId) as GroupType;
                    if (groupType == null) continue;
                    
                    string groupTypeName = groupType.Name ?? "Unknown";
                    diagnosticLog.AppendLine($"\n--- Processing Group Type: {groupTypeName} ---");
                    
                    // Use the first containing group as reference
                    Group referenceGroup = containingGroupsOfThisType.First();
                    diagnosticLog.AppendLine($"Reference group ID: {referenceGroup.Id}");
                    
                    // Get all instances of this group type
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    IList<Group> allGroupInstances = collector
                        .OfClass(typeof(Group))
                        .WhereElementIsNotElementType()
                        .Where(g => g.GetTypeId() == groupTypeId)
                        .Cast<Group>()
                        .ToList();
                    
                    diagnosticLog.AppendLine($"Total instances of this group type: {allGroupInstances.Count}");
                    
                    if (allGroupInstances.Count < 2)
                    {
                        diagnosticLog.AppendLine("Skipped: Less than 2 instances");
                        continue;
                    }
                    
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
                    
                    diagnosticLog.AppendLine($"Elements to copy from reference group: {elementsToCopy.Count}");
                    
                    if (elementsToCopy.Count == 0)
                    {
                        diagnosticLog.AppendLine("Skipped: No elements to copy");
                        continue;
                    }
                    
                    groupTypesProcessed++;
                    
                    // Get reference elements for transformation calculation
                    ReferenceElements refElements = GetReferenceElements(referenceGroup, doc);
                    
                    if (refElements == null || refElements.Elements.Count < 2)
                    {
                        groupTypesSkippedNoRefElements++;
                        diagnosticLog.AppendLine($"Skipped: Insufficient reference elements (found {refElements?.Elements.Count ?? 0})");
                        transformFailureReasons["No Reference Elements"]++;
                        continue;
                    }
                    
                    diagnosticLog.AppendLine($"Reference elements found: {refElements.Elements.Count}");
                    foreach (var refElem in refElements.Elements)
                    {
                        diagnosticLog.AppendLine($"  - {refElem.UniqueKey}");
                    }
                    
                    // Track successful and failed transformations
                    int successfulTransforms = 0;
                    int failedTransforms = 0;
                    List<string> failedGroupIds = new List<string>();
                    
                    // Collect batch copy data for all target instances
                    foreach (Group otherGroup in allGroupInstances)
                    {
                        if (otherGroup.Id == referenceGroup.Id) continue;
                        
                        totalGroupInstancesProcessed++;
                        
                        XYZ otherOrigin = (otherGroup.Location as LocationPoint).Point;
                        
                        // Log which group we're processing
                        diagnosticLog.AppendLine($"  Processing target group {otherGroup.Id} at Z={otherOrigin.Z:F2}");
                        
                        TransformResult transformResult = CalculateTransformation(refElements, otherGroup, doc);
                        
                        if (transformResult != null)
                        {
                            Transform transform = CreateTransform(transformResult, 
                                refElements.GroupOrigin, otherOrigin);
                            Level targetGroupLevel = GetGroupLevel(otherGroup, doc);
                            
                            // Log the transform details
                            diagnosticLog.AppendLine($"    Transform created: Origin Z={transform.Origin.Z:F6}");
                            
                            // Add batch copy data for each element to copy
                            foreach (Element elem in elementsToCopy)
                            {
                                allBatchCopyData.Add(new BatchCopyData
                                {
                                    SourceElementId = elem.Id,
                                    Transform = transform,
                                    TargetGroup = otherGroup,
                                    TargetLevel = targetGroupLevel,
                                    GroupTypeNames = elementToGroupTypeNames.ContainsKey(elem.Id) 
                                        ? elementToGroupTypeNames[elem.Id] 
                                        : new List<string>()
                                });
                            }
                            successfulTransforms++;
                        }
                        else
                        {
                            failedTransforms++;
                            failedGroupIds.Add(otherGroup.Id.ToString());
                            totalGroupInstancesSkipped++;
                            
                            // Log why this specific group failed
                            diagnosticLog.AppendLine($"  Failed to transform to group {otherGroup.Id}");
                        }
                    }
                    
                    diagnosticLog.AppendLine($"Transformation results: {successfulTransforms} successful, {failedTransforms} failed");
                    if (failedGroupIds.Count > 0 && failedGroupIds.Count <= 20)
                    {
                        diagnosticLog.AppendLine($"Failed group IDs: {string.Join(", ", failedGroupIds)}");
                    }
                    
                    diagnosticInfo[groupTypeName] = groupDiagnostics;
                }
                
                diagnosticLog.AppendLine($"\n=== TRANSFORMATION FAILURE SUMMARY ===");
                foreach (var kvp in transformFailureReasons)
                {
                    if (kvp.Value > 0)
                    {
                        diagnosticLog.AppendLine($"{kvp.Key}: {kvp.Value}");
                    }
                }
                
                // PHASE 2: Execute TRUE BATCH COPY - All elements at once!
                if (allBatchCopyData.Count > 0)
                {
                    // IMPORTANT: For vertically stacked groups, we need to ensure each group gets its own transform
                    // Instead of grouping by transform string, group by target group to ensure unique copies
                    var transformGroups = allBatchCopyData
                        .GroupBy(bcd => bcd.TargetGroup.Id)
                        .ToList();
                    
                    totalCopyOperations = transformGroups.Count;  // Store count for use outside transaction
                    
                    diagnosticLog.AppendLine($"\nUnique target groups to receive copies: {transformGroups.Count}");
                    
                    foreach (var transformGroup in transformGroups)
                    {
                        List<BatchCopyData> batchItems = transformGroup.ToList();
                        Group targetGroup = batchItems.First().TargetGroup;
                        Transform sharedTransform = batchItems.First().Transform;
                        
                        // Get unique element IDs for this group
                        List<ElementId> elementIdsToCopy = batchItems
                            .Select(b => b.SourceElementId)
                            .Distinct()
                            .ToList();
                        
                        try
                        {
                            // Log transform details
                            diagnosticLog.AppendLine($"\nCopying to group {targetGroup.Id}:");
                            diagnosticLog.AppendLine($"  Elements to copy: {elementIdsToCopy.Count}");
                            diagnosticLog.AppendLine($"  Transform Origin Z: {sharedTransform.Origin.Z:F6}");
                            
                            // SINGLE COPY CALL for all elements for this specific group
                            ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                doc, 
                                elementIdsToCopy,
                                doc,
                                sharedTransform,
                                null);
                            
                            diagnosticLog.AppendLine($"  Copy result: {copiedIds.Count} elements created (expected {elementIdsToCopy.Count})");
                            
                            // Log the actual copied element IDs and their locations
                            if (copiedIds.Count > 0)
                            {
                                diagnosticLog.AppendLine($"  Copied element IDs: {string.Join(", ", copiedIds.Take(3))}...");
                                
                                // Check positions of copied elements
                                Element firstCopied = doc.GetElement(copiedIds.First());
                                Element firstSource = doc.GetElement(elementIdsToCopy.First());
                                
                                if (firstCopied != null && firstSource != null)
                                {
                                    // Get locations
                                    XYZ copiedLoc = null;
                                    XYZ sourceLoc = null;
                                    
                                    if (firstCopied.Location is LocationPoint copiedLocPt)
                                        copiedLoc = copiedLocPt.Point;
                                    else if (firstCopied.Location is LocationCurve copiedLocCrv)
                                        copiedLoc = copiedLocCrv.Curve.GetEndPoint(0);
                                    
                                    if (firstSource.Location is LocationPoint sourceLocPt)
                                        sourceLoc = sourceLocPt.Point;
                                    else if (firstSource.Location is LocationCurve sourceLocCrv)
                                        sourceLoc = sourceLocCrv.Curve.GetEndPoint(0);
                                    
                                    if (copiedLoc != null && sourceLoc != null)
                                    {
                                        diagnosticLog.AppendLine($"  Source element at: X={sourceLoc.X:F2}, Y={sourceLoc.Y:F2}, Z={sourceLoc.Z:F2}");
                                        diagnosticLog.AppendLine($"  Copied element at: X={copiedLoc.X:F2}, Y={copiedLoc.Y:F2}, Z={copiedLoc.Z:F2}");
                                        
                                        XYZ actualDisplacement = copiedLoc - sourceLoc;
                                        diagnosticLog.AppendLine($"  Actual displacement: X={actualDisplacement.X:F2}, Y={actualDisplacement.Y:F2}, Z={actualDisplacement.Z:F2}");
                                        
                                        // Compare with expected displacement
                                        XYZ expectedDisplacement = targetGroup.Location is LocationPoint grpLoc 
                                            ? grpLoc.Point - (batchItems.First().TargetGroup.Location as LocationPoint).Point
                                            : XYZ.Zero;
                                        
                                        // Check if position is correct relative to target group
                                        XYZ targetGroupOrigin = (targetGroup.Location as LocationPoint)?.Point ?? XYZ.Zero;
                                        XYZ relativePos = copiedLoc - targetGroupOrigin;
                                        diagnosticLog.AppendLine($"  Position relative to target group: X={relativePos.X:F2}, Y={relativePos.Y:F2}, Z={relativePos.Z:F2}");
                                    }
                                    
                                    // Check if element is visible
                                    if (firstCopied.IsHidden(doc.ActiveView))
                                    {
                                        diagnosticLog.AppendLine($"  WARNING: Copied element is hidden in active view!");
                                    }
                                    
                                    // Check element's level
                                    Parameter levelParam = firstCopied.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                                    if (levelParam != null && levelParam.HasValue)
                                    {
                                        ElementId levelId = levelParam.AsElementId();
                                        Level level = doc.GetElement(levelId) as Level;
                                        diagnosticLog.AppendLine($"  Copied element level: {level?.Name ?? "Unknown"}");
                                    }
                                }
                            }
                            
                            if (copiedIds.Count != elementIdsToCopy.Count)
                            {
                                diagnosticLog.AppendLine($"  WARNING: Mismatch in copy count!");
                            }
                            
                            if (copiedIds.Count > 0)
                            {
                                // Map copied elements back to their batch data
                                List<ElementId> copiedIdsList = copiedIds.ToList();
                                Dictionary<ElementId, ElementId> sourceToCopieMa
                                    = new Dictionary<ElementId, ElementId>();
                                
                                // Create mapping of source to copied elements
                                for (int i = 0; i < copiedIdsList.Count && i < elementIdsToCopy.Count; i++)
                                {
                                    sourceToCopieMa[elementIdsToCopy[i]] = copiedIdsList[i];
                                }
                                
                                // Update properties for copied elements based on their target groups
                                int updatedCount = 0;
                                foreach (var batchItem in batchItems)
                                {
                                    if (sourceToCopieMa.ContainsKey(batchItem.SourceElementId))
                                    {
                                        ElementId copiedId = sourceToCopieMa[batchItem.SourceElementId];
                                        Element copiedElem = doc.GetElement(copiedId);
                                        
                                        if (copiedElem != null)
                                        {
                                            updatedCount++;
                                            
                                            // Update Comments parameter
                                            Parameter commentsParam = copiedElem.get_Parameter(
                                                BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                            if (commentsParam != null && !commentsParam.IsReadOnly && 
                                                batchItem.GroupTypeNames.Count > 0)
                                            {
                                                string groupNames = string.Join(", ", batchItem.GroupTypeNames);
                                                commentsParam.Set(groupNames);
                                            }
                                            
                                            // Update level if needed
                                            if (batchItem.TargetLevel != null)
                                            {
                                                UpdateElementLevel(copiedElem, batchItem.TargetLevel);
                                            }
                                        }
                                    }
                                }
                                
                                diagnosticLog.AppendLine($"  Updated properties for {updatedCount} copied elements");
                                totalCopied += copiedIds.Count;
                            }
                            else
                            {
                                diagnosticLog.AppendLine($"  WARNING: No elements were copied!");
                            }
                        }
                        catch (Exception ex)
                        {
                            diagnosticLog.AppendLine($"  ERROR in copy operation: {ex.Message}");
                            diagnosticLog.AppendLine($"  Stack trace: {ex.StackTrace}");
                        }
                    }
                }
                
                // Add final verification check
                diagnosticLog.AppendLine($"\n=== FINAL VERIFICATION ===");
                
                // Count how many groups actually have the copied elements
                int groupsWithCopiedElements = 0;
                List<string> groupsWithoutElements = new List<string>();
                
                // Store reference group IDs to skip them
                HashSet<ElementId> referenceGroupIds = new HashSet<ElementId>();
                foreach (var kvp in elementsInGroups)
                {
                    foreach (Group g in kvp.Value)
                    {
                        referenceGroupIds.Add(g.Id);
                    }
                }
                
                foreach (var kvp in groupsByType)
                {
                    ElementId groupTypeId = kvp.Key;
                    
                    // Get all instances of this group type
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    IList<Group> allGroupInstances = collector
                        .OfClass(typeof(Group))
                        .WhereElementIsNotElementType()
                        .Where(g => g.GetTypeId() == groupTypeId)
                        .Cast<Group>()
                        .ToList();
                    
                    foreach (Group group in allGroupInstances)
                    {
                        // Skip any reference groups
                        if (referenceGroupIds.Contains(group.Id)) continue;
                        
                        // Check if this group's location has any of the copied elements nearby
                        XYZ groupOrigin = (group.Location as LocationPoint)?.Point;
                        if (groupOrigin != null)
                        {
                            // Create a small bounding box around the group origin
                            BoundingBoxXYZ searchBox = new BoundingBoxXYZ();
                            searchBox.Min = new XYZ(groupOrigin.X - 50, groupOrigin.Y - 50, groupOrigin.Z - 5);
                            searchBox.Max = new XYZ(groupOrigin.X + 50, groupOrigin.Y + 50, groupOrigin.Z + 15);
                            
                            // Use a collector to find sprinklers in this area
                            BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(new Outline(searchBox.Min, searchBox.Max));
                            FilteredElementCollector sprinklerCollector = new FilteredElementCollector(doc);
                            
                            // Assuming sprinklers are in the Fire Protection category
                            var nearbyElements = sprinklerCollector
                                .OfCategory(BuiltInCategory.OST_Sprinklers)
                                .WherePasses(bbFilter)
                                .ToElements();
                            
                            if (nearbyElements.Count > 0)
                            {
                                groupsWithCopiedElements++;
                                diagnosticLog.AppendLine($"Group {group.Id} at Z={groupOrigin.Z:F2}: Found {nearbyElements.Count} sprinklers");
                            }
                            else
                            {
                                groupsWithoutElements.Add($"{group.Id} (Z={groupOrigin.Z:F2})");
                                diagnosticLog.AppendLine($"Group {group.Id} at Z={groupOrigin.Z:F2}: NO SPRINKLERS FOUND!");
                            }
                        }
                    }
                }
                
                diagnosticLog.AppendLine($"\nGroups with copied elements: {groupsWithCopiedElements}");
                if (groupsWithoutElements.Count > 0)
                {
                    diagnosticLog.AppendLine($"Groups WITHOUT elements: {string.Join(", ", groupsWithoutElements.Take(10))}");
                    if (groupsWithoutElements.Count > 10)
                    {
                        diagnosticLog.AppendLine($"  ... and {groupsWithoutElements.Count - 10} more");
                    }
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
            }
            else
            {
                resultMessage.AppendLine($"Copied {totalCopied} elements to group instances.");
                resultMessage.AppendLine($"\nGroup types processed: {groupTypesProcessed}");
                resultMessage.AppendLine($"Total copy operations: {totalCopyOperations}");
                resultMessage.AppendLine($"Groups checked (spatial filtering): {spatiallyRelevantGroups.Count} of {allGroups.Count}");
                
                // NEW: Show group instance processing stats
                resultMessage.AppendLine($"\n=== GROUP INSTANCE PROCESSING ===");
                resultMessage.AppendLine($"Total group instances processed: {totalGroupInstancesProcessed}");
                resultMessage.AppendLine($"Successful transformations: {totalGroupInstancesProcessed - totalGroupInstancesSkipped}");
                resultMessage.AppendLine($"Failed transformations: {totalGroupInstancesSkipped}");
                
                if (transformFailureReasons.Any(kvp => kvp.Value > 0))
                {
                    resultMessage.AppendLine($"\nTransformation Failure Reasons:");
                    foreach (var kvp in transformFailureReasons)
                    {
                        if (kvp.Value > 0)
                        {
                            resultMessage.AppendLine($"- {kvp.Key}: {kvp.Value}");
                        }
                    }
                }
                
                // Show diagnostic info about group processing
                if (enableDiagnostics)
                {
                    resultMessage.AppendLine($"\n=== GROUP FILTERING DIAGNOSTICS ===");
                    resultMessage.AppendLine($"Groups analyzed: {spatiallyRelevantGroups.Count}");
                    foreach (var kvp in skipReasons)
                    {
                        if (kvp.Key == "Contained Elements")
                        {
                            resultMessage.AppendLine($"- Groups with selected elements: {kvp.Value}");
                        }
                        else
                        {
                            resultMessage.AppendLine($"- Groups skipped ({kvp.Key}): {kvp.Value}");
                        }
                    }
                }
                
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
                
                // Add option to save detailed log
                if (enableDiagnostics && diagnosticLog.Length > 0)
                {
                    resultMessage.AppendLine("\n[Detailed diagnostic log available - check journal file or copy from next dialog]");
                }
            }
            
            TaskDialog.Show("Copy Complete", resultMessage.ToString());
            
            // ALWAYS show detailed diagnostics if we have any log content
            if (enableDiagnostics && diagnosticLog.Length > 0)
            {
                // Write to file as well
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"RevitCopyDiagnostics_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                string filePath = System.IO.Path.Combine(desktopPath, fileName);
                
                try
                {
                    System.IO.File.WriteAllText(filePath, diagnosticLog.ToString());
                    
                    TaskDialog detailedDialog = new TaskDialog("Detailed Diagnostics");
                    detailedDialog.MainInstruction = "Transformation Diagnostic Details";
                    detailedDialog.MainContent = $"Diagnostics saved to:\n{filePath}\n\n" + 
                                                 "First 5000 characters shown below:\n\n" +
                                                 (diagnosticLog.Length > 5000 
                                                    ? diagnosticLog.ToString().Substring(0, 5000) + "\n\n[... truncated, see file for full log]"
                                                    : diagnosticLog.ToString());
                    detailedDialog.ExpandedContent = "Full diagnostics have been saved to your desktop.";
                    detailedDialog.Show();
                }
                catch (Exception ex)
                {
                    // If file write fails, still show dialog
                    TaskDialog detailedDialog = new TaskDialog("Detailed Diagnostics");
                    detailedDialog.MainInstruction = "Transformation Diagnostic Details";
                    detailedDialog.MainContent = diagnosticLog.ToString();
                    detailedDialog.ExpandedContent = $"Could not save to file: {ex.Message}\n\nCopy this text manually if needed.";
                    detailedDialog.Show();
                }
            }
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
    
    // Build room cache for all rooms
    private void BuildRoomCache(Document doc)
    {
        FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
        List<Room> allRooms = roomCollector
            .OfClass(typeof(SpatialElement))
            .WhereElementIsNotElementType()
            .OfType<Room>()
            .Where(r => r != null && r.Area > 0)
            .ToList();
        
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
                Level = roomLevel,
                LevelZ = roomLevelElevation,
                Height = roomHeight,
                MinZ = roomLevelElevation,
                MaxZ = roomLevelElevation + roomHeight,
                BoundingBox = room.get_BoundingBox(null),
                IsUnbound = room.Volume <= 0.001
            };
            
            _roomDataCache[room.Id] = data;
        }
    }
    
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
    
    // Pre-calculate all group bounding boxes and build spatial index
    private void PreCalculateGroupDataAndSpatialIndex(IList<Group> allGroups, Document doc)
    {
        _spatialIndex.Clear();
        
        foreach (Group group in allGroups)
        {
            // Calculate and cache bounding box
            BoundingBoxXYZ bb = CalculateGroupBoundingBox(group, doc);
            if (bb != null)
            {
                _groupBoundingBoxCache[group.Id] = bb;
                
                // Add to spatial index
                AddToSpatialIndex(group, bb);
            }
        }
    }
    
    // Calculate group bounding box (without checking cache)
    private BoundingBoxXYZ CalculateGroupBoundingBox(Group group, Document doc)
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
    
    // Add group to spatial index
    private void AddToSpatialIndex(Group group, BoundingBoxXYZ bb)
    {
        // Get all grid cells that this bounding box overlaps
        int minGridX = (int)Math.Floor(bb.Min.X / SPATIAL_GRID_SIZE);
        int maxGridX = (int)Math.Floor(bb.Max.X / SPATIAL_GRID_SIZE);
        int minGridY = (int)Math.Floor(bb.Min.Y / SPATIAL_GRID_SIZE);
        int maxGridY = (int)Math.Floor(bb.Max.Y / SPATIAL_GRID_SIZE);
        
        // Add group to all overlapping cells
        for (int x = minGridX; x <= maxGridX; x++)
        {
            for (int y = minGridY; y <= maxGridY; y++)
            {
                int cellKey = GetSpatialCellKey(x, y);
                if (!_spatialIndex.ContainsKey(cellKey))
                {
                    _spatialIndex[cellKey] = new List<Group>();
                }
                _spatialIndex[cellKey].Add(group);
            }
        }
    }
    
    // Get spatial cell key from grid coordinates
    private int GetSpatialCellKey(int gridX, int gridY)
    {
        // Simple hash combining x and y coordinates
        // Assumes grid coordinates are within reasonable bounds
        return (gridX + 10000) * 20000 + (gridY + 10000);
    }
    
    // Get groups that might intersect with the given bounding box
    private List<Group> GetSpatiallyRelevantGroups(BoundingBoxXYZ targetBB)
    {
        HashSet<Group> relevantGroups = new HashSet<Group>();
        
        // Get all grid cells that the target bounding box overlaps
        int minGridX = (int)Math.Floor(targetBB.Min.X / SPATIAL_GRID_SIZE);
        int maxGridX = (int)Math.Floor(targetBB.Max.X / SPATIAL_GRID_SIZE);
        int minGridY = (int)Math.Floor(targetBB.Min.Y / SPATIAL_GRID_SIZE);
        int maxGridY = (int)Math.Floor(targetBB.Max.Y / SPATIAL_GRID_SIZE);
        
        // Collect all groups from overlapping cells
        for (int x = minGridX; x <= maxGridX; x++)
        {
            for (int y = minGridY; y <= maxGridY; y++)
            {
                int cellKey = GetSpatialCellKey(x, y);
                if (_spatialIndex.ContainsKey(cellKey))
                {
                    foreach (Group group in _spatialIndex[cellKey])
                    {
                        relevantGroups.Add(group);
                    }
                }
            }
        }
        
        return relevantGroups.ToList();
    }
    
    // Convert transform to string for grouping
    private string TransformToString(Transform t)
    {
        return $"{t.BasisX.X:F6},{t.BasisX.Y:F6},{t.BasisX.Z:F6}," +
               $"{t.BasisY.X:F6},{t.BasisY.Y:F6},{t.BasisY.Z:F6}," +
               $"{t.BasisZ.X:F6},{t.BasisZ.Y:F6},{t.BasisZ.Z:F6}," +
               $"{t.Origin.X:F6},{t.Origin.Y:F6},{t.Origin.Z:F6}";
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
            // For curves, check endpoints and midpoint
            Curve curve = locCurve.Curve;
            testPoints.Add(curve.GetEndPoint(0));
            testPoints.Add(curve.GetEndPoint(1));
            testPoints.Add(curve.Evaluate(0.5, true));
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
    
    // Check if two bounding boxes intersect
    private bool BoundingBoxesIntersect(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        return !(bb1.Max.X < bb2.Min.X || bb2.Max.X < bb1.Min.X ||
                 bb1.Max.Y < bb2.Min.Y || bb2.Max.Y < bb1.Min.Y ||
                 bb1.Max.Z < bb2.Min.Z || bb2.Max.Z < bb1.Min.Z);
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
        // Calculate the translation needed
        XYZ translation = targetOrigin - refOrigin;
        
        // Log the translation for debugging
        diagnosticLog.AppendLine($"      Translation: X={translation.X:F3}, Y={translation.Y:F3}, Z={translation.Z:F3}");
        diagnosticLog.AppendLine($"      Rotation: {transformResult.Rotation:F2} degrees, Mirrored: {transformResult.IsMirrored}");
        
        // BACK TO SIMPLE TRANSLATION - the mirroring is already handled by the group itself!
        // Groups maintain their own internal coordinate system
        Transform transform = Transform.CreateTranslation(translation);
        
        diagnosticLog.AppendLine($"      Using simple translation (ignoring mirror/rotation - handled by group)");
        diagnosticLog.AppendLine($"      Transform Origin: X={transform.Origin.X:F3}, Y={transform.Origin.Y:F3}, Z={transform.Origin.Z:F3}");
        
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
                diagnosticLog.AppendLine($"    Group {otherGroup.Id}: Only {otherElements?.Count ?? 0} corresponding elements found");
                
                if (otherElements == null || otherElements.Count == 0)
                    transformFailureReasons["No Corresponding Elements"]++;
                else
                    transformFailureReasons["Insufficient Matching Elements"]++;
                
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
                
                diagnosticLog.AppendLine($"    Group {otherGroup.Id}: Could not find sufficient matching element pairs");
                transformFailureReasons["Insufficient Matching Elements"]++;
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
            diagnosticLog.AppendLine($"    Group {otherGroup.Id}: Exception during transform calculation: {ex.Message}");
            transformFailureReasons["Exception During Transform"]++;
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
