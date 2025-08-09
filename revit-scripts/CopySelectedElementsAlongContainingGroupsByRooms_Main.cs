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
public partial class CopySelectedElementsAlongContainingGroupsByRooms : IExternalCommand
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
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        try
        {
            diagnosticLog.Clear();
            diagnosticLog.AppendLine("=== DIAGNOSTIC LOG START ===");
            
            // Get selected elements
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            
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
            
            // Process copying
            var copyResult = ProcessCopying(doc, groupsByType, elementsInGroups, elementToGroupTypeNames);
            
            // Build and show result message
            ShowResults(doc, selectedElements, elementsInGroups, groupsByType, 
                        elementToGroupTypeNames, spatiallyRelevantGroups, allGroups,
                        skipReasons, copyResult);
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
    
    
    // NEW METHOD: Get scope box name for a group
    private string GetGroupScopeBox(Group group, Document doc)
    {
        // Try to get scope box from group members
        ICollection<ElementId> memberIds = group.GetMemberIds();
        foreach (ElementId id in memberIds)
        {
            Element member = doc.GetElement(id);
            if (member != null)
            {
                Parameter scopeBoxParam = member.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                if (scopeBoxParam != null && scopeBoxParam.HasValue)
                {
                    ElementId scopeBoxId = scopeBoxParam.AsElementId();
                    if (scopeBoxId != null && scopeBoxId != ElementId.InvalidElementId)
                    {
                        Element scopeBox = doc.GetElement(scopeBoxId);
                        if (scopeBox != null)
                        {
                            return scopeBox.Name;
                        }
                    }
                }
            }
        }
        return "None";
    }
    
    private CopyResult ProcessCopying(Document doc, 
                                      Dictionary<ElementId, List<Group>> groupsByType,
                                      Dictionary<ElementId, List<Group>> elementsInGroups,
                                      Dictionary<ElementId, List<string>> elementToGroupTypeNames)
    {
        CopyResult result = new CopyResult();
        List<BatchCopyData> allBatchCopyData = new List<BatchCopyData>();
        
        // Track all unique target groups across all group types
        HashSet<ElementId> allPossibleTargetGroups = new HashSet<ElementId>();
        HashSet<ElementId> successfulTargetGroups = new HashSet<ElementId>();
        
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
                        // Build enhanced comment for source elements
                        List<string> commentParts = new List<string>();
                        
                        // Add group type names
                        commentParts.Add(string.Join(", ", elemKvp.Value));
                        
                        // Add source group IDs and scope boxes for this element
                        if (elementsInGroups.ContainsKey(elemKvp.Key))
                        {
                            List<Group> containingGroups = elementsInGroups[elemKvp.Key];
                            if (containingGroups.Count > 0)
                            {
                                // Get first containing group for source info
                                Group sourceGroup = containingGroups.First();
                                commentParts.Add($"source id: {sourceGroup.Id}");
                                
                                string scopeBox = GetGroupScopeBox(sourceGroup, doc);
                                if (!string.IsNullOrEmpty(scopeBox) && scopeBox != "None")
                                {
                                    commentParts.Add($"source scope box: {scopeBox}");
                                }
                            }
                        }
                        
                        string fullComment = string.Join(", ", commentParts);
                        commentsParam.Set(fullComment);
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
                
                // Get all instances of this group type
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Group> allGroupInstances = collector
                    .OfClass(typeof(Group))
                    .WhereElementIsNotElementType()
                    .Where(g => g.GetTypeId() == groupTypeId)
                    .Cast<Group>()
                    .ToList();
                
                diagnosticLog.AppendLine($"Total instances of this group type: {allGroupInstances.Count}");
                diagnosticLog.AppendLine($"Source groups containing selected elements: {containingGroupsOfThisType.Count}");
                
                if (allGroupInstances.Count < 2)
                {
                    diagnosticLog.AppendLine("Skipped: Less than 2 instances");
                    continue;
                }
                
                // NEW: Process EACH containing group as a potential source
                int sourceGroupCount = 0;
                foreach (Group sourceGroup in containingGroupsOfThisType)
                {
                    sourceGroupCount++;
                    diagnosticLog.AppendLine($"\n  Processing source group {sourceGroupCount}/{containingGroupsOfThisType.Count} - ID: {sourceGroup.Id}");
                    
                    // Determine which elements to copy from THIS specific source group
                    List<Element> elementsToCopyFromThisGroup = new List<Element>();
                    foreach (var elemKvp in elementsInGroups)
                    {
                        Element elem = doc.GetElement(elemKvp.Key);
                        if (elemKvp.Value.Contains(sourceGroup))
                        {
                            elementsToCopyFromThisGroup.Add(elem);
                        }
                    }
                    
                    diagnosticLog.AppendLine($"  Elements to copy from this source group: {elementsToCopyFromThisGroup.Count}");
                    
                    if (elementsToCopyFromThisGroup.Count == 0)
                    {
                        diagnosticLog.AppendLine("  Skipped: No elements to copy from this source");
                        continue;
                    }
                    
                    // Process transformations using this source group as reference
                    ProcessGroupTypeTransformations(sourceGroup, allGroupInstances, 
                                                   elementsToCopyFromThisGroup, elementToGroupTypeNames,
                                                   doc, allBatchCopyData, result,
                                                   allPossibleTargetGroups, successfulTargetGroups);
                }
                
                result.GroupTypesProcessed++;
            }
            
            // PHASE 2: Execute batch copy
            ExecuteBatchCopy(doc, allBatchCopyData, result);
            
            // Calculate final statistics from tracked groups
            result.TotalGroupInstancesProcessed = successfulTargetGroups.Count;
            result.TotalGroupInstancesSkipped = allPossibleTargetGroups.Count - successfulTargetGroups.Count;
            
            trans.Commit();
        }
        
        return result;
    }
    
    private void ProcessGroupTypeTransformations(Group referenceGroup, 
                                                 IList<Group> allGroupInstances,
                                                 List<Element> elementsToCopy,
                                                 Dictionary<ElementId, List<string>> elementToGroupTypeNames,
                                                 Document doc,
                                                 List<BatchCopyData> allBatchCopyData,
                                                 CopyResult result,
                                                 HashSet<ElementId> allPossibleTargetGroups,
                                                 HashSet<ElementId> successfulTargetGroups)
    {
        // Get reference elements for transformation calculation
        ReferenceElements refElements = GetReferenceElements(referenceGroup, doc);
        
        if (refElements == null || refElements.Elements.Count < 2)
        {
            result.GroupTypesSkippedNoRefElements++;
            diagnosticLog.AppendLine($"Skipped: Insufficient reference elements (found {refElements?.Elements.Count ?? 0})");
            transformFailureReasons["No Reference Elements"]++;
            return;
        }
        
        diagnosticLog.AppendLine($"Reference elements found: {refElements.Elements.Count}");
        foreach (var refElem in refElements.Elements)
        {
            diagnosticLog.AppendLine($"  - {refElem.UniqueKey}");
        }
        
        // Get scope box for the reference group
        string sourceScopeBox = GetGroupScopeBox(referenceGroup, doc);
        
        // Track successful and failed transformations
        int successfulTransforms = 0;
        int failedTransforms = 0;
        List<string> failedGroupIds = new List<string>();
        
        // Collect batch copy data for all target instances
        foreach (Group otherGroup in allGroupInstances)
        {
            // Skip the source group itself
            if (otherGroup.Id == referenceGroup.Id) continue;
            
            // Track this as a possible target
            allPossibleTargetGroups.Add(otherGroup.Id);
            
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
                
                // Track as successful
                successfulTargetGroups.Add(otherGroup.Id);
                
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
                            : new List<string>(),
                        SourceGroup = referenceGroup,
                        SourceScopeBox = sourceScopeBox
                    });
                }
                successfulTransforms++;
            }
            else
            {
                failedTransforms++;
                failedGroupIds.Add(otherGroup.Id.ToString());
                
                // Log why this specific group failed
                diagnosticLog.AppendLine($"  Failed to transform to group {otherGroup.Id}");
            }
        }
        
        diagnosticLog.AppendLine($"Transformation results: {successfulTransforms} successful, {failedTransforms} failed");
        if (failedGroupIds.Count > 0 && failedGroupIds.Count <= 20)
        {
            diagnosticLog.AppendLine($"Failed group IDs: {string.Join(", ", failedGroupIds)}");
        }
    }
    
    private void ExecuteBatchCopy(Document doc, List<BatchCopyData> allBatchCopyData, CopyResult result)
    {
        diagnosticLog.AppendLine($"\n=== TRANSFORMATION FAILURE SUMMARY ===");
        foreach (var kvp in transformFailureReasons)
        {
            if (kvp.Value > 0)
            {
                diagnosticLog.AppendLine($"{kvp.Key}: {kvp.Value}");
            }
        }
        
        if (allBatchCopyData.Count > 0)
        {
            // Group by target group to ensure unique copies
            var transformGroups = allBatchCopyData
                .GroupBy(bcd => bcd.TargetGroup.Id)
                .ToList();
            
            result.TotalCopyOperations = transformGroups.Count;
            
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
                    
                    if (copiedIds.Count > 0)
                    {
                        UpdateCopiedElements(doc, copiedIds, elementIdsToCopy, batchItems, targetGroup);
                        result.TotalCopied += copiedIds.Count;
                    }
                }
                catch (Exception ex)
                {
                    diagnosticLog.AppendLine($"  ERROR in copy operation: {ex.Message}");
                    diagnosticLog.AppendLine($"  Stack trace: {ex.StackTrace}");
                }
            }
        }
    }
    
    private void UpdateCopiedElements(Document doc, ICollection<ElementId> copiedIds, 
                                     List<ElementId> elementIdsToCopy, 
                                     List<BatchCopyData> batchItems,
                                     Group targetGroup)
    {
        // Map copied elements back to their batch data
        List<ElementId> copiedIdsList = copiedIds.ToList();
        Dictionary<ElementId, ElementId> sourceToCopiedMap = new Dictionary<ElementId, ElementId>();
        
        // Create mapping of source to copied elements
        for (int i = 0; i < copiedIdsList.Count && i < elementIdsToCopy.Count; i++)
        {
            sourceToCopiedMap[elementIdsToCopy[i]] = copiedIdsList[i];
        }
        
        // Update properties for copied elements based on their target groups
        int updatedCount = 0;
        foreach (var batchItem in batchItems)
        {
            if (sourceToCopiedMap.ContainsKey(batchItem.SourceElementId))
            {
                ElementId copiedId = sourceToCopiedMap[batchItem.SourceElementId];
                Element copiedElem = doc.GetElement(copiedId);
                
                if (copiedElem != null)
                {
                    updatedCount++;
                    
                    // Update Comments parameter with enhanced information
                    Parameter commentsParam = copiedElem.get_Parameter(
                        BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                    if (commentsParam != null && !commentsParam.IsReadOnly)
                    {
                        // Build enhanced comment string
                        List<string> commentParts = new List<string>();
                        
                        // Add group type names
                        if (batchItem.GroupTypeNames.Count > 0)
                        {
                            commentParts.Add(string.Join(", ", batchItem.GroupTypeNames));
                        }
                        
                        // Add source group ID
                        if (batchItem.SourceGroup != null)
                        {
                            commentParts.Add($"source id: {batchItem.SourceGroup.Id}");
                        }
                        
                        // Add source scope box
                        if (!string.IsNullOrEmpty(batchItem.SourceScopeBox))
                        {
                            commentParts.Add($"source scope box: {batchItem.SourceScopeBox}");
                        }
                        
                        string fullComment = string.Join(", ", commentParts);
                        commentsParam.Set(fullComment);
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
    }
    
    private void ShowResults(Document doc, List<Element> selectedElements,
                            Dictionary<ElementId, List<Group>> elementsInGroups,
                            Dictionary<ElementId, List<Group>> groupsByType,
                            Dictionary<ElementId, List<string>> elementToGroupTypeNames,
                            List<Group> spatiallyRelevantGroups,
                            IList<Group> allGroups,
                            Dictionary<string, int> skipReasons,
                            CopyResult copyResult)
    {
        StringBuilder resultMessage = new StringBuilder();
        
        if (copyResult.TotalCopied == 0)
        {
            resultMessage.AppendLine($"No elements were copied.");
            resultMessage.AppendLine($"\nDiagnostic Information:");
            resultMessage.AppendLine($"- Selected elements: {selectedElements.Count}");
            resultMessage.AppendLine($"- Elements in groups: {elementsInGroups.Count}");
            resultMessage.AppendLine($"- Group types found: {groupsByType.Count}");
            resultMessage.AppendLine($"- Group types processed: {copyResult.GroupTypesProcessed}");
            resultMessage.AppendLine($"- Group types skipped (no reference elements): {copyResult.GroupTypesSkippedNoRefElements}");
        }
        else
        {
            resultMessage.AppendLine($"Copied {copyResult.TotalCopied} elements to group instances.");
            resultMessage.AppendLine($"\nGroup types processed: {copyResult.GroupTypesProcessed}");
            resultMessage.AppendLine($"Total copy operations: {copyResult.TotalCopyOperations}");
            resultMessage.AppendLine($"Groups checked (spatial filtering): {spatiallyRelevantGroups.Count} of {allGroups.Count}");
            
            // Show group instance processing stats
            resultMessage.AppendLine($"\n=== GROUP INSTANCE PROCESSING ===");
            resultMessage.AppendLine($"Total group instances processed: {copyResult.TotalGroupInstancesProcessed}");
            resultMessage.AppendLine($"Successful transformations: {copyResult.TotalGroupInstancesProcessed - copyResult.TotalGroupInstancesSkipped}");
            resultMessage.AppendLine($"Failed transformations: {copyResult.TotalGroupInstancesSkipped}");
            
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
        
        // Show detailed diagnostics if available
        if (enableDiagnostics && diagnosticLog.Length > 0)
        {
            SaveDiagnostics();
        }
    }
    
    private void SaveDiagnostics()
    {
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
    
    // Helper class for copy results
    private class CopyResult
    {
        public int TotalCopied { get; set; }
        public int GroupTypesProcessed { get; set; }
        public int GroupTypesSkippedNoRefElements { get; set; }
        public int TotalGroupInstancesProcessed { get; set; }
        public int TotalGroupInstancesSkipped { get; set; }
        public int TotalCopyOperations { get; set; }
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
}
