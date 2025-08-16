using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using System.Threading;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
    // Structure for true batch copying
    private class BatchCopyData
    {
        public ElementId SourceElementId { get; set; }
        public Transform Transform { get; set; }
        public Group TargetGroup { get; set; }
        public Level TargetLevel { get; set; }
        public List<string> GroupTypeNames { get; set; }
        public Group SourceGroup { get; set; } // Track which group instance this element came from
    }

    private CopyResult ProcessCopying(Document doc,
                                      Dictionary<ElementId, List<Group>> groupsByType,
                                      Dictionary<ElementId, List<Group>> elementsInGroups,
                                      Dictionary<ElementId, List<string>> elementToGroupTypeNames,
                                      int totalSelectedElements = 0,
                                      int totalElementsInGroups = 0)
    {
        CopyResult result = new CopyResult();
        List<BatchCopyData> allBatchCopyData = new List<BatchCopyData>();

        // Track all unique target groups across all group types
        HashSet<ElementId> allPossibleTargetGroups = new HashSet<ElementId>();
        HashSet<ElementId> successfulTargetGroups = new HashSet<ElementId>();
        
        int totalGroupTypes = groupsByType.Count;
        int currentGroupType = 0;

        using (Transaction trans = new Transaction(doc, "Copy Elements Following Containing Groups By Rooms"))
        {
            trans.Start();

            // First, update Comments parameter for all selected elements that are in groups
            UpdateSourceElementComments(doc, elementToGroupTypeNames, elementsInGroups);

            if (progressForm != null)
            {
                progressForm.SetPhase("Collecting Copy Operations");
                progressForm.UpdateProgress(0, totalGroupTypes, "Analyzing group types...");
            }

            // PHASE 1: Collect all copy operations across all group types
            foreach (var kvp in groupsByType)
            {
                currentGroupType++;
                ElementId groupTypeId = kvp.Key;
                List<Group> containingGroupsOfThisType = kvp.Value;

                GroupType groupType = doc.GetElement(groupTypeId) as GroupType;
                
                if (CheckCancellation())
                    break;
                
                if (groupType == null) continue;

                // Get all instances of this group type
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Group> allGroupInstances = collector
                    .OfClass(typeof(Group))
                    .WhereElementIsNotElementType()
                    .Where(g => g.GetTypeId() == groupTypeId)
                    .Cast<Group>()
                    .ToList();

                if (allGroupInstances.Count < 2)
                {
                    continue;
                }
                
                if (progressForm != null)
                {
                    progressForm.UpdateProgress(currentGroupType, totalGroupTypes, 
                        $"Processing group type {currentGroupType} of {totalGroupTypes}",
                        $"{groupType.Name} ({allGroupInstances.Count} instances)");
                    progressForm.AddGroupTypeProcessed(groupType.Name, allGroupInstances.Count);
                    Application.DoEvents();
                }

                // Process EACH containing group as a potential source
                foreach (Group sourceGroup in containingGroupsOfThisType)
                {
                    if (CheckCancellation())
                        break;
                        
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

                    if (elementsToCopyFromThisGroup.Count == 0)
                    {
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

            if (CheckCancellation())
            {
                trans.RollBack();
                return result;
            }

            // PHASE 2: Execute batch copy
            if (progressForm != null)
            {
                progressForm.SetPhase("Executing Copy Operations");
                progressForm.UpdateProgress(0, allBatchCopyData.Count, "Copying elements...");
            }
            
            ExecuteBatchCopy(doc, allBatchCopyData, result, totalSelectedElements, totalElementsInGroups);

            // Calculate final statistics from tracked groups
            result.TotalGroupInstancesProcessed = successfulTargetGroups.Count;
            result.TotalGroupInstancesSkipped = allPossibleTargetGroups.Count - successfulTargetGroups.Count;

            trans.Commit();
        }

        return result;
    }

    private void UpdateSourceElementComments(Document doc,
                                            Dictionary<ElementId, List<string>> elementToGroupTypeNames,
                                            Dictionary<ElementId, List<Group>> elementsInGroups)
    {
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

                    // Add source group IDs for this element
                    if (elementsInGroups.ContainsKey(elemKvp.Key))
                    {
                        List<Group> containingGroups = elementsInGroups[elemKvp.Key];
                        if (containingGroups.Count > 0)
                        {
                            // Get first containing group for source info
                            Group sourceGroup = containingGroups.First();
                            commentParts.Add($"source id: {sourceGroup.Id}");

                        }
                    }

                    string fullComment = string.Join(", ", commentParts);
                    commentsParam.Set(fullComment);
                }
            }
        }
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
            return;
        }

        // Collect batch copy data for all target instances
        foreach (Group otherGroup in allGroupInstances)
        {
            // Skip the source group itself
            if (otherGroup.Id == referenceGroup.Id) continue;

            if (CheckCancellation())
                break;

            // OPTIMIZATION: Skip only if groups are at VERY different elevations
            // Changed from 50.0 to 200.0 to allow copying between more floors
            // This was causing groups on different floors to be skipped
            LocationPoint sourceLoc = referenceGroup.Location as LocationPoint;
            LocationPoint otherLoc = otherGroup.Location as LocationPoint;
            if (sourceLoc != null && otherLoc != null)
            {
                double zDiff = Math.Abs(sourceLoc.Point.Z - otherLoc.Point.Z);
                if (zDiff > 200.0) // Only skip if groups are on VERY different floors
                {
                    LogGroupSkipped("Elevation difference > 200ft", otherGroup, zDiff);
                    continue;
                }
            }

            // Track this as a possible target
            allPossibleTargetGroups.Add(otherGroup.Id);

            XYZ otherOrigin = (otherGroup.Location as LocationPoint).Point;

            TransformResult transformResult = CalculateTransformation(refElements, otherGroup, doc);

            if (transformResult != null)
            {
                Transform transform = CreateTransform(transformResult,
                    refElements.GroupOrigin, otherOrigin);
                Level targetGroupLevel = GetGroupLevel(otherGroup, doc);

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
                    });
                }
            }
            else
            {
                LogGroupSkipped("No valid transformation found", otherGroup);
            }
        }
    }
    

    private void ExecuteBatchCopy(Document doc, List<BatchCopyData> allBatchCopyData, CopyResult result,
                                  int totalSelectedElements = 0,
                                  int totalElementsInGroups = 0)
    {
        if (allBatchCopyData.Count > 0)
        {
            // Group by target group to ensure unique copies
            var transformGroups = allBatchCopyData
                .GroupBy(bcd => bcd.TargetGroup.Id)
                .ToList();

            result.TotalCopyOperations = transformGroups.Count;
            int currentOperation = 0;

            foreach (var transformGroup in transformGroups)
            {
                List<BatchCopyData> batchItems = transformGroup.ToList();
                Group targetGroup = batchItems.First().TargetGroup;
                Transform sharedTransform = batchItems.First().Transform;

                currentOperation++;
                // Stop the timer before the last copy operation
                if (currentOperation == transformGroups.Count && progressForm != null)
                {
                    progressForm.StopTimer();
                }

                
                if (progressForm != null)
                {
                    GroupType gt = doc.GetElement(targetGroup.GetTypeId()) as GroupType;
                    string targetGroupName = gt?.Name ?? "Unknown";
                    
                    
                    // Get level for target group
                    Level targetLevel = GetGroupLevel(targetGroup, doc);
                    string levelName = targetLevel?.Name ?? "Unknown Level";
                    
                    string targetInfo = $"{targetGroupName} ({levelName})";
                    progressForm.UpdateProgress(currentOperation, transformGroups.Count,
                        $"Copying to group {currentOperation} of {transformGroups.Count}",
                        $"Target: {targetInfo}");
                    Application.DoEvents();
                }
                
                if (CheckCancellation())
                    break;

                // Get unique element IDs for this group
                List<ElementId> elementIdsToCopy = batchItems
                    .Select(b => b.SourceElementId)
                    .Distinct()
                    .ToList();

                try
                {
                    // SINGLE COPY CALL for all elements for this specific group
                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                        doc,
                        elementIdsToCopy,
                        doc,
                        sharedTransform,
                        null);

                    if (copiedIds.Count > 0)
                    {
                        UpdateCopiedElements(doc, copiedIds, elementIdsToCopy, batchItems, targetGroup);
                        result.TotalCopied += copiedIds.Count;
                        
                        if (progressForm != null)
                        {
                            // Don't reset the selected/in groups counts, just update copied count
                            // The form will preserve the original values
                            progressForm.UpdateElementCounts(totalSelectedElements, totalElementsInGroups, result.TotalCopied);
                            GroupType targetGt = doc.GetElement(targetGroup.GetTypeId()) as GroupType;
                            Level targetLevel = GetGroupLevel(targetGroup, doc);
                            string levelName = targetLevel?.Name ?? "Unknown Level";
                            string targetInfo = $"{targetGt?.Name ?? "Unknown"} ({levelName})";
                            
                            progressForm.AddCopyOperation(targetGt?.Name ?? "Unknown", targetInfo);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Silent fail - log if needed
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
        foreach (var batchItem in batchItems)
        {
            if (sourceToCopiedMap.ContainsKey(batchItem.SourceElementId))
            {
                ElementId copiedId = sourceToCopiedMap[batchItem.SourceElementId];
                Element copiedElem = doc.GetElement(copiedId);

                if (copiedElem != null)
                {
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
