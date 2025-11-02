using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class CombineSelectedElementsContainingSelectedModelGroups : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get current selection
        var selectedIds = uidoc.GetSelectionIds();

        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "No elements selected. Please select elements and groups first.");
            return Result.Failed;
        }

        // Separate groups from other elements in the selection
        var groups = new List<Group>();
        var elementsToCheck = new List<Element>();

        foreach (var id in selectedIds)
        {
            var element = doc.GetElement(id);
            if (element is Group group && group.GroupType != null)
            {
                // Only add model groups (not detail groups)
                if (group.Category != null && 
                    group.Category.Id.IntegerValue == (int)BuiltInCategory.OST_IOSModelGroups)
                {
                    groups.Add(group);
                }
            }
            else
            {
                // Check if element can be added to a group
                if (CanBeAddedToGroup(element))
                {
                    elementsToCheck.Add(element);
                }
            }
        }

        // Check if we have groups and elements
        if (groups.Count == 0)
        {
            TaskDialog.Show("Error", "No model groups found in selection. Please select at least one model group.");
            return Result.Failed;
        }

        if (elementsToCheck.Count == 0)
        {
            TaskDialog.Show("Error", "No valid elements to add to groups found in selection.");
            return Result.Failed;
        }

        // Create a mapping of groups to elements that intersect/are contained
        var groupToElementsMap = new Dictionary<Group, List<Element>>();
        
        foreach (var group in groups)
        {
            var groupBB = GetGroupBoundingBox(group);
            if (groupBB == null) continue;

            var intersectingElements = new List<Element>();
            
            foreach (var element in elementsToCheck)
            {
                var elementBB = element.get_BoundingBox(null);
                if (elementBB != null)
                {
                    // Check if element bounding box intersects or is contained within group bounding box
                    if (BoundingBoxIntersectsOrContains(groupBB, elementBB))
                    {
                        intersectingElements.Add(element);
                    }
                }
            }
            
            if (intersectingElements.Count > 0)
            {
                groupToElementsMap[group] = intersectingElements;
            }
        }

        if (groupToElementsMap.Count == 0)
        {
            TaskDialog.Show("Info", "No elements found intersecting or contained within the selected groups.");
            return Result.Cancelled;
        }

        // Show confirmation dialog with details
        var summaryText = "Element to Group Mapping:\n\n";
        foreach (var kvp in groupToElementsMap)
        {
            summaryText += $"Group '{kvp.Key.GroupType.Name}' (Id: {kvp.Key.Id}):\n";
            summaryText += $"  - {kvp.Value.Count} element(s) will be added\n";
        }
        
        // Check for elements that will be added to multiple groups
        var elementToGroupsMap = new Dictionary<ElementId, List<Group>>();
        foreach (var kvp in groupToElementsMap)
        {
            foreach (var elem in kvp.Value)
            {
                if (!elementToGroupsMap.ContainsKey(elem.Id))
                    elementToGroupsMap[elem.Id] = new List<Group>();
                elementToGroupsMap[elem.Id].Add(kvp.Key);
            }
        }
        
        var multiGroupElements = elementToGroupsMap.Where(x => x.Value.Count > 1).ToList();
        if (multiGroupElements.Any())
        {
            summaryText += $"\nNote: {multiGroupElements.Count} element(s) will be copied to multiple groups.\n";
        }

        var td = new TaskDialog("Confirm Operation")
        {
            MainContent = summaryText + 
                         "\nGroups will be ungrouped and recreated with the additional elements.\n" +
                         "All instance parameters will be preserved.\n\n" +
                         "Continue?",
            MainInstruction = "Add Intersecting Elements to Groups",
            CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
        };

        if (td.Show() != TaskDialogResult.Yes)
        {
            return Result.Cancelled;
        }

        // Process groups
        using (Transaction trans = new Transaction(doc, "Add Intersecting Elements to Model Groups"))
        {
            trans.Start();

            var processedGroups = new List<ElementId>();
            var errors = new List<string>();

            try
            {
                foreach (var kvp in groupToElementsMap)
                {
                    try
                    {
                        var targetGroup = kvp.Key;
                        var elementsToAdd = kvp.Value;
                        
                        // Store original group info
                        var originalGroupType = targetGroup.GroupType;
                        string originalGroupName = originalGroupType.Name;

                        // Store parameters
                        var parameterValues = StoreGroupParameters(targetGroup);
                        
                        // Store location
                        XYZ groupLocation = null;
                        double groupRotation = 0;
                        bool hasRotation = false;
                        if (targetGroup.Location is LocationPoint locPoint)
                        {
                            groupLocation = locPoint.Point;
                            try
                            {
                                groupRotation = locPoint.Rotation;
                                hasRotation = true;
                            }
                            catch { }
                        }

                        // Create copies of elements to add
                        var copiedElementIds = new List<ElementId>();
                        foreach (var elem in elementsToAdd)
                        {
                            try
                            {
                                var copiedIds = ElementTransformUtils.CopyElements(
                                    doc, 
                                    new List<ElementId> { elem.Id }, 
                                    XYZ.Zero);
                                
                                if (copiedIds.Count > 0)
                                {
                                    copiedElementIds.Add(copiedIds.First());
                                }
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"Failed to copy element {elem.Id}: {ex.Message}");
                            }
                        }

                        if (copiedElementIds.Count == 0)
                        {
                            errors.Add($"Failed to copy any elements for group '{originalGroupName}'");
                            continue;
                        }

                        // Ungroup
                        var ungroupedIds = targetGroup.UngroupMembers();
                        
                        // Verify ungrouped elements
                        var validUngroupedIds = new List<ElementId>();
                        foreach (var id in ungroupedIds)
                        {
                            var elem = doc.GetElement(id);
                            if (elem != null && elem.IsValidObject)
                            {
                                validUngroupedIds.Add(id);
                            }
                        }

                        // Combine ungrouped elements with copied elements
                        var allElementIds = new List<ElementId>(validUngroupedIds);
                        allElementIds.AddRange(copiedElementIds);
                        
                        // Create new group
                        Group newGroup = doc.Create.NewGroup(allElementIds);

                        // Set new group type name
                        var newGroupType = newGroup.GroupType;
                        try
                        {
                            newGroupType.Name = originalGroupName + "_Updated";
                        }
                        catch { }
                        
                        // Restore parameters
                        RestoreGroupParameters(newGroup, parameterValues);
                        
                        // Restore location
                        if (groupLocation != null && newGroup.Location is LocationPoint newLocPoint)
                        {
                            // Move to original position
                            var translation = groupLocation - newLocPoint.Point;
                            if (!translation.IsZeroLength())
                            {
                                ElementTransformUtils.MoveElement(doc, newGroup.Id, translation);
                            }
                            
                            // Restore rotation if supported
                            if (hasRotation)
                            {
                                try
                                {
                                    double currentRotation = 0;
                                    try { currentRotation = newLocPoint.Rotation; } catch { }
                                    
                                    if (Math.Abs(groupRotation - currentRotation) > 0.0001)
                                    {
                                        var axis = Line.CreateBound(groupLocation, groupLocation + XYZ.BasisZ);
                                        var rotationAngle = groupRotation - currentRotation;
                                        ElementTransformUtils.RotateElement(doc, newGroup.Id, axis, rotationAngle);
                                    }
                                }
                                catch { }
                            }
                        }

                        processedGroups.Add(newGroup.Id);
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Failed to process group '{kvp.Key.GroupType.Name}': {ex.Message}");
                    }
                }

                // Delete the original elements (cleanup)
                try
                {
                    var originalElementIds = elementsToCheck.Select(e => e.Id).ToList();
                    doc.Delete(originalElementIds);
                }
                catch { }

                trans.Commit();

                // Update selection to show the new groups
                uidoc.SetSelectionIds(processedGroups);

                // Show results
                var resultMessage = $"Successfully processed {processedGroups.Count} group(s).";
                if (errors.Any())
                {
                    resultMessage += $"\n\nWarnings:\n" + string.Join("\n", errors);
                }
                
                TaskDialog.Show("Operation Complete", resultMessage);
            }
            catch (Exception ex)
            {
                trans.RollBack();
                TaskDialog.Show("Error", $"Failed to add elements to groups: {ex.Message}");
                return Result.Failed;
            }
        }

        return Result.Succeeded;
    }

    // Get the bounding box of a group
    private BoundingBoxXYZ GetGroupBoundingBox(Group group)
    {
        var bb = group.get_BoundingBox(null);
        if (bb != null) return bb;

        // If group doesn't have a bounding box, calculate from members
        var memberIds = group.GetMemberIds();
        if (!memberIds.Any()) return null;

        var doc = group.Document;
        BoundingBoxXYZ combinedBB = null;

        foreach (var memberId in memberIds)
        {
            var member = doc.GetElement(memberId);
            if (member == null) continue;

            var memberBB = member.get_BoundingBox(null);
            if (memberBB == null) continue;

            if (combinedBB == null)
            {
                combinedBB = new BoundingBoxXYZ();
                combinedBB.Min = memberBB.Min;
                combinedBB.Max = memberBB.Max;
            }
            else
            {
                // Expand bounding box to include this member
                combinedBB.Min = new XYZ(
                    Math.Min(combinedBB.Min.X, memberBB.Min.X),
                    Math.Min(combinedBB.Min.Y, memberBB.Min.Y),
                    Math.Min(combinedBB.Min.Z, memberBB.Min.Z)
                );
                combinedBB.Max = new XYZ(
                    Math.Max(combinedBB.Max.X, memberBB.Max.X),
                    Math.Max(combinedBB.Max.Y, memberBB.Max.Y),
                    Math.Max(combinedBB.Max.Z, memberBB.Max.Z)
                );
            }
        }

        return combinedBB;
    }

    // Check if two bounding boxes intersect or if bb2 is contained within bb1
    private bool BoundingBoxIntersectsOrContains(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        // Check if bb2 is completely contained within bb1
        bool contained = bb2.Min.X >= bb1.Min.X && bb2.Max.X <= bb1.Max.X &&
                        bb2.Min.Y >= bb1.Min.Y && bb2.Max.Y <= bb1.Max.Y &&
                        bb2.Min.Z >= bb1.Min.Z && bb2.Max.Z <= bb1.Max.Z;
        
        if (contained) return true;

        // Check for intersection
        bool intersects = !(bb2.Max.X < bb1.Min.X || bb2.Min.X > bb1.Max.X ||
                           bb2.Max.Y < bb1.Min.Y || bb2.Min.Y > bb1.Max.Y ||
                           bb2.Max.Z < bb1.Min.Z || bb2.Min.Z > bb1.Max.Z);
        
        return intersects;
    }

    // Store group parameters
    private Dictionary<string, object> StoreGroupParameters(Group group)
    {
        var parameterValues = new Dictionary<string, object>();
        foreach (Parameter param in group.Parameters)
        {
            if (!param.IsReadOnly && param.HasValue)
            {
                try
                {
                    string paramName = param.Definition.Name;
                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            parameterValues[paramName] = param.AsDouble();
                            break;
                        case StorageType.Integer:
                            parameterValues[paramName] = param.AsInteger();
                            break;
                        case StorageType.String:
                            parameterValues[paramName] = param.AsString();
                            break;
                        case StorageType.ElementId:
                            parameterValues[paramName] = param.AsElementId();
                            break;
                    }
                }
                catch { }
            }
        }
        return parameterValues;
    }

    // Restore group parameters
    private void RestoreGroupParameters(Group group, Dictionary<string, object> parameterValues)
    {
        foreach (var kvp in parameterValues)
        {
            try
            {
                Parameter newParam = group.LookupParameter(kvp.Key);
                if (newParam != null && !newParam.IsReadOnly)
                {
                    switch (newParam.StorageType)
                    {
                        case StorageType.Double:
                            newParam.Set((double)kvp.Value);
                            break;
                        case StorageType.Integer:
                            newParam.Set((int)kvp.Value);
                            break;
                        case StorageType.String:
                            string strValue = kvp.Value as string;
                            if (strValue != null)
                                newParam.Set(strValue);
                            break;
                        case StorageType.ElementId:
                            newParam.Set((ElementId)kvp.Value);
                            break;
                    }
                }
            }
            catch { }
        }
    }

    // Helper method to check if an element can be added to a group
    private bool CanBeAddedToGroup(Element element)
    {
        // Filter out elements that cannot be grouped
        if (element == null) return false;
        
        // Skip element types
        if (element is ElementType) return false;
        
        // Skip view-specific elements
        if (element.OwnerViewId != ElementId.InvalidElementId) return false;
        
        // Skip groups themselves
        if (element is Group) return false;
        
        // Skip elements that are already in a group
        if (element.GroupId != ElementId.InvalidElementId) return false;
        
        // Skip certain categories that cannot be grouped
        var category = element.Category;
        if (category != null)
        {
            var builtInCat = (BuiltInCategory)category.Id.IntegerValue;
            
            // Add categories that should not be grouped
            var excludedCategories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Views,
                BuiltInCategory.OST_Sheets,
                BuiltInCategory.OST_IOSModelGroups,
                BuiltInCategory.OST_IOSDetailGroups,
                BuiltInCategory.OST_Constraints,
                BuiltInCategory.OST_RvtLinks,
                BuiltInCategory.OST_Cameras,
                BuiltInCategory.OST_Elev,
                BuiltInCategory.OST_GridChains,
                BuiltInCategory.OST_SectionBox,
                BuiltInCategory.OST_RoomSeparationLines,
                BuiltInCategory.OST_MEPSpaceSeparationLines,
                BuiltInCategory.OST_AreaSchemeLines
            };
            
            if (excludedCategories.Contains(builtInCat))
                return false;
        }
        
        return true;
    }
}
