using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class CombineSelectedElementsWithSelectedModelGroups : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get current selection using the SelectionModeManager extension method
        var selectedIds = uidoc.GetSelectionIds();

        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "No elements selected. Please select elements and groups first.");
            return Result.Failed;
        }

        // Separate groups from other elements in the selection
        var groups = new List<Group>();
        var elementIdsToAdd = new List<ElementId>();

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
                    elementIdsToAdd.Add(id);
                }
            }
        }

        // Check if we have groups and elements
        if (groups.Count == 0)
        {
            TaskDialog.Show("Error", "No model groups found in selection. Please select at least one model group.");
            return Result.Failed;
        }

        if (elementIdsToAdd.Count == 0)
        {
            TaskDialog.Show("Error", "No valid elements to add to groups found in selection.");
            return Result.Failed;
        }

        // Determine which groups to add elements to
        List<Group> targetGroups;

        if (groups.Count == 1)
        {
            // Only one group selected, use it directly
            targetGroups = groups;
        }
        else
        {
            // Multiple groups selected, prompt user to choose
            var entries = new List<Dictionary<string, object>>();
            foreach (var group in groups)
            {
                var groupType = group.GroupType;
                var entry = new Dictionary<string, object>
                {
                    { "Group Name", groupType.Name },
                    { "Group Id", group.Id.IntegerValue },
                    { "Level", GetGroupLevel(doc, group) },
                    { "Member Count", group.GetMemberIds().Count }
                };
                entries.Add(entry);
            }

            // Define columns for DataGrid
            var propertyNames = new List<string> { "Group Name", "Group Id", "Level", "Member Count" };

            // Prompt user to select target groups
            var selectedEntries = CustomGUIs.DataGrid(
                entries, 
                propertyNames, 
                spanAllScreens: false
            );

            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                TaskDialog.Show("Info", "No groups selected. Operation cancelled.");
                return Result.Cancelled;
            }

            // Get the selected groups
            targetGroups = new List<Group>();
            foreach (var entry in selectedEntries)
            {
                int groupId = (int)entry["Group Id"];
                var group = groups.FirstOrDefault(g => g.Id.IntegerValue == groupId);
                if (group != null)
                {
                    targetGroups.Add(group);
                }
            }
        }

        // Show confirmation dialog
        var td = new TaskDialog("Confirm Operation")
        {
            MainContent = $"This will ungroup {targetGroups.Count} group(s) and create new groups with {elementIdsToAdd.Count} additional element(s).\n\nContinue?",
            MainInstruction = "Add Elements to Groups",
            CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
        };

        if (td.Show() != TaskDialogResult.Yes)
        {
            return Result.Cancelled;
        }

        // Process groups
        using (Transaction trans = new Transaction(doc, "Add Elements to Model Groups"))
        {
            trans.Start();

            int successCount = 0;
            var results = new List<string>();
            var newGroupIds = new List<ElementId>();

            foreach (var group in targetGroups)
            {
                string originalGroupName = "";
                try
                {
                    // Store original group info
                    originalGroupName = group.GroupType.Name;
                    
                    // Ungroup - this returns the ungrouped element IDs
                    var ungroupedIds = group.UngroupMembers();
                    
                    // Create a new list combining ungrouped elements with new elements
                    var allElementIds = new List<ElementId>(ungroupedIds);
                    
                    // Add new elements, avoiding duplicates
                    foreach (var newId in elementIdsToAdd)
                    {
                        if (!allElementIds.Contains(newId))
                        {
                            allElementIds.Add(newId);
                        }
                    }
                    
                    // Create new group
                    Group newGroup = doc.Create.NewGroup(allElementIds);
                    
                    // Try to set a meaningful name for the new group type
                    try
                    {
                        string newGroupName = originalGroupName;
                        
                        // Check if the original name is already taken by another group type
                        var existingTypes = new FilteredElementCollector(doc)
                            .OfClass(typeof(GroupType))
                            .Cast<GroupType>()
                            .Where(gt => gt.Category != null && 
                                   gt.Category.Id.IntegerValue == (int)BuiltInCategory.OST_IOSModelGroups)
                            .Select(gt => gt.Name)
                            .ToList();
                        
                        // If name exists, append a number
                        if (existingTypes.Contains(newGroupName))
                        {
                            int counter = 2;
                            while (existingTypes.Contains($"{originalGroupName}_{counter}"))
                            {
                                counter++;
                            }
                            newGroupName = $"{originalGroupName}_{counter}";
                        }
                        
                        newGroup.GroupType.Name = newGroupName;
                    }
                    catch
                    {
                        // If renaming fails, continue with auto-generated name
                    }
                    
                    newGroupIds.Add(newGroup.Id);
                    successCount++;
                    results.Add($"✓ {originalGroupName} → {newGroup.GroupType.Name} (ID: {newGroup.Id})");
                }
                catch (System.Exception ex)
                {
                    results.Add($"✗ {originalGroupName}: {ex.Message}");
                }
            }

            trans.Commit();

            // Update selection to show the new groups
            if (newGroupIds.Count > 0)
            {
                uidoc.SetSelectionIds(newGroupIds);
            }

            // Show results
            string resultMessage = $"Operation Complete\n\n";
            resultMessage += $"Successfully created {successCount} new group(s) from {targetGroups.Count} original group(s).\n\n";
            resultMessage += "Details:\n" + string.Join("\n", results);

            TaskDialog.Show("Add Elements to Groups", resultMessage);
        }

        return Result.Succeeded;
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

    // Helper method to get group level
    private string GetGroupLevel(Document doc, Group group)
    {
        try
        {
            // Try to get level from parameter
            var levelParam = group.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
            if (levelParam != null && levelParam.HasValue)
            {
                var level = doc.GetElement(levelParam.AsElementId()) as Level;
                if (level != null)
                    return level.Name;
            }
            
            // Try base constraint parameter
            var baseConstraintParam = group.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM);
            if (baseConstraintParam != null && baseConstraintParam.HasValue)
            {
                var level = doc.GetElement(baseConstraintParam.AsElementId()) as Level;
                if (level != null)
                    return level.Name;
            }
            
            // Try to find level from group members
            var memberIds = group.GetMemberIds();
            foreach (var memberId in memberIds)
            {
                var member = doc.GetElement(memberId);
                if (member != null)
                {
                    var memberLevelParam = member.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                    if (memberLevelParam != null && memberLevelParam.HasValue)
                    {
                        var level = doc.GetElement(memberLevelParam.AsElementId()) as Level;
                        if (level != null)
                            return level.Name;
                    }
                }
            }
        }
        catch { }
        
        return "N/A";
    }
}
