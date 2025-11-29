using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SwapSelectedModelGroups : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Check if active view supports placing model groups
        View activeView = doc.ActiveView;
        ViewType viewType = activeView.ViewType;
        if (viewType == ViewType.Legend || viewType == ViewType.DrawingSheet || viewType == ViewType.Schedule ||
            viewType == ViewType.PanelSchedule || viewType == ViewType.ColumnSchedule ||
            viewType == ViewType.ProjectBrowser || viewType == ViewType.SystemBrowser ||
            viewType == ViewType.Report || viewType == ViewType.Walkthrough || viewType == ViewType.Rendering)
        {
            TaskDialog.Show("Error", "This command must be run from a model view such as a floor plan, ceiling plan, elevation, section, detail, or 3D view.");
            return Result.Failed;
        }

        // Get current selection using the SelectionModeManager extension method
        var selectedIds = uidoc.GetSelectionIds();

        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "No elements selected. Please select groups to swap.");
            return Result.Failed;
        }

        // Get selected groups
        var selectedGroups = new List<Group>();
        foreach (var id in selectedIds)
        {
            var element = doc.GetElement(id);
            if (element is Group group && group.GroupType != null)
            {
                // Only add model groups (not detail groups)
                if (group.Category != null && 
                    group.Category.Id.AsLong() == (int)BuiltInCategory.OST_IOSModelGroups)
                {
                    selectedGroups.Add(group);
                }
            }
        }

        if (selectedGroups.Count == 0)
        {
            TaskDialog.Show("Error", "No model groups found in selection.");
            return Result.Failed;
        }

        // Get all model group types in the project
        var modelGroupTypes = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                                .WhereElementIsElementType()
                                .Cast<GroupType>()
                                .ToList();

        if (modelGroupTypes.Count == 0)
        {
            TaskDialog.Show("Error", "No model group types found in the project.");
            return Result.Failed;
        }

        // Get current group types from selection
        var currentTypeIds = selectedGroups.Select(g => g.GroupType.Id).Distinct().ToList();
        
        // Prepare entries for the DataGrid
        var entries = new List<Dictionary<string, object>>();
        foreach (var groupType in modelGroupTypes)
        {
            // Count instances of this type
            var instanceCount = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsNotElementType()
                .Cast<Group>()
                .Count(g => g.GroupType.Id == groupType.Id);

            // Get member count from first instance if available
            int memberCount = 0;
            var firstInstance = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsNotElementType()
                .Cast<Group>()
                .FirstOrDefault(g => g.GroupType.Id == groupType.Id);
            
            if (firstInstance != null)
            {
                memberCount = firstInstance.GetMemberIds().Count;
            }

            bool isCurrentType = currentTypeIds.Contains(groupType.Id);
            string displayName = isCurrentType ? $"[CURRENT] {groupType.Name}" : groupType.Name;

            var entry = new Dictionary<string, object>
            {
                { "Group Name", displayName },
                { "Instance Count", instanceCount },
                { "Member Count", memberCount },
                { "Original Name", groupType.Name } // Store original name for lookup
            };
            entries.Add(entry);
        }

        // Define the columns to display in the DataGrid
        var propertyNames = new List<string> { "Group Name", "Instance Count", "Member Count" };

        // Show info about what will be swapped
        var currentTypes = selectedGroups.Select(g => g.GroupType.Name).Distinct().ToList();
        string currentTypesStr = string.Join(", ", currentTypes);
        
        TaskDialog.Show("Select Replacement Group", 
            $"Swapping {selectedGroups.Count} group(s) of type(s):\n{currentTypesStr}\n\nSelect a replacement group type.");
        
        // Prompt the user to select a replacement group type
        var selectedEntries = CustomGUIs.DataGrid(
            entries, 
            propertyNames, 
            false
        );

        if (selectedEntries == null || selectedEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No model group type selected.");
            return Result.Cancelled;
        }

        if (selectedEntries.Count > 1)
        {
            TaskDialog.Show("Error", "Please select only one replacement group type.");
            return Result.Failed;
        }

        // Get the selected replacement group type
        string selectedGroupName = (string)selectedEntries.First()["Original Name"];
        GroupType replacementGroupType = modelGroupTypes.FirstOrDefault(g => g.Name == selectedGroupName);

        if (replacementGroupType == null)
        {
            TaskDialog.Show("Error", $"Unable to find the model group type: {selectedGroupName}");
            return Result.Failed;
        }

        // Show confirmation dialog
        var td = new TaskDialog("Confirm Group Swap")
        {
            MainContent = $"Swap {selectedGroups.Count} group(s) to:\n'{replacementGroupType.Name}'?\n\nAll instance parameters and locations will be preserved.",
            MainInstruction = "Swap Model Groups",
            CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
        };

        if (td.Show() != TaskDialogResult.Yes)
        {
            return Result.Cancelled;
        }

        // Perform the swap
        using (Transaction trans = new Transaction(doc, "Swap Model Groups"))
        {
            trans.Start();

            int successCount = 0;
            var results = new List<string>();
            var newGroupIds = new List<ElementId>();

            foreach (var group in selectedGroups)
            {
                string originalGroupName = group.GroupType.Name;
                try
                {
                    // Store all instance parameters before swapping
                    var parameterValues = new Dictionary<string, object>();
                    foreach (Parameter param in group.Parameters)
                    {
                        if (!param.IsReadOnly && param.HasValue)
                        {
                            try
                            {
                                string paramName = param.Definition.Name;
                                
                                // Store parameter value based on storage type
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
                            catch { /* Skip parameters that can't be read */ }
                        }
                    }
                    
                    // Store location info
                    XYZ groupLocation = null;
                    double groupRotation = 0;
                    bool supportsRotation = false;
                    
                    if (group.Location is LocationPoint locPoint)
                    {
                        groupLocation = locPoint.Point;
                        
                        // Check if rotation is supported
                        try
                        {
                            groupRotation = locPoint.Rotation;
                            supportsRotation = true;
                        }
                        catch
                        {
                            // Rotation not supported for this group
                            supportsRotation = false;
                        }
                    }
                    if (groupLocation == null)
                    {
                        throw new Exception("Group has no valid location.");
                    }
                    
                    // Create new group instance of the replacement type
                    Group newGroup = doc.Create.PlaceGroup(groupLocation, replacementGroupType);
                    
                    // Restore all instance parameters
                    foreach (var kvp in parameterValues)
                    {
                        try
                        {
                            Parameter newParam = newGroup.LookupParameter(kvp.Key);
                            
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
                        catch { /* Skip parameters that can't be set */ }
                    }
                    
                    // Restore rotation if supported
                    if (supportsRotation && newGroup.Location is LocationPoint newLocPoint)
                    {
                        try
                        {
                            var currentRotation = newLocPoint.Rotation;
                            if (Math.Abs(groupRotation - currentRotation) > 0.0001)
                            {
                                var axis = Line.CreateBound(groupLocation, groupLocation + XYZ.BasisZ);
                                var rotationAngle = groupRotation - currentRotation;
                                ElementTransformUtils.RotateElement(doc, newGroup.Id, axis, rotationAngle);
                            }
                        }
                        catch
                        {
                            // If rotation fails, skip it silently
                        }
                    }
                    
                    // Delete the old group
                    var deletedIds = doc.Delete(group.Id);
                    if (deletedIds.Count == 0)
                    {
                        throw new Exception("Failed to delete original group.");
                    }
                    
                    newGroupIds.Add(newGroup.Id);
                    successCount++;
                    results.Add($"✓ {originalGroupName} → {replacementGroupType.Name} (ID: {newGroup.Id})");
                }
                catch (System.Exception ex)
                {
                    results.Add($"✗ {originalGroupName} (ID: {group.Id}): {ex.Message}");
                }
            }

            trans.Commit();

            // Update selection to show the new groups
            if (newGroupIds.Count > 0)
            {
                uidoc.Selection.SetElementIds(newGroupIds);
            }

            // Show results
            string resultMessage = $"Operation Complete\n\n";
            resultMessage += $"Successfully swapped {successCount} of {selectedGroups.Count} group(s).\n\n";
            resultMessage += "Details:\n" + string.Join("\n", results);

            TaskDialog.Show("Swap Model Groups", resultMessage);
        }

        return Result.Succeeded;
    }
}
