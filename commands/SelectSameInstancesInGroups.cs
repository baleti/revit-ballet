using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectSameInstancesInGroups : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;
            
            // Get currently selected elements using the extension method from SelectionModeManager
            var selectedIds = uidoc.GetSelectionIds();
            
            if (!selectedIds.Any())
            {
                TaskDialog.Show("Selection", "Please select at least one element that belongs to a group.");
                return Result.Cancelled;
            }
            
            // Dictionary to map GroupType to list of indices of selected elements within that group type
            var groupTypeToIndices = new Dictionary<ElementId, HashSet<int>>();
            
            // Dictionary to cache group instances by type
            var groupTypeToInstances = new Dictionary<ElementId, List<Group>>();
            
            // Process selected elements
            foreach (var selectedId in selectedIds)
            {
                var element = doc.GetElement(selectedId);
                if (element == null) continue;
                
                // Skip if not in a group
                if (element.GroupId == ElementId.InvalidElementId) continue;
                
                var group = doc.GetElement(element.GroupId) as Group;
                if (group == null) continue;
                
                var groupTypeId = group.GetTypeId();
                
                // Get member index efficiently
                var memberIds = group.GetMemberIds();
                int memberIndex = -1;
                
                // Find index by comparing ElementIds directly (much faster than loading elements)
                for (int i = 0; i < memberIds.Count; i++)
                {
                    if (memberIds.ElementAt(i) == selectedId)
                    {
                        memberIndex = i;
                        break;
                    }
                }
                
                if (memberIndex == -1) continue;
                
                // Store the member index for this group type
                if (!groupTypeToIndices.ContainsKey(groupTypeId))
                {
                    groupTypeToIndices[groupTypeId] = new HashSet<int>();
                }
                groupTypeToIndices[groupTypeId].Add(memberIndex);
                
                // Cache group instances for this type
                if (!groupTypeToInstances.ContainsKey(groupTypeId))
                {
                    // Get all instances of this group type in one query
                    groupTypeToInstances[groupTypeId] = new FilteredElementCollector(doc)
                        .OfClass(typeof(Group))
                        .Cast<Group>()
                        .Where(g => g.GetTypeId() == groupTypeId)
                        .ToList();
                }
            }
            
            if (!groupTypeToIndices.Any())
            {
                TaskDialog.Show("Selection", "No selected elements belong to groups.");
                return Result.Cancelled;
            }
            
            // Collect all matching elements
            var allElementIds = new HashSet<ElementId>(selectedIds);
            
            // For each group type with selected elements
            foreach (var kvp in groupTypeToIndices)
            {
                var groupTypeId = kvp.Key;
                var memberIndices = kvp.Value;
                var groupInstances = groupTypeToInstances[groupTypeId];
                
                // For each group instance of this type
                foreach (var groupInstance in groupInstances)
                {
                    var memberIds = groupInstance.GetMemberIds();
                    
                    // Add elements at the stored indices
                    foreach (int index in memberIndices)
                    {
                        if (index >= 0 && index < memberIds.Count)
                        {
                            allElementIds.Add(memberIds.ElementAt(index));
                        }
                    }
                }
            }
            
            // Set the new selection using the extension method from SelectionModeManager
            uidoc.SetSelectionIds(allElementIds.ToList());
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
