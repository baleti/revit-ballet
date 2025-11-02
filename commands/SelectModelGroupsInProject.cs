using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectModelGroupsInProject : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

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

        // Get all model group instances and count them by type in one pass
        var allGroupInstances = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                                .WhereElementIsNotElementType()
                                .Cast<Group>()
                                .ToList();

        // Create a dictionary to store instance counts by GroupType Id
        var instanceCountByTypeId = new Dictionary<ElementId, int>();
        foreach (var instance in allGroupInstances)
        {
            var typeId = instance.GroupType.Id;
            if (instanceCountByTypeId.ContainsKey(typeId))
            {
                instanceCountByTypeId[typeId]++;
            }
            else
            {
                instanceCountByTypeId[typeId] = 1;
            }
        }

        // Prepare entries for the DataGrid
        var entries = new List<Dictionary<string, object>>();
        foreach (var groupType in modelGroupTypes)
        {
            // Get count from dictionary (0 if not found)
            int instanceCount = instanceCountByTypeId.ContainsKey(groupType.Id) 
                                ? instanceCountByTypeId[groupType.Id] 
                                : 0;

            var entry = new Dictionary<string, object>
            {
                { "Group Name", groupType.Name },
                { "Instances", instanceCount }
            };
            entries.Add(entry);
        }

        // Define the columns to display in the DataGrid
        var propertyNames = new List<string> { "Group Name", "Instances" };

        // Prompt the user to select one or more group types using the custom DataGrid GUI
        var selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);

        if (selectedEntries == null || selectedEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No model group type selected.");
            return Result.Cancelled;
        }

        // Get the current selection of elements in the model
        var selection = uidoc.Selection;
        var currentSelectionIds = uidoc.GetSelectionIds();

        // Iterate over all selected group types
        foreach (var selectedEntry in selectedEntries)
        {
            string selectedGroupName = (string)selectedEntry["Group Name"];

            // Find the corresponding GroupType by name
            GroupType selectedGroupType = modelGroupTypes.FirstOrDefault(g => g.Name == selectedGroupName);

            if (selectedGroupType == null)
            {
                TaskDialog.Show("Error", $"Unable to find the model group type: {selectedGroupName}");
                return Result.Failed;
            }

            // Find instances of the selected GroupType in the model
            var groupInstances = allGroupInstances
                                    .Where(g => g.GroupType.Id == selectedGroupType.Id)
                                    .ToList();

            if (groupInstances.Count == 0)
            {
                TaskDialog.Show("Error", $"No instances of the selected group type: {selectedGroupType.Name} found in the project.");
                continue; // Continue checking the other selected groups
            }

            // Add the new group instances to the current selection
            var groupInstanceIds = groupInstances.Select(g => g.Id).ToList();
            foreach (var id in groupInstanceIds)
            {
                if (!currentSelectionIds.Contains(id))
                {
                    currentSelectionIds.Add(id);
                }
            }
        }

        // Update the selection with the combined set of elements
        uidoc.SetSelectionIds(currentSelectionIds);

        return Result.Succeeded;
    }
}
