using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectByModelGroupsInViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Determine target views (selected views or active view)
        ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();
        List<View> targetViews = new List<View>();

        bool hasSelectedViews = currentSelection.Any(id =>
        {
            Element elem = doc.GetElement(id);
            return elem is View;
        });

        if (hasSelectedViews)
        {
            // Use selected views
            targetViews = currentSelection
                .Select(id => doc.GetElement(id))
                .OfType<View>()
                .ToList();
        }
        else
        {
            // Use current view only
            targetViews.Add(doc.ActiveView);
        }

        // Collect all model group instances visible in target views and count them by type
        var typeToInstancesMap = new Dictionary<ElementId, List<ElementId>>();

        foreach (View view in targetViews)
        {
            var groupInstances = new FilteredElementCollector(doc, view.Id)
                                    .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                                    .WhereElementIsNotElementType()
                                    .Cast<Group>()
                                    .ToList();

            foreach (var group in groupInstances)
            {
                ElementId typeId = group.GroupType.Id;
                if (!typeToInstancesMap.ContainsKey(typeId))
                {
                    typeToInstancesMap[typeId] = new List<ElementId>();
                }
                // Avoid duplicates when same instance is visible in multiple views
                if (!typeToInstancesMap[typeId].Contains(group.Id))
                {
                    typeToInstancesMap[typeId].Add(group.Id);
                }
            }
        }

        if (typeToInstancesMap.Count == 0)
        {
            TaskDialog.Show("Error", "No model groups found visible in the target view(s).");
            return Result.Failed;
        }

        // Get the group types that have instances in the target views
        var modelGroupTypes = typeToInstancesMap.Keys
            .Select(id => doc.GetElement(id) as GroupType)
            .Where(gt => gt != null)
            .ToList();

        // Prepare entries for the DataGrid with instance counts
        var entries = new List<Dictionary<string, object>>();
        foreach (var groupType in modelGroupTypes)
        {
            int instanceCount = typeToInstancesMap[groupType.Id].Count;
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
        var selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, false);

        if (selectedEntries == null || selectedEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No model group type selected.");
            return Result.Cancelled;
        }

        // Build the final selection set
        HashSet<ElementId> finalSelection = new HashSet<ElementId>(currentSelection);

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

            // Add the instances we already collected from the target views
            if (typeToInstancesMap.ContainsKey(selectedGroupType.Id))
            {
                foreach (var instanceId in typeToInstancesMap[selectedGroupType.Id])
                {
                    finalSelection.Add(instanceId);
                }
            }
        }

        // Update the selection with the combined set of elements
        uidoc.SetSelectionIds(finalSelection.ToList());

        return Result.Succeeded;
    }
}
