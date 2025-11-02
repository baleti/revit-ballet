using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectDetailGroupsInCurrentView : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;
        View currentView = doc.ActiveView;

        // Get all detail group instances in the current view
        var detailGroupInstances = new FilteredElementCollector(doc, currentView.Id)
                                    .OfCategory(BuiltInCategory.OST_IOSDetailGroups)
                                    .WhereElementIsNotElementType()
                                    .Cast<Group>()
                                    .ToList();

        if (detailGroupInstances.Count == 0)
        {
            TaskDialog.Show("Error", "No detail groups found in the current view.");
            return Result.Failed;
        }

        // Get unique group types from the instances in the current view
        var detailGroupTypes = detailGroupInstances
                                .Select(g => g.GroupType)
                                .Distinct()
                                .ToList();

        // Prepare entries for the DataGrid
        var entries = new List<Dictionary<string, object>>();
        foreach (var groupType in detailGroupTypes)
        {
            var entry = new Dictionary<string, object>
            {
                { "Group Name", groupType.Name }
            };
            entries.Add(entry);
        }

        // Define the columns to display in the DataGrid
        var propertyNames = new List<string> { "Group Name" };

        // Prompt the user to select one or more group types using the custom DataGrid GUI
        var selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);

        if (selectedEntries == null || selectedEntries.Count == 0)
        {
            TaskDialog.Show("Info", "No detail group type selected.");
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
            GroupType selectedGroupType = detailGroupTypes.FirstOrDefault(g => g.Name == selectedGroupName);

            if (selectedGroupType == null)
            {
                TaskDialog.Show("Error", $"Unable to find the detail group type: {selectedGroupName}");
                return Result.Failed;
            }

            // Get instances of the selected GroupType in the current view
            var groupInstances = detailGroupInstances
                                .Where(g => g.GroupType.Id == selectedGroupType.Id)
                                .ToList();

            if (groupInstances.Count == 0)
            {
                TaskDialog.Show("Error", $"No instances of the selected group type: {selectedGroupType.Name} found in the current view.");
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
