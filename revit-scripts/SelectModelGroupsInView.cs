using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectModelGroupsInView : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = doc.ActiveView;

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

        // Prepare entries for the DataGrid
        var entries = new List<Dictionary<string, object>>();
        foreach (var groupType in modelGroupTypes)
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

            // Find instances of the selected GroupType in the current view
            var groupInstances = new FilteredElementCollector(doc, activeView.Id)
                                    .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                                    .WhereElementIsNotElementType()
                                    .Cast<Group>()
                                    .Where(g => g.GroupType.Id == selectedGroupType.Id)
                                    .ToList();

            // Add the visible group instances to the current selection
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
