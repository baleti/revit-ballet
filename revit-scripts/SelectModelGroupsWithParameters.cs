using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectModelGroupsWithParameters : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Step 1: Prepare a list of unique model groups in the current view
        List<Dictionary<string, object>> groupEntries = new List<Dictionary<string, object>>();
        Dictionary<string, Group> groupElementMap = new Dictionary<string, Group>(); // Map unique keys to Groups

        // Get the current view's Id
        ElementId currentViewId = doc.ActiveView.Id;

        // Collect all visible Group elements in the current view
        var groupsInView = new FilteredElementCollector(doc, currentViewId)
            .OfClass(typeof(Group))
            .Cast<Group>()
            .Where(g => IsElementVisibleInView(g, doc.ActiveView))
            .ToList();

        // Distinguish groups by their unique Id since names can be duplicated
        foreach (Group grp in groupsInView)
        {
            GroupType grpType = grp.GroupType; // Directly access the GroupType

            var entry = new Dictionary<string, object>
            {
                { "Group Name", grp.Name },
                { "Group Type", grpType != null ? grpType.Name : "N/A" },
                { "Category", grp.Category?.Name ?? "N/A" }
            };

            string uniqueKey = $"{grpType?.Name ?? "Unknown"}:{grp.Name}:{grp.Id}";
            groupElementMap[uniqueKey] = grp;

            groupEntries.Add(entry);
        }

        // Step 2: Display the list of groups using CustomGUIs.DataGrid
        var propertyNames = new List<string> { "Group Name", "Group Type", "Category" };
        var selectedEntries = CustomGUIs.DataGrid(groupEntries, propertyNames, spanAllScreens: false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        // Step 3: Collect ElementIds of the selected groups
        List<ElementId> selectedGroupIds = selectedEntries
            .Select(entry =>
            {
                string uniqueKey = $"{entry["Group Type"]}:{entry["Group Name"]}:{FindMatchingIdInMap(groupElementMap, entry)}";
                return groupElementMap[uniqueKey].Id;
            })
            .ToList();

        // Step 4: Retrieve and display all available parameters for the selected groups
        List<Dictionary<string, object>> parameterEntries = new List<Dictionary<string, object>>();

        foreach (ElementId grpId in selectedGroupIds)
        {
            Group grp = doc.GetElement(grpId) as Group;
            if (grp != null)
            {
                GroupType grpType = grp.GroupType; // Directly access the GroupType

                var entry = new Dictionary<string, object>
                {
                    { "Group Name", grp.Name },
                    { "Group Type", grpType?.Name ?? "N/A" },
                    { "Category", grp.Category?.Name ?? "N/A" }
                };

                // Add all available parameters from the group itself
                foreach (Parameter param in grp.Parameters)
                {
                    string paramName = param.Definition.Name;
                    string paramValue = param.AsValueString() ?? param.AsString() ?? "None";
                    entry[paramName] = paramValue;
                }

                // Optionally, also add parameters from the GroupType
                if (grpType != null)
                {
                    foreach (Parameter param in grpType.Parameters)
                    {
                        string paramName = "Type - " + param.Definition.Name;
                        string paramValue = param.AsValueString() ?? param.AsString() ?? "None";
                        entry[paramName] = paramValue;
                    }
                }

                parameterEntries.Add(entry);
            }
        }

        // Step 5: Display the parameters of the selected groups in a second DataGrid
        var paramPropertyNames = parameterEntries.FirstOrDefault()?.Keys.ToList();
        var finalSelection = CustomGUIs.DataGrid(parameterEntries, paramPropertyNames, spanAllScreens: false);

        if (finalSelection.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        // Step 6: Select the final chosen groups
        List<ElementId> finalSelectedGroupIds = finalSelection
            .Select(entry =>
            {
                string uniqueKey = $"{entry["Group Type"]}:{entry["Group Name"]}:{FindMatchingIdInMap(groupElementMap, entry)}";
                return groupElementMap[uniqueKey].Id;
            })
            .ToList();

        uidoc.Selection.SetElementIds(finalSelectedGroupIds);

        return Result.Succeeded;
    }

    // Helper method to match the chosen entry back to the unique key in the map
    private string FindMatchingIdInMap(Dictionary<string, Group> groupMap, Dictionary<string, object> entry)
    {
        string grpType = entry["Group Type"].ToString();
        string grpName = entry["Group Name"].ToString();

        // Attempt to find a single matching key in the dictionary. 
        // Keys are in the form: "GroupTypeName:GroupName:ElementId"
        var match = groupMap.Keys.FirstOrDefault(k => k.StartsWith($"{grpType}:{grpName}:"));
        if (match != null)
        {
            // Extract the element Id part
            var parts = match.Split(':');
            if (parts.Length == 3)
                return parts[2];
        }
        return string.Empty;
    }

    // Helper method to check if an element is visible in a given view
    private bool IsElementVisibleInView(Element element, View view)
    {
        BoundingBoxXYZ bbox = element.get_BoundingBox(view);
        return bbox != null;
    }
}
