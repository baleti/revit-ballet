using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.ReadOnly)]
public class FilterDoors : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get and filter selected elements
        ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
        if (!selectedIds.Any())
        {
            TaskDialog.Show("Warning", "Please select doors before running the command.");
            return Result.Failed;
        }

        List<Element> selectedDoors = selectedIds
            .Select(id => doc.GetElement(id))
            .Where(e => e?.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
            .ToList();

        if (!selectedDoors.Any())
        {
            TaskDialog.Show("Warning", "No doors in current selection.");
            return Result.Failed;
        }

        // Collect shared parameters once before processing
        var sharedParams = selectedDoors
            .SelectMany(d => d.Parameters.OfType<Parameter>())
            .Where(p => p.IsShared)
            .Select(p => p.Definition.Name)
            .Distinct()
            .OrderBy(n => n)
            .ToList();

        // Prepare data structures
        List<string> propertyNames = new List<string>
        {
            "Element Id", "Family Name", "Instance Name", "Level", "Mark",
            "Head Height", "Comments", "Group", "FacingFlipped", "HandFlipped",
            "Width", "Height", "Room From", "Room To"
        };
        propertyNames.AddRange(sharedParams);

        List<Dictionary<string, object>> doorData = new List<Dictionary<string, object>>();

        foreach (Element door in selectedDoors)
        {
            if (!(door is FamilyInstance doorInst)) continue;

            Dictionary<string, object> doorProperties = new Dictionary<string, object>();
            ElementType doorType = doc.GetElement(door.GetTypeId()) as ElementType;

            // Built-in parameters
            doorProperties["Element Id"] = door.Id.IntegerValue;
            doorProperties["Family Name"] = doorType?.FamilyName ?? "";
            doorProperties["Instance Name"] = door.Name;
            doorProperties["Level"] = doc.GetElement(door.LevelId)?.Name ?? "";
            doorProperties["Mark"] = door.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString() ?? "";
            doorProperties["Head Height"] = door.get_Parameter(BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM)?.AsDouble() ?? 0.0;
            doorProperties["Comments"] = door.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.AsString() ?? "";
            doorProperties["Group"] = door.GroupId != ElementId.InvalidElementId ? doc.GetElement(door.GroupId)?.Name : "";
            doorProperties["FacingFlipped"] = doorInst.FacingFlipped;
            doorProperties["HandFlipped"] = doorInst.HandFlipped;
            doorProperties["Width"] = doorType?.get_Parameter(BuiltInParameter.DOOR_WIDTH)?.AsDouble() ?? 0.0;
            doorProperties["Height"] = doorType?.get_Parameter(BuiltInParameter.DOOR_HEIGHT)?.AsDouble() ?? 0.0;
            doorProperties["Room From"] = doorInst.FromRoom?.Name ?? "";
            doorProperties["Room To"] = doorInst.ToRoom?.Name ?? "";

            // Shared parameters
            foreach (var paramName in sharedParams)
            {
                var param = door.LookupParameter(paramName);
                doorProperties[paramName] = param?.AsValueString() ?? param?.AsString() ?? "";
            }

            doorData.Add(doorProperties);
        }

        // DataGrid handling
        List<Dictionary<string, object>> selectedFromGrid = CustomGUIs.DataGrid(doorData, propertyNames, false);
        
        if (selectedFromGrid?.Any() == true)
        {
            var finalSelection = selectedDoors
                .Where(d => selectedFromGrid.Any(s => 
                    (int)s["Element Id"] == d.Id.IntegerValue))
                .Select(d => d.Id)
                .ToList();
            
            uidoc.Selection.SetElementIds(finalSelection);
        }

        return Result.Succeeded;
    }
}
