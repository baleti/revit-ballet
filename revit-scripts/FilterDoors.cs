using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
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
        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
        if (!selectedIds.Any())
        {
            TaskDialog.Show("Warning", "Please select doors or windows before running the command.");
            return Result.Failed;
        }

        List<Element> selectedOpenings = selectedIds
            .Select(id => doc.GetElement(id))
            .Where(e => e?.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_Doors ||
                       e?.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
            .ToList();

        if (!selectedOpenings.Any())
        {
            TaskDialog.Show("Warning", "No doors or windows in current selection.");
            return Result.Failed;
        }

        // Collect all parameters (not just shared) once before processing
        var allParams = selectedOpenings
            .SelectMany(d => d.Parameters.OfType<Parameter>())
            .Select(p => p.Definition.Name)
            .Distinct()
            .OrderBy(n => n)
            .ToList();

        // Prepare data structures
        List<string> propertyNames = new List<string>
        {
            "Element Id", "Category", "Family Name", "Instance Name", "Level", "Mark",
            "Head Height", "Comments", "Group", "FacingFlipped", "HandFlipped", "Wall Direction",
            "Room From", "Room To", "Width", "Height"
        };
        propertyNames.AddRange(allParams);

        List<Dictionary<string, object>> openingData = new List<Dictionary<string, object>>();

        foreach (Element opening in selectedOpenings)
        {
            if (!(opening is FamilyInstance openingInst)) continue;

            Dictionary<string, object> openingProperties = new Dictionary<string, object>();
            ElementType openingType = doc.GetElement(opening.GetTypeId()) as ElementType;

            // Built-in parameters
            openingProperties["Element Id"] = opening.Id.IntegerValue;
            openingProperties["Category"] = opening.Category.Name;
            openingProperties["Family Name"] = openingType?.FamilyName ?? "";
            openingProperties["Instance Name"] = opening.Name;
            openingProperties["Level"] = doc.GetElement(opening.LevelId)?.Name ?? "";
            openingProperties["Mark"] = opening.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString() ?? "";
            openingProperties["Head Height"] = opening.get_Parameter(BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM)?.AsDouble() ?? 0.0;
            openingProperties["Comments"] = opening.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.AsString() ?? "";
            openingProperties["Group"] = opening.GroupId != ElementId.InvalidElementId ? doc.GetElement(opening.GroupId)?.Name : "";
            openingProperties["FacingFlipped"] = openingInst.FacingFlipped;
            openingProperties["HandFlipped"] = openingInst.HandFlipped;
            
            // Calculate Wall Direction
            openingProperties["Wall Direction"] = GetWallDirection(openingInst);
            
            openingProperties["Room From"] = openingInst.FromRoom?.Name ?? "";
            openingProperties["Room To"] = openingInst.ToRoom?.Name ?? "";
            
            // Width and Height parameters differ between doors and windows
            if (opening.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
            {
                openingProperties["Width"] = openingType?.get_Parameter(BuiltInParameter.DOOR_WIDTH)?.AsDouble() ?? 0.0;
                openingProperties["Height"] = openingType?.get_Parameter(BuiltInParameter.DOOR_HEIGHT)?.AsDouble() ?? 0.0;
            }
            else if (opening.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
            {
                openingProperties["Width"] = openingType?.get_Parameter(BuiltInParameter.WINDOW_WIDTH)?.AsDouble() ?? 0.0;
                openingProperties["Height"] = openingType?.get_Parameter(BuiltInParameter.WINDOW_HEIGHT)?.AsDouble() ?? 0.0;
            }

            // All parameters (built-in, shared, project, family)
            foreach (var paramName in allParams)
            {
                var param = opening.LookupParameter(paramName);
                if (param != null)
                {
                    // Use AsValueString() first for formatted values, then AsString() for text
                    openingProperties[paramName] = param.AsValueString() ?? param.AsString() ?? "None";
                }
                else
                {
                    openingProperties[paramName] = "";
                }
            }

            openingData.Add(openingProperties);
        }

        // DataGrid handling
        List<Dictionary<string, object>> selectedFromGrid = CustomGUIs.DataGrid(openingData, propertyNames, false);
        
        if (selectedFromGrid?.Any() == true)
        {
            var finalSelection = selectedOpenings
                .Where(o => selectedFromGrid.Any(s => 
                    (int)s["Element Id"] == o.Id.IntegerValue))
                .Select(o => o.Id)
                .ToList();
            
            uidoc.SetSelectionIds(finalSelection);
        }

        return Result.Succeeded;
    }
    
    private string GetWallDirection(FamilyInstance opening)
    {
        try
        {
            // Get the host wall
            Element host = opening.Host;
            if (!(host is Wall wall)) return "";
            
            // Get wall location curve
            LocationCurve wallLocation = wall.Location as LocationCurve;
            if (wallLocation == null) return "";
            
            Curve wallCurve = wallLocation.Curve;
            if (wallCurve == null) return "";
            
            // Get opening location
            LocationPoint openingLocation = opening.Location as LocationPoint;
            if (openingLocation == null) return "";
            
            XYZ openingPoint = openingLocation.Point;
            
            // Project opening point onto wall curve to get parameter
            IntersectionResult projection = wallCurve.Project(openingPoint);
            if (projection == null) return "";
            
            double parameter = projection.Parameter;
            
            // Get tangent at the parameter
            Transform transform = wallCurve.ComputeDerivatives(parameter, false);
            XYZ tangent = transform.BasisX.Normalize(); // Tangent is the first derivative
            
            // Use global coordinate system for consistent reference
            // Assume looking in positive Y direction (north), so right is positive X
            XYZ globalRight = XYZ.BasisX; // (1, 0, 0)
            
            // Check if wall tangent aligns with global right direction
            double dotProduct = tangent.DotProduct(globalRight);
            
            // If dot product is positive, wall extends to the right (positive X)
            // If dot product is negative, wall extends to the left (negative X)
            return dotProduct > 0 ? "Right" : "Left";
        }
        catch
        {
            return "";
        }
    }
}
