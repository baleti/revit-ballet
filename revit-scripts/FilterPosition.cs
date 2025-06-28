using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.ReadOnly)]
public class FilterPosition : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the active UIDocument and Document.
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Retrieve the current selection.
        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
        if (selectedIds == null || !selectedIds.Any())
        {
            TaskDialog.Show("Warning", "Please select elements before running the command.");
            return Result.Failed;
        }

        // Get the selected elements.
        List<Element> selectedElements = selectedIds.Select(id => doc.GetElement(id)).ToList();

        // Define the property names (columns) for the data grid.
        List<string> propertyNames = new List<string>
        {
            "Element Id",
            "Category",
            "Name",
            "LocationType",
            "X",
            "Y",
            "Z",
            "Rotation",
            "FacingFlipped",
            "HandFlipped",
            "Curve Start",
            "Curve End"
        };

        // Prepare the list of dictionaries that will hold each element’s orientation data.
        List<Dictionary<string, object>> elementData = new List<Dictionary<string, object>>();

        foreach (Element elem in selectedElements)
        {
            Dictionary<string, object> properties = new Dictionary<string, object>();

            // Basic element properties.
            properties["Element Id"] = elem.Id.IntegerValue;
            properties["Category"] = elem.Category != null ? elem.Category.Name : "";
            properties["Name"] = elem.Name;

            // Initialize orientation fields.
            properties["LocationType"] = "";
            properties["X"] = "";
            properties["Y"] = "";
            properties["Z"] = "";
            properties["Rotation"] = "";
            properties["FacingFlipped"] = "";
            properties["HandFlipped"] = "";
            properties["Curve Start"] = "";
            properties["Curve End"] = "";

            // Check if the element has a location.
            Location loc = elem.Location;
            if (loc != null)
            {
                // For elements with a point location.
                if (loc is LocationPoint locPoint)
                {
                    properties["LocationType"] = "Point";
                    properties["X"] = locPoint.Point.X;
                    properties["Y"] = locPoint.Point.Y;
                    properties["Z"] = locPoint.Point.Z;
                    properties["Rotation"] = locPoint.Rotation;
                }
                // For elements with a curve location.
                else if (loc is LocationCurve locCurve)
                {
                    properties["LocationType"] = "Curve";
                    Curve curve = locCurve.Curve;
                    if (curve != null)
                    {
                        XYZ start = curve.GetEndPoint(0);
                        XYZ end = curve.GetEndPoint(1);
                        properties["Curve Start"] = $"({start.X:F2}, {start.Y:F2}, {start.Z:F2})";
                        properties["Curve End"] = $"({end.X:F2}, {end.Y:F2}, {end.Z:F2})";
                    }
                }
            }

            // If the element is a FamilyInstance, record its flip properties.
            if (elem is FamilyInstance familyInstance)
            {
                properties["FacingFlipped"] = familyInstance.FacingFlipped;
                properties["HandFlipped"] = familyInstance.HandFlipped;
            }

            // (Optional) If the element is a View, and it has an active CropBox,
            // use the CropBox center as the location.
            if (elem is View view)
            {
                if (view.CropBoxActive && view.CropBox != null)
                {
                    XYZ center = (view.CropBox.Min + view.CropBox.Max) / 2;
                    properties["LocationType"] = "View CropBox Center";
                    properties["X"] = center.X;
                    properties["Y"] = center.Y;
                    properties["Z"] = center.Z;
                }
            }

            elementData.Add(properties);
        }

        // Display the data grid. (Assuming CustomGUIs.DataGrid is a helper method that
        // shows a grid for the user to review and optionally modify the selection.)
        List<Dictionary<string, object>> selectedFromGrid = CustomGUIs.DataGrid(elementData, propertyNames, false);

        // If the user made a selection in the grid, update the active selection.
        if (selectedFromGrid?.Any() == true)
        {
            var finalSelection = selectedElements
                .Where(e => selectedFromGrid.Any(s =>
                    (int)s["Element Id"] == e.Id.IntegerValue))
                .Select(e => e.Id)
                .ToList();

            uidoc.SetSelectionIds(finalSelection);
        }

        return Result.Succeeded;
    }
}
