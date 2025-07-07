using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.ReadOnly)]
public class FilterPosition : IExternalCommand
{
    // Helper class to store element data along with its reference
    private class ElementDataWithReference
    {
        public Dictionary<string, object> Data { get; set; }
        public Reference Reference { get; set; }
        public Element Element { get; set; }
        public bool IsLinked { get; set; }
        public string DocumentName { get; set; }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the active UIDocument and Document.
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Retrieve the current selection using the SelectionModeManager which supports linked references
        IList<Reference> selectedRefs = uidoc.GetReferences();
        if (selectedRefs == null || !selectedRefs.Any())
        {
            // Try to get regular selection IDs as fallback
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            if (selectedIds == null || !selectedIds.Any())
            {
                TaskDialog.Show("Warning", "Please select elements before running the command.");
                return Result.Failed;
            }
            
            // Convert ElementIds to References for consistency
            selectedRefs = selectedIds.Select(id => new Reference(doc.GetElement(id))).ToList();
        }

        // Process references to get elements (both from current doc and linked docs)
        List<ElementDataWithReference> elementDataList = new List<ElementDataWithReference>();

        foreach (Reference reference in selectedRefs)
        {
            Element elem = null;
            bool isLinked = false;
            string documentName = doc.Title;

            try
            {
                if (reference.LinkedElementId != ElementId.InvalidElementId)
                {
                    // This is a linked element
                    isLinked = true;
                    RevitLinkInstance linkInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkInstance != null)
                    {
                        Document linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            elem = linkedDoc.GetElement(reference.LinkedElementId);
                            documentName = linkedDoc.Title;
                        }
                    }
                }
                else
                {
                    // Regular element in current document
                    elem = doc.GetElement(reference.ElementId);
                }

                if (elem != null)
                {
                    elementDataList.Add(new ElementDataWithReference
                    {
                        Element = elem,
                        Reference = reference,
                        IsLinked = isLinked,
                        DocumentName = documentName
                    });
                }
            }
            catch
            {
                // Skip problematic references
                continue;
            }
        }

        if (!elementDataList.Any())
        {
            TaskDialog.Show("Warning", "No valid elements found in selection.");
            return Result.Failed;
        }

        // Define the property names (columns) for the data grid.
        List<string> propertyNames = new List<string>
        {
            "Element Id",
            "Document",
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

        // Prepare the list of dictionaries that will hold each element's orientation data.
        List<Dictionary<string, object>> gridData = new List<Dictionary<string, object>>();

        foreach (var elemData in elementDataList)
        {
            Element elem = elemData.Element;
            Dictionary<string, object> properties = new Dictionary<string, object>();

            // Basic element properties.
            properties["Element Id"] = elem.Id.IntegerValue;
            properties["Document"] = elemData.DocumentName + (elemData.IsLinked ? " (Linked)" : "");
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

            // Store the data along with the index for later reference
            elemData.Data = properties;
            gridData.Add(properties);
        }

        // Display the data grid. (Assuming CustomGUIs.DataGrid is a helper method that
        // shows a grid for the user to review and optionally modify the selection.)
        List<Dictionary<string, object>> selectedFromGrid = CustomGUIs.DataGrid(gridData, propertyNames, false);

        // If the user made a selection in the grid, update the active selection.
        if (selectedFromGrid?.Any() == true)
        {
            // Build a new list of references based on the selected items
            List<Reference> finalReferences = new List<Reference>();

            foreach (var selectedData in selectedFromGrid)
            {
                // Find the matching element data by comparing element ID and document name
                var matchingElemData = elementDataList.FirstOrDefault(ed =>
                    (int)ed.Data["Element Id"] == (int)selectedData["Element Id"] &&
                    ed.Data["Document"].ToString() == selectedData["Document"].ToString());

                if (matchingElemData != null)
                {
                    finalReferences.Add(matchingElemData.Reference);
                }
            }

            // Update the selection using SetReferences to maintain linked element references
            uidoc.SetReferences(finalReferences);
        }

        return Result.Succeeded;
    }
}
