using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.ReadOnly)]
public class FilterDimensions : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get active UIDocument and Document.
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Retrieve the current selection.
        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
        if (selectedIds == null || !selectedIds.Any())
        {
            TaskDialog.Show("Warning", "Please select dimension elements before running the command.");
            return Result.Failed;
        }

        // Filter the selection to include only Dimension elements.
        List<Dimension> selectedDimensions = selectedIds
            .Select(id => doc.GetElement(id))
            .OfType<Dimension>()
            .ToList();

        if (!selectedDimensions.Any())
        {
            TaskDialog.Show("Warning", "No dimension elements selected.");
            return Result.Failed;
        }

        // Determine maximum number of segments among selected dimensions.
        int maxSegmentColumns = selectedDimensions.Max(d => Math.Max(1, d.NumberOfSegments));

        // Define the columns for the data grid in the desired order.
        // New columns: Level, Owner View, Sheet Number, Sheet Name.
        List<string> propertyNames = new List<string>
        {
            "Element Id",
            "Name",
            "Level",
            "Owner View",
            "Sheet Number",
            "Sheet Name",
            "Begin X",
            "Begin Y",
            "Begin Z",
            "End X",
            "End Y",
            "End Z",
            "Total Length",
            "Segment Count"
        };

        // Add segment length columns immediately after "Segment Count".
        for (int i = 0; i < maxSegmentColumns; i++)
        {
            propertyNames.Add($"Length {i + 1}");
        }

        // Then add referenced element columns.
        propertyNames.Add("Referenced Element 1");
        propertyNames.Add("Referenced Element 2");
        propertyNames.Add("Referenced Element 3");
        propertyNames.Add("Referenced Element 4");

        List<Dictionary<string, object>> elementData = new List<Dictionary<string, object>>();

        // Collect all view sheets for later lookup.
        List<ViewSheet> allSheets = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet))
            .Cast<ViewSheet>()
            .ToList();

        foreach (Dimension dim in selectedDimensions)
        {
            Dictionary<string, object> props = new Dictionary<string, object>();

            // Basic properties.
            props["Element Id"] = dim.Id.IntegerValue;
            props["Name"] = dim.Name;

            // Retrieve Owner View, Level, Sheet Number, and Sheet Name.
            string ownerViewName = "N/A";
            string levelName = "N/A";
            string sheetNumber = "N/A";
            string sheetName = "N/A";
            if (dim.OwnerViewId != ElementId.InvalidElementId)
            {
                View ownerView = doc.GetElement(dim.OwnerViewId) as View;
                if (ownerView != null)
                {
                    ownerViewName = ownerView.Name;
                    // Try to get the level from the owner view.
                    if (ownerView.LevelId != ElementId.InvalidElementId)
                    {
                        Level lvl = doc.GetElement(ownerView.LevelId) as Level;
                        if (lvl != null)
                            levelName = lvl.Name;
                    }
                    else
                    {
                        Parameter assocLevelParam = ownerView.LookupParameter("Associated Level");
                        if (assocLevelParam != null && assocLevelParam.HasValue)
                        {
                            levelName = assocLevelParam.AsString() ?? "N/A";
                        }
                    }
                    // Look for a sheet placing this owner view.
                    foreach (ViewSheet sheet in allSheets)
                    {
                        if (sheet.GetAllPlacedViews().Contains(ownerView.Id))
                        {
                            sheetNumber = sheet.SheetNumber;
                            sheetName = sheet.Name;
                            break;
                        }
                    }
                }
            }
            props["Level"] = levelName;
            props["Owner View"] = ownerViewName;
            props["Sheet Number"] = sheetNumber;
            props["Sheet Name"] = sheetName;

            // Initialize coordinate and length properties.
            props["Begin X"] = "";
            props["Begin Y"] = "";
            props["Begin Z"] = "";
            props["End X"] = "";
            props["End Y"] = "";
            props["End Z"] = "";
            props["Total Length"] = "";
            props["Segment Count"] = "";

            // Initialize segment length columns with default "N/A".
            for (int i = 0; i < maxSegmentColumns; i++)
            {
                props[$"Length {i + 1}"] = "N/A";
            }

            // Initialize referenced element columns with "N/A".
            for (int i = 1; i <= 4; i++)
            {
                props[$"Referenced Element {i}"] = "N/A";
            }

            // Attempt to get endpoints from the dimension line.
            XYZ ptBegin = null;
            XYZ ptEnd = null;
            bool boundFound = false;
            Line dimLine = dim.Curve as Line;
            if (dimLine != null)
            {
                if (dimLine.IsBound)
                {
                    ptBegin = dimLine.GetEndPoint(0);
                    ptEnd = dimLine.GetEndPoint(1);
                    boundFound = true;
                }
                else
                {
                    // If the line is unbound, try to use reference points from the dimension.
                    List<XYZ> refPoints = new List<XYZ>();
                    ReferenceArray refArray = dim.References;
                    if (refArray != null)
                    {
                        foreach (Reference r in refArray)
                        {
                            XYZ pt = GetReferencePoint(doc, r);
                            if (pt != null)
                                refPoints.Add(pt);
                        }
                    }
                    if (refPoints.Count >= 2)
                    {
                        ptBegin = refPoints.First();
                        ptEnd = refPoints.Last();
                        boundFound = true;
                    }
                }
            }

            if (boundFound && ptBegin != null && ptEnd != null)
            {
                // Set begin/end coordinates rounded to three decimals.
                props["Begin X"] = ptBegin.X.ToString("F3");
                props["Begin Y"] = ptBegin.Y.ToString("F3");
                props["Begin Z"] = ptBegin.Z.ToString("F3");
                props["End X"] = ptEnd.X.ToString("F3");
                props["End Y"] = ptEnd.Y.ToString("F3");
                props["End Z"] = ptEnd.Z.ToString("F3");

                // Get total length (in feet) from built-in parameter or by computing distance.
                Parameter totalLengthParam = dim.get_Parameter(BuiltInParameter.DIM_TOTAL_LENGTH);
                double totalLengthFeet = (totalLengthParam != null && totalLengthParam.HasValue)
                    ? totalLengthParam.AsDouble()
                    : ptBegin.DistanceTo(ptEnd);
                double totalLengthMm = UnitUtils.ConvertFromInternalUnits(totalLengthFeet, UnitTypeId.Millimeters);
                props["Total Length"] = totalLengthMm.ToString("F2");
            }
            else
            {
                props["Begin X"] = "Unbound";
                props["Begin Y"] = "Unbound";
                props["Begin Z"] = "Unbound";
                props["End X"] = "Unbound";
                props["End Y"] = "Unbound";
                props["End Z"] = "Unbound";
                Parameter totalLengthParam = dim.get_Parameter(BuiltInParameter.DIM_TOTAL_LENGTH);
                props["Total Length"] = (totalLengthParam != null && totalLengthParam.HasValue)
                    ? totalLengthParam.AsValueString()
                    : "N/A";
            }

            // Segment information.
            int segCount = dim.NumberOfSegments;
            props["Segment Count"] = segCount > 0 ? segCount : 1;
            if (dim.NumberOfSegments > 0)
            {
                int idx = 0;
                foreach (DimensionSegment seg in dim.Segments)
                {
                    double? segValueFeet = seg.Value;
                    if (segValueFeet.HasValue)
                    {
                        double segLengthMm = UnitUtils.ConvertFromInternalUnits(segValueFeet.Value, UnitTypeId.Millimeters);
                        props[$"Length {idx + 1}"] = segLengthMm.ToString("F2");
                    }
                    else
                    {
                        props[$"Length {idx + 1}"] = "N/A";
                    }
                    idx++;
                }
                // For any missing segment columns, ensure "N/A" is assigned.
                for (int i = idx; i < maxSegmentColumns; i++)
                {
                    props[$"Length {i + 1}"] = "N/A";
                }
            }
            else
            {
                props["Length 1"] = props["Total Length"];
            }

            // Place referenced elements into separate columns (up to 4).
            int maxRefs = 4;
            ReferenceArray rArray = dim.References;
            for (int i = 0; i < maxRefs; i++)
            {
                string colName = $"Referenced Element {i + 1}";
                if (rArray != null && i < rArray.Size)
                {
                    Reference r = rArray.get_Item(i);
                    Element refElem = doc.GetElement(r);
                    props[colName] = (refElem != null) ? refElem.Name : "N/A";
                }
                else
                {
                    props[colName] = "N/A";
                }
            }

            elementData.Add(props);
        }

        List<Dictionary<string, object>> selectedFromGrid = CustomGUIs.DataGrid(elementData, propertyNames, false);
        if (selectedFromGrid?.Any() == true)
        {
            List<ElementId> selectedDimensionIds = selectedFromGrid
                .Select(dict => new ElementId((int)dict["Element Id"]))
                .ToList();
            uidoc.SetSelectionIds(selectedDimensionIds);
        }

        return Result.Succeeded;
    }

    /// <summary>
    /// Returns an approximate reference point for a given Reference.
    /// </summary>
    private XYZ GetReferencePoint(Document doc, Reference reference)
    {
        Element elem = doc.GetElement(reference);
        if (elem == null)
            return null;
        Location loc = elem.Location;
        if (loc is LocationPoint locPt)
            return locPt.Point;
        else if (loc is LocationCurve locCurve)
            return locCurve.Curve.Evaluate(0.5, true);
        else
        {
            BoundingBoxXYZ bb = elem.get_BoundingBox(null);
            if (bb != null)
                return (bb.Min + bb.Max) / 2;
        }
        return null;
    }
}
