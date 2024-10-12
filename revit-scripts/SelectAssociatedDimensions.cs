using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using System;
using System.Drawing;
using System.ComponentModel;

[Transaction(TransactionMode.Manual)]
public class SelectAssociatedDimensions : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the application and document
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the current selection
        Selection sel = uiDoc.Selection;
        IList<ElementId> selectedIds = sel.GetElementIds().ToList();

        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "Please select one or more elements.");
            return Result.Cancelled;
        }

        // Collect all dimensions in the document
        FilteredElementCollector dimCollector = new FilteredElementCollector(doc);
        IList<Dimension> allDimensions = dimCollector.OfClass(typeof(Dimension)).Cast<Dimension>().ToList();

        // List to store associated dimensions
        List<Dimension> associatedDimensions = new List<Dimension>();

        foreach (Dimension dim in allDimensions)
        {
            ReferenceArray refs = dim.References;
            if (refs != null)
            {
                foreach (Reference r in refs)
                {
                    ElementId refElemId = r.ElementId;
                    if (selectedIds.Contains(refElemId))
                    {
                        associatedDimensions.Add(dim);
                        break; // No need to check other references for this dimension
                    }
                }
            }
        }

        if (associatedDimensions.Count > 0)
        {
            // Prepare data for DataGrid
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            foreach (var dim in associatedDimensions)
            {
                Dictionary<string, object> entry = new Dictionary<string, object>();
                entry["Name"] = dim.Name;

                // Retrieve the sheet name where OwnerView is placed
                string sheetName = GetSheetNameForOwnerView(dim.OwnerViewId, doc);
                entry["Sheet Name"] = sheetName;

                // Get the owner view name
                entry["OwnerView"] = doc.GetElement(dim.OwnerViewId)?.Name ?? "N/A";
                entry["Id"] = dim.Id.IntegerValue;

                // Retrieve Total Length
                string totalLength = GetTotalLength(dim, doc);
                entry["Total Length"] = totalLength;

                // Set Count to the number of segments (or 1 if count is 0)
                entry["Count"] = dim.NumberOfSegments == 0 ? 1 : dim.NumberOfSegments;

                // Get segment lengths
                string segmentLengths = GetSegmentLengths(dim, doc);
                entry["Segment Lengths"] = segmentLengths;

                entries.Add(entry);
            }

            // Show DataGrid to user
            List<string> propertyNames = new List<string> { "Name", "Sheet Name", "OwnerView", "Id", "Total Length", "Count", "Segment Lengths" };
            List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, false);

            if (selectedEntries != null && selectedEntries.Count > 0)
            {
                // Get the selected dimension IDs
                List<ElementId> selectedDimensionIds = new List<ElementId>();
                foreach (var entry in selectedEntries)
                {
                    int idValue = Convert.ToInt32(entry["Id"]);
                    selectedDimensionIds.Add(new ElementId(idValue));
                }

                // Select the dimensions in Revit
                sel.SetElementIds(selectedDimensionIds);
            }
        }
        else
        {
            TaskDialog.Show("Result", "No dimensions are associated with the selected elements.");
        }

        return Result.Succeeded;
    }

    private string GetTotalLength(Dimension dim, Document doc)
    {
        Parameter totalLengthParam = dim.get_Parameter(BuiltInParameter.DIM_TOTAL_LENGTH);
        if (totalLengthParam != null && totalLengthParam.HasValue)
        {
            return totalLengthParam.AsValueString();
        }
        else
        {
            // Fallback to using Value property
            double? length = dim.Value;
            if (length.HasValue)
            {
                return FormatLength(length.Value, doc);
            }
            else
            {
                return "N/A";
            }
        }
    }

    private string GetSegmentLengths(Dimension dim, Document doc)
    {
        List<string> segmentLengthStrings = new List<string>();
        if (dim.NumberOfSegments > 1)
        {
            foreach (DimensionSegment seg in dim.Segments)
            {
                // Use Value property to get segment length
                double? segLength = seg.Value;
                if (segLength.HasValue)
                {
                    string formattedSegLength = FormatLength(segLength.Value, doc);
                    segmentLengthStrings.Add(formattedSegLength);
                }
                else
                {
                    segmentLengthStrings.Add("N/A");
                }
            }
        }
        else
        {
            // Single-segment dimension
            string totalLength = GetTotalLength(dim, doc);
            segmentLengthStrings.Add(totalLength);
        }
        return string.Join(" ", segmentLengthStrings);
    }

    private string GetSheetNameForOwnerView(ElementId ownerViewId, Document doc)
    {
        if (ownerViewId == ElementId.InvalidElementId)
        {
            return "N/A";
        }

        Autodesk.Revit.DB.View ownerView = doc.GetElement(ownerViewId) as Autodesk.Revit.DB.View;
        if (ownerView == null)
        {
            return "N/A";
        }

        // Use a collector to find the sheet where the view is placed
        FilteredElementCollector sheetCollector = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet));

        foreach (ViewSheet sheet in sheetCollector)
        {
            if (sheet.GetAllPlacedViews().Contains(ownerViewId))
            {
                return sheet.Name;  // Return the sheet name where the view is placed
            }
        }

        return "N/A";  // If not placed on a sheet
    }

    private string FormatLength(double lengthInFeet, Document doc)
    {
        return UnitFormatUtils.Format(doc.GetUnits(), SpecTypeId.Length, lengthInFeet, false);
    }
}
