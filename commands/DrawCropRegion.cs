using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class DrawCropRegion : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View currentView = doc.ActiveView;

        // Get all views in the document
        var allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => v.ViewType == ViewType.Detail || v.ViewType == ViewType.FloorPlan || v.ViewType == ViewType.Section || v.ViewType == ViewType.Elevation)
            .Select(v => new Dictionary<string, object>
            {
                { "Id", v.Id.IntegerValue },
                { "Name", v.Name },
                { "Type", v.ViewType.ToString() }
            })
            .ToList();

        if (!allViews.Any())
        {
            TaskDialog.Show("Error", "No views found.");
            return Result.Failed;
        }

        List<string> propertyNames = new List<string> { "Id", "Name", "Type" };
        var selectedEntries = CustomGUIs.DataGrid(allViews, propertyNames, false);

        if (!selectedEntries.Any())
        {
            return Result.Cancelled;
        }

        List<ElementId> detailCurveIds = new List<ElementId>();

        // Start a transaction to draw the crop region rectangles
        using (Transaction trans = new Transaction(doc, "Draw Crop Region"))
        {
            trans.Start();

            foreach (var entry in selectedEntries)
            {
                int selectedViewId = (int)entry["Id"];
                ElementId viewElementId = new ElementId(selectedViewId);
                View selectedView = doc.GetElement(viewElementId) as View;

                if (selectedView == null)
                {
                    TaskDialog.Show("Error", "Selected view could not be found.");
                    continue;
                }

                // Check if the selected view has an active crop region
                if (selectedView.CropBox != null && selectedView.CropBoxActive)
                {
                    // Use the ViewCropRegionShapeManager to get the crop region shape
                    ViewCropRegionShapeManager cropManager = selectedView.GetCropRegionShapeManager();
                    IList<CurveLoop> cropRegionLoops = cropManager.GetCropShape();

                    if (cropRegionLoops != null && cropRegionLoops.Count > 0)
                    {
                        foreach (CurveLoop loop in cropRegionLoops)
                        {
                            foreach (Curve curve in loop)
                            {
                                // Draw detail lines along each curve in the crop region
                                DetailCurve detailCurve = doc.Create.NewDetailCurve(currentView, curve);
                                detailCurveIds.Add(detailCurve.Id);
                            }
                        }
                    }
                    else
                    {
                        TaskDialog.Show("Error", "Selected view does not have a valid crop region shape.");
                        trans.RollBack();
                        return Result.Failed;
                    }
                }
                else
                {
                    TaskDialog.Show("Error", "Selected view does not have an active crop region.");
                    trans.RollBack();
                    return Result.Failed;
                }
            }

            trans.Commit();
        }

        // Set the detail curves as the current selection
        uidoc.SetSelectionIds(detailCurveIds);

        return Result.Succeeded;
    }
}
