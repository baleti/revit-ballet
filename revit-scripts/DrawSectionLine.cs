using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class DrawSectionLine : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View currentView = doc.ActiveView;

        // Abort if active view is not a plan view
        if (currentView.ViewType != ViewType.AreaPlan && currentView.ViewType != ViewType.FloorPlan && currentView.ViewType != ViewType.CeilingPlan)
        {
            TaskDialog.Show("Error", "The active view is not a plan view.");
            return Result.Failed;
        }

        // Get all section views in the document
        var sectionViews = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSection))
            .Cast<ViewSection>()
            .Select(v => new Dictionary<string, object>
            {
                { "Id", v.Id.IntegerValue },
                { "Name", v.Name },
                { "Type", v.ViewType.ToString() }
            })
            .ToList();

        if (!sectionViews.Any())
        {
            TaskDialog.Show("Error", "No section views found.");
            return Result.Failed;
        }

        List<string> propertyNames = new List<string> { "Id", "Name", "Type" };
        var selectedEntries = CustomGUIs.DataGrid(sectionViews, propertyNames, false);

        if (!selectedEntries.Any())
        {
            TaskDialog.Show("Cancelled", "No section view selected.");
            return Result.Cancelled;
        }

        List<ElementId> detailLineIds = new List<ElementId>();

        // Start a transaction to draw the section line
        using (Transaction trans = new Transaction(doc, "Draw Section Line"))
        {
            trans.Start();

            foreach (var entry in selectedEntries)
            {
                int selectedViewId = (int)entry["Id"];
                ElementId viewElementId = new ElementId(selectedViewId);
                ViewSection selectedView = doc.GetElement(viewElementId) as ViewSection;

                if (selectedView == null)
                {
                    TaskDialog.Show("Error", "Selected section view could not be found.");
                    continue;
                }

                // Get the section line
                Line sectionLine = GetSectionLine(selectedView);

                if (sectionLine != null)
                {
                    // Draw detail lines along the section line in the current view
                    DetailCurve detailCurve = doc.Create.NewDetailCurve(currentView, sectionLine);
                    detailLineIds.Add(detailCurve.Id);
                }
                else
                {
                    TaskDialog.Show("Error", "Selected section view does not have a valid section line.");
                    trans.RollBack();
                    return Result.Failed;
                }
            }

            trans.Commit();
        }

        // Set the detail lines as the current selection
        uidoc.SetSelectionIds(detailLineIds);

        return Result.Succeeded;
    }

    private Line GetSectionLine(ViewSection sectionView)
    {
        BoundingBoxXYZ boundingBox = sectionView.get_BoundingBox(null);
        if (boundingBox != null)
        {
            XYZ startPoint = boundingBox.Min;
            XYZ endPoint = boundingBox.Max;
            return Line.CreateBound(startPoint, endPoint);
        }
        return null;
    }
}
