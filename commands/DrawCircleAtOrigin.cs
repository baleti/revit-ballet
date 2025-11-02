using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class DrawCircleAtOrigin : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View view = doc.ActiveView;

        // Define circle center at point (0,0,0)
        XYZ center = new XYZ(0, 0, 0);

        // Define circle radius (100 mm = 0.1 meter)
        double radius = UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Millimeters);

        // Define the circle geometry
        Arc circle = Arc.Create(center, radius, 0, 2 * Math.PI, XYZ.BasisX, XYZ.BasisY);

        // Start a new transaction
        using (Transaction trans = new Transaction(doc, "Draw Circle"))
        {
            trans.Start();

            // Create a detail curve in the current view
            DetailCurve detailCircle = doc.Create.NewDetailCurve(view, circle);

            trans.Commit();

            // Set the detail curve as the active selection
            ICollection<ElementId> selectedIds = new List<ElementId> { detailCircle.Id };
            uidoc.SetSelectionIds(selectedIds);
        }

        return Result.Succeeded;
    }
}
