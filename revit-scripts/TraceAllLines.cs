// DON'T USE IT
// EXPLODE IMPORTED CAD AND COPY ELEMENTS INTO A NEW FAMILY
// CLEANING IT UP ISN'T DANGEROUS AND IT'S MUCH FASTER THAN THIS SCRIPT DUE TO REVIT API LIMITATIONS
// https://forums.autodesk.com/t5/revit-api-forum/creating-lines-increasingly-slower-from-3ms-to-200ms/td-p/12895581
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class TraceAllLines : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Prompt the user to select an imported CAD element
        Reference selectedReference = uidoc.Selection.PickObject(ObjectType.Element, "Select an imported CAD element");
        if (selectedReference == null)
        {
            return Result.Cancelled;
        }

        Element selectedElement = doc.GetElement(selectedReference);
        if (!(selectedElement is ImportInstance))
        {
            message = "Selected element is not an imported CAD element.";
            return Result.Failed;
        }

        ImportInstance importInstance = selectedElement as ImportInstance;
        GeometryElement geoElement = importInstance.get_Geometry(new Options());

        using (Transaction tx = new Transaction(doc))
        {
            tx.Start("Trace CAD Lines");

            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is GeometryInstance)
                {
                    GeometryInstance geoInstance = geoObject as GeometryInstance;
                    GeometryElement geoInstanceElement = geoInstance.GetInstanceGeometry();
                    foreach (GeometryObject geoInstanceObject in geoInstanceElement)
                    {
                        ProcessGeometryObject(doc, geoInstanceObject);
                    }
                }
            }

            tx.Commit();
        }

        return Result.Succeeded;
    }

    private void ProcessGeometryObject(Document doc, GeometryObject geoObject)
    {
        if (geoObject is Curve curve)
        {
            CreateDetailCurve(doc, curve.Clone());
        }
        else if (geoObject is PolyLine polyLine)
        {
            IList<XYZ> points = polyLine.GetCoordinates();
            for (int i = 0; i < points.Count - 1; i++)
            {
                XYZ start = points[i];
                XYZ end = points[i + 1];
                if (start.DistanceTo(end) > doc.Application.ShortCurveTolerance)
                {
                    Line line = Line.CreateBound(start, end);
                    CreateDetailCurve(doc, line);
                }
            }
        }
        else if (geoObject is Arc arc)
        {
            CreateDetailCurve(doc, arc.Clone());
        }
        else if (geoObject is Line line)
        {
            CreateDetailCurve(doc, line.Clone());
        }
        else if (geoObject is Ellipse ellipse)
        {
            CreateDetailCurve(doc, ellipse.Clone());
        }
        else if (geoObject is HermiteSpline spline)
        {
            CreateDetailCurve(doc, spline.Clone());
        }
    }

    private void CreateDetailCurve(Document doc, Curve curve)
    {
        if (doc.IsFamilyDocument)
        {
            doc.FamilyCreate.NewDetailCurve(doc.ActiveView, curve);
        }
        else
        {
            doc.Create.NewDetailCurve(doc.ActiveView, curve);
        }
    }
}
