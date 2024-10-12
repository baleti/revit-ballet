using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class MapLinesToSheet : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View activeView = uidoc.ActiveView;

        // Find the viewport that contains this active view
        Viewport viewport = GetViewportForActiveView(doc, activeView);
        if (viewport == null)
        {
            message = "No viewport found for the active view.";
            return Result.Failed;
        }

        // Get the sheet that contains the viewport
        ViewSheet sheet = doc.GetElement(viewport.SheetId) as ViewSheet;
        if (sheet == null)
        {
            message = "No sheet found for the active view.";
            return Result.Failed;
        }

        // Use the existing selection
        ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
        if (selectedIds.Count == 0)
        {
            message = "No elements selected. Please select some lines first.";
            return Result.Failed;
        }

        // Begin transaction
        using (Transaction trans = new Transaction(doc, "Map Lines to Sheet Space"))
        {
            trans.Start();

            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element is CurveElement curveElement)
                {
                    Curve curve = curveElement.GeometryCurve;

                    // Get the combined transformation from view space to sheet space
                    Transform transform = GetViewToSheetTransform(viewport, activeView);

                    // Apply the transform to the curve
                    Curve transformedCurve = curve.CreateTransformed(transform);

                    // Project the transformed curve to the sheet plane
                    Curve projectedCurve = ProjectCurveToSheetPlane(transformedCurve);

                    // Create the new curve on the sheet
                    if (projectedCurve != null)
                    {
                        doc.Create.NewDetailCurve(sheet, projectedCurve);
                    }
                    else
                    {
                        message = "Failed to project curve to sheet plane.";
                        return Result.Failed;
                    }
                }
            }

            trans.Commit();
        }

        return Result.Succeeded;
    }

    private Viewport GetViewportForActiveView(Document doc, View activeView)
    {
        // Find the viewport that displays the active view
        FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Viewport));
        foreach (Viewport vp in collector.Cast<Viewport>())
        {
            if (vp.ViewId == activeView.Id)
            {
                return vp;
            }
        }

        return null; // No matching viewport found
    }

    private Curve ProjectCurveToSheetPlane(Curve curve)
    {
        // Define the XY plane of the sheet (Z = 0)
        Plane sheetPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero);

        // Project each endpoint of the curve onto the sheet plane
        XYZ start = ProjectPointToPlane(curve.GetEndPoint(0), sheetPlane);
        XYZ end = ProjectPointToPlane(curve.GetEndPoint(1), sheetPlane);

        // Return a new line or arc based on the projected points
        if (curve is Line)
        {
            return Line.CreateBound(start, end);
        }
        else if (curve is Arc arc)
        {
            // For arcs, we also need to project a mid-point to define the arc
            XYZ mid = ProjectPointToPlane(arc.Evaluate(0.5, true), sheetPlane);
            return Arc.Create(start, end, mid);
        }

        // Handle other curve types if needed
        return null; // If the curve type is not supported
    }

    private XYZ ProjectPointToPlane(XYZ point, Plane plane)
    {
        // Project a point onto a plane
        XYZ originToPt = point - plane.Origin;
        double distance = originToPt.DotProduct(plane.Normal);
        XYZ projection = point - distance * plane.Normal;

        return projection;
    }

    private Transform GetViewToSheetTransform(Viewport viewport, View view)
    {
        // Get the projection transform from the viewport
        Transform projectionTransform = viewport.GetProjectionToSheetTransform();

        // Get the rotation of the viewport (if any)
        double rotationAngle = -viewport.GetBoxCenter().AngleTo(XYZ.BasisX);
        Transform rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, rotationAngle, viewport.GetBoxCenter());

        // Combine the projection and rotation transforms
        Transform viewportTransform = rotationTransform.Multiply(projectionTransform);

        // Get the crop region's transformation
        Transform cropTransform = GetCropRegionTransform(view);

        // Combine the crop transform with the viewport transform
        Transform totalTransform = viewportTransform.Multiply(cropTransform);

        return totalTransform;
    }

    private Transform GetCropRegionTransform(View view)
    {
        // Get the crop box of the view
        BoundingBoxXYZ cropBox = view.CropBox;

        // Get the transformation of the crop box
        Transform cropTransform = cropBox.Transform;

        // Adjust for the crop box origin
        XYZ cropOrigin = cropBox.Min;

        // Create a translation transform to move the origin
        Transform translationTransform = Transform.CreateTranslation(-cropOrigin);

        // Combine the crop transform and translation
        Transform totalCropTransform = cropTransform.Multiply(translationTransform);

        return totalCropTransform;
    }
}
