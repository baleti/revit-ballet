using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ExtrudeAreas : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
        if (selectedIds.Count == 0)
        {
            message = "Please select one or more areas.";
            return Result.Cancelled;
        }

        List<Area> selectedAreas = new List<Area>();
        foreach (ElementId id in selectedIds)
        {
            if (doc.GetElement(id) is Area area)
                selectedAreas.Add(area);
        }

        if (selectedAreas.Count == 0)
        {
            message = "No valid areas selected.";
            return Result.Cancelled;
        }

        using (Transaction trans = new Transaction(doc, "Create Area Extrusions"))
        {
            trans.Start();
            foreach (Area area in selectedAreas)
            {
                CreateExtrusionForArea(doc, area);
            }
            trans.Commit();
        }

        return Result.Succeeded;
    }

    private void CreateExtrusionForArea(Document doc, Area area)
    {
        SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
        IList<IList<BoundarySegment>> boundaries = area.GetBoundarySegments(options);
        if (boundaries == null || boundaries.Count == 0) return;

        List<Curve> curves = boundaries[0].Select(seg => seg.GetCurve()).ToList();
        if (curves.Count == 0) return;

        CurveLoop profileLoop = CreateContinuousCurveLoop(curves);
        if (profileLoop == null || profileLoop.IsOpen()) return;

        double height = GetExtrusionHeight(area);
        Solid extrusion = GeometryCreationUtilities.CreateExtrusionGeometry(
            new List<CurveLoop> { profileLoop },
            XYZ.BasisZ,
            height
        );

        DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        ds.SetShape(new GeometryObject[] { extrusion });
        ds.Name = $"{area.Name} Extrusion";
    }

    private CurveLoop CreateContinuousCurveLoop(List<Curve> curves)
    {
        CurveLoop loop = new CurveLoop();
        if (curves.Count == 0) return null;

        List<Curve> remainingCurves = new List<Curve>(curves);
        XYZ currentEnd = null;

        Curve firstCurve = remainingCurves[0];
        loop.Append(firstCurve);
        currentEnd = firstCurve.GetEndPoint(1);
        remainingCurves.RemoveAt(0);

        while (remainingCurves.Count > 0)
        {
            bool found = false;
            
            for (int i = 0; i < remainingCurves.Count; i++)
            {
                Curve curve = remainingCurves[i];
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                if (start.IsAlmostEqualTo(currentEnd))
                {
                    loop.Append(curve);
                    currentEnd = end;
                    remainingCurves.RemoveAt(i);
                    found = true;
                    break;
                }
                else if (end.IsAlmostEqualTo(currentEnd))
                {
                    Curve reversed = curve.CreateReversed();
                    loop.Append(reversed);
                    currentEnd = reversed.GetEndPoint(1);
                    remainingCurves.RemoveAt(i);
                    found = true;
                    break;
                }
            }
            if (!found) break;
        }
        return loop.IsOpen() ? null : loop;
    }

    private double GetExtrusionHeight(Area area)
    {
        Parameter heightParam = area.LookupParameter("Height");
        if (heightParam != null && heightParam.StorageType == StorageType.Double)
        {
            return heightParam.AsDouble();
        }
        return UnitUtils.ConvertToInternalUnits(3.0, UnitTypeId.Meters);
    }
}
