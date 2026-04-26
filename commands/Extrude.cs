#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using RevitAddin;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[CommandMeta("Area, Room, Floor")]
[CommandOutput("DirectShape")]
public class Extrude : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        var resolved = InputResolver.ResolveElements(uiDoc,
            BuiltInCategory.OST_Areas,
            BuiltInCategory.OST_Rooms,
            BuiltInCategory.OST_Floors);

        if (resolved == null)
        {
            // Nothing relevant selected — show DataGrid picker for all Areas, Rooms, Floors
            var candidates = new FilteredElementCollector(doc)
                .WherePasses(new LogicalOrFilter(new List<ElementFilter>
                {
                    new ElementCategoryFilter(BuiltInCategory.OST_Areas),
                    new ElementCategoryFilter(BuiltInCategory.OST_Rooms),
                    new ElementCategoryFilter(BuiltInCategory.OST_Floors),
                }))
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .Where(e => e.IsValidObject)
                .ToList();

            if (candidates.Count == 0)
            {
                message = "No Areas, Rooms, or Floors found in document.";
                return Result.Cancelled;
            }

            var pickerData = candidates.Select(e => new Dictionary<string, object>
            {
                { "Name", e.Name },
                { "Category", e.Category?.Name ?? "" },
                { "Level", (doc.GetElement(e.LevelId) as Level)?.Name ?? "" },
                { "ElementIdObject", e.Id }
            }).ToList();

            var picked = CustomGUIs.DataGrid(pickerData, new List<string> { "Name", "Category", "Level" }, false);
            if (picked == null || picked.Count == 0)
                return Result.Cancelled;

            resolved = new List<Element>();
            foreach (var row in picked)
            {
                if (row.TryGetValue("ElementIdObject", out object idObj) && idObj is ElementId eid)
                {
                    Element elem = doc.GetElement(eid);
                    if (elem != null) resolved.Add(elem);
                }
            }
        }

        if (resolved.Count == 0)
            return Result.Cancelled;

        using (Transaction trans = new Transaction(doc, "Extrude"))
        {
            trans.Start();
            foreach (Element elem in resolved)
            {
                if (elem is Area area)        CreateExtrusion(doc, area);
                else if (elem is Room room)   CreateExtrusion(doc, room);
                else if (elem is Floor floor) CreateFloorExtrusion(doc, floor);
            }
            trans.Commit();
        }

        return Result.Succeeded;
    }

    private void CreateExtrusion(Document doc, SpatialElement spatial)
    {
        SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
        IList<IList<BoundarySegment>> boundaries = spatial.GetBoundarySegments(options);
        if (boundaries == null || boundaries.Count == 0) return;

        List<Curve> curves = boundaries[0].Select(seg => seg.GetCurve()).ToList();
        if (curves.Count == 0) return;

        CurveLoop profileLoop = BuildCurveLoop(curves);
        if (profileLoop == null || profileLoop.IsOpen()) return;

        double height = GetHeight(spatial);
        Solid extrusion = GeometryCreationUtilities.CreateExtrusionGeometry(
            new List<CurveLoop> { profileLoop }, XYZ.BasisZ, height);

        DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        ds.SetShape(new GeometryObject[] { extrusion });
        ds.Name = $"{spatial.Name} Extrusion";
    }

    private void CreateFloorExtrusion(Document doc, Floor floor)
    {
        Options geomOptions = new Options { ComputeReferences = false };
        GeometryElement geomElem = floor.get_Geometry(geomOptions);
        if (geomElem == null) return;

        Solid floorSolid = geomElem.OfType<Solid>().OrderByDescending(s => s.Volume).FirstOrDefault();
        if (floorSolid == null || floorSolid.Volume <= 0) return;

        DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        ds.SetShape(new GeometryObject[] { floorSolid });
        ds.Name = $"{floor.Name} Extrusion";
    }

    private CurveLoop BuildCurveLoop(List<Curve> curves)
    {
        CurveLoop loop = new CurveLoop();
        List<Curve> remaining = new List<Curve>(curves);

        Curve first = remaining[0];
        loop.Append(first);
        XYZ currentEnd = first.GetEndPoint(1);
        remaining.RemoveAt(0);

        while (remaining.Count > 0)
        {
            bool found = false;
            for (int i = 0; i < remaining.Count; i++)
            {
                Curve c = remaining[i];
                if (c.GetEndPoint(0).IsAlmostEqualTo(currentEnd))
                {
                    loop.Append(c);
                    currentEnd = c.GetEndPoint(1);
                    remaining.RemoveAt(i);
                    found = true;
                    break;
                }
                if (c.GetEndPoint(1).IsAlmostEqualTo(currentEnd))
                {
                    Curve rev = c.CreateReversed();
                    loop.Append(rev);
                    currentEnd = rev.GetEndPoint(1);
                    remaining.RemoveAt(i);
                    found = true;
                    break;
                }
            }
            if (!found) break;
        }
        return loop.IsOpen() ? null : loop;
    }

    private double GetHeight(SpatialElement spatial)
    {
        Parameter p = spatial.LookupParameter("Height") ?? spatial.LookupParameter("Limit Offset");
        if (p != null && p.StorageType == StorageType.Double && p.AsDouble() > 0)
            return p.AsDouble();
        return UnitUtils.ConvertToInternalUnits(3.0, UnitTypeId.Meters);
    }
}

#endif
