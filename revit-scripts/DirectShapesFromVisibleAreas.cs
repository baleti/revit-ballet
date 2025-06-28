using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

[Transaction(TransactionMode.Manual)]
public class DirectShapesFromVisibleAreas : IExternalCommand
{
    public Result Execute(ExternalCommandData data, ref string message, ElementSet elems)
    {
        UIDocument uidoc = data.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;
        View       view  = doc.ActiveView;

        // -------------------------------------------------- 1. collect visible areas
        IList<Area> areas = GetVisibleAreas(doc, view);
        if (areas.Count == 0)
        {
            message = "No visible areas in this view / section box.";
            return Result.Failed;
        }

        // -------------------------------------------------- 2. DataGrid picker
        var table = areas.Select(a => new Dictionary<string, object>
        {
            ["Id"]     = a.Id.IntegerValue,
            ["Number"] = a.Number,
            ["Name"]   = StripNumber(a.Name, a.Number),                // display name without number
            ["Level"]  = (doc.GetElement(a.LevelId) as Level)?.Name ?? "",
            ["Area"]   = $"{UnitUtils.ConvertFromInternalUnits(a.Area, UnitTypeId.SquareMeters):F2} m²"
        }).ToList();

        var visibleCols = new List<string> { "Id", "Number", "Name", "Level", "Area" };
        var pick        = CustomGUIs.DataGrid(table, visibleCols, /*multiSelect*/ false);
        if (pick == null || pick.Count == 0) return Result.Cancelled;

        var pickedIds = pick.Select(d => (int)d["Id"]).ToHashSet();
        var chosen    = areas.Where(a => pickedIds.Contains(a.Id.IntegerValue)).ToList();
        if (chosen.Count == 0)
        {
            message = "Selected area not found.";
            return Result.Failed;
        }

        // -------------------------------------------------- 3. create DirectShapes
        using (Transaction t = new Transaction(doc, "Create DirectShapes for Areas"))
        {
            t.Start();
            int created = 0;
            foreach (Area a in chosen)
                if (TryCreateDirectShape(doc, a)) created++;

            if (created == 0)
            {
                message = "No DirectShapes could be created (invalid area boundaries?).";
                t.RollBack();
                return Result.Failed;
            }
            t.Commit();
        }
        return Result.Succeeded;
    }

    // --------------------------------------------------------------------------
    IList<Area> GetVisibleAreas(Document doc, View view)
    {
        var areas = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Areas)
            .OfClass(typeof(SpatialElement))
            .Cast<Area>()
            .Where(a => a.Area > 0)
            .ToList();

        // 3-D view with section box
        if (view is View3D v3 && v3.IsSectionBoxActive)
        {
            BoundingBoxXYZ sb = v3.GetSectionBox();
            Transform      tr = sb.Transform;

            XYZ[] c =
            {
                tr.OfPoint(sb.Min), tr.OfPoint(sb.Max),
                tr.OfPoint(new XYZ(sb.Min.X, sb.Max.Y, sb.Min.Z)),
                tr.OfPoint(new XYZ(sb.Max.X, sb.Min.Y, sb.Min.Z)),
                tr.OfPoint(new XYZ(sb.Min.X, sb.Min.Y, sb.Max.Z)),
                tr.OfPoint(new XYZ(sb.Max.X, sb.Max.Y, sb.Min.Z)),
                tr.OfPoint(new XYZ(sb.Min.X, sb.Max.Y, sb.Max.Z)),
                tr.OfPoint(new XYZ(sb.Max.X, sb.Min.Y, sb.Max.Z))
            };
            XYZ wMin = new XYZ(c.Min(p => p.X), c.Min(p => p.Y), c.Min(p => p.Z));
            XYZ wMax = new XYZ(c.Max(p => p.X), c.Max(p => p.Y), c.Max(p => p.Z));

            var outline = new Outline(wMin, wMax);
            var filter  = new BoundingBoxIntersectsFilter(outline);

            areas = areas.Where(a => a.get_BoundingBox(null) != null && filter.PassesFilter(a))
                         .ToList();
        }
        // 2-D cropped view
        else if (view.CropBoxActive)
        {
            areas = new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_Areas)
                .OfClass(typeof(SpatialElement))
                .Cast<Area>()
                .Where(a => a.Area > 0)
                .ToList();
        }
        return areas;
    }

    // --------------------------------------------------------------------------
    bool TryCreateDirectShape(Document doc, Area area)
    {
        var loops = area
            .GetBoundarySegments(new SpatialElementBoundaryOptions())?
            .Select(bl => bl.Select(s => s.GetCurve()).Where(c => c != null).ToList())
            .Where(cs => cs.Count > 2)
            .ToList();
        if (loops == null || loops.Count == 0) return false;

        CurveLoop loop;
        try { loop = CurveLoop.Create(loops.First()); } catch { return false; }
        if (loop.IsOpen()) return false;

        // Areas typically don't have a Height parameter, so use a default
        // You can adjust this value or add a parameter lookup if needed
        double height = UnitUtils.ConvertToInternalUnits(3.0, UnitTypeId.Meters);

        Solid solid;
        try
        {
            solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                        new[] { loop }, XYZ.BasisZ, height);
        }
        catch { return false; }

        DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        if (ds == null) return false;

        try { ds.SetShape(new GeometryObject[] { solid }); }
        catch { doc.Delete(ds.Id); return false; }

        // Name retains Revit's original (includes number)
        ds.Name = Sanitize(area.Name);

        // Comments: "<Number> - <stripped area name>"
        string stripped = StripNumber(area.Name, area.Number);
        if (string.IsNullOrWhiteSpace(stripped)) stripped = area.Name;
        string comment  = $"{area.Number} - {stripped}";

        Parameter p = ds.LookupParameter("Comments") ??
                      ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (p != null && p.StorageType == StorageType.String)
            p.Set(Sanitize(comment));

        return true;
    }

    // --------------------------------------------------------------------------
    static string StripNumber(string name, string number)
    {
        // Remove the exact area number at start or end (with optional separators)
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(number))
            return name ?? "";

        return Regex.Replace(
            name,
            $@"^\s*{Regex.Escape(number)}[\s\-\.:]*|\s*[\-\.:]*{Regex.Escape(number)}\s*$",
            "",
            RegexOptions.IgnoreCase).Trim();
    }

    static string Sanitize(string raw)
    {
        var illegal = new Regex(@"[<>:{}|;?*\\/\[\]]");
        string clean = illegal.Replace(raw ?? "", "_").Trim();
        return clean.Length > 250 ? clean.Substring(0, 250) : clean;
    }
}
