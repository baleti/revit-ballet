using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

[Transaction(TransactionMode.Manual)]
public class DirectShapesFromVisibleRooms : IExternalCommand
{
    public Result Execute(ExternalCommandData data, ref string message, ElementSet elems)
    {
        UIDocument uidoc = data.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;
        View       view  = doc.ActiveView;

        // -------------------------------------------------- 1. collect visible rooms
        IList<Room> rooms = GetVisibleRooms(doc, view);
        if (rooms.Count == 0)
        {
            message = "No visible rooms in this view / section box.";
            return Result.Failed;
        }

        // -------------------------------------------------- 2. DataGrid picker
        var table = rooms.Select(r => new Dictionary<string, object>
        {
            ["Id"]     = r.Id.IntegerValue,
            ["Number"] = r.Number,
            ["Name"]   = StripNumber(r.Name, r.Number),                // display name without number
            ["Level"]  = (doc.GetElement(r.LevelId) as Level)?.Name ?? "",
            ["Area"]   = $"{UnitUtils.ConvertFromInternalUnits(r.Area, UnitTypeId.SquareMeters):F2} m²"
        }).ToList();

        var visibleCols = new List<string> { "Id", "Number", "Name", "Level", "Area" };
        var pick        = CustomGUIs.DataGrid(table, visibleCols, /*multiSelect*/ false);
        if (pick == null || pick.Count == 0) return Result.Cancelled;

        var pickedIds = pick.Select(d => (int)d["Id"]).ToHashSet();
        var chosen    = rooms.Where(r => pickedIds.Contains(r.Id.IntegerValue)).ToList();
        if (chosen.Count == 0)
        {
            message = "Selected room not found.";
            return Result.Failed;
        }

        // -------------------------------------------------- 3. create DirectShapes
        using (Transaction t = new Transaction(doc, "Create DirectShapes for Rooms"))
        {
            t.Start();
            int created = 0;
            foreach (Room r in chosen)
                if (TryCreateDirectShape(doc, r)) created++;

            if (created == 0)
            {
                message = "No DirectShapes could be created (invalid room boundaries?).";
                t.RollBack();
                return Result.Failed;
            }
            t.Commit();
        }
        return Result.Succeeded;
    }

    // --------------------------------------------------------------------------
    IList<Room> GetVisibleRooms(Document doc, View view)
    {
        var rooms = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Rooms)
            .OfClass(typeof(SpatialElement))
            .Cast<Room>()
            .Where(r => r.Area > 0)
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

            rooms = rooms.Where(r => r.get_BoundingBox(null) != null && filter.PassesFilter(r))
                         .ToList();
        }
        // 2-D cropped view
        else if (view.CropBoxActive)
        {
            rooms = new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .OfClass(typeof(SpatialElement))
                .Cast<Room>()
                .Where(r => r.Area > 0)
                .ToList();
        }
        return rooms;
    }

    // --------------------------------------------------------------------------
    bool TryCreateDirectShape(Document doc, Room room)
    {
        var loops = room
            .GetBoundarySegments(new SpatialElementBoundaryOptions())?
            .Select(bl => bl.Select(s => s.GetCurve()).Where(c => c != null).ToList())
            .Where(cs => cs.Count > 2)
            .ToList();
        if (loops == null || loops.Count == 0) return false;

        CurveLoop loop;
        try { loop = CurveLoop.Create(loops.First()); } catch { return false; }
        if (loop.IsOpen()) return false;

        double height = room.LookupParameter("Height")?.AsDouble()
            ?? UnitUtils.ConvertToInternalUnits(3.0, UnitTypeId.Meters);

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
        ds.Name = Sanitize(room.Name);

        // Comments: "<Number> - <stripped room name>"
        string stripped = StripNumber(room.Name, room.Number);
        if (string.IsNullOrWhiteSpace(stripped)) stripped = room.Name;
        string comment  = $"{room.Number} - {stripped}";

        Parameter p = ds.LookupParameter("Comments") ??
                      ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (p != null && p.StorageType == StorageType.String)
            p.Set(Sanitize(comment));

        return true;
    }

    // --------------------------------------------------------------------------
    static string StripNumber(string name, string number)
    {
        // Remove the exact room number at start or end (with optional separators)
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
