#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using RevitDirectShape = Autodesk.Revit.DB.DirectShape;

namespace RevitCommands
{

[Transaction(TransactionMode.Manual)]
[CommandMeta("Area, Room")]
[CommandOutput("DirectShape")]
public class DirectShape : IExternalCommand
{
    public Result Execute(ExternalCommandData data, ref string message, ElementSet elems)
    {
        UIDocument uidoc = data.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;

        List<Area> areas   = new List<Area>();
        List<Room> rooms   = new List<Room>();

        foreach (ElementId id in uidoc.GetSelectionIds())
        {
            Element elem = doc.GetElement(id);
            if (elem is Area a && a.Area > 0) areas.Add(a);
            else if (elem is Room r && r.Area > 0) rooms.Add(r);
        }

        if (areas.Count == 0 && rooms.Count == 0)
        {
            CustomGUIs.SetCurrentUIDocument(uidoc);
            var items = new List<Dictionary<string, object>>();

            foreach (Area a in new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Areas).OfClass(typeof(SpatialElement))
                .Cast<Area>().Where(a => a.Area > 0))
                items.Add(new Dictionary<string, object> { { "Name", a.Name }, { "Type", "Area" }, { "ElementIdObject", a.Id } });

            foreach (Room r in new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms).OfClass(typeof(SpatialElement))
                .Cast<Room>().Where(r => r.Area > 0))
                items.Add(new Dictionary<string, object> { { "Name", r.Name }, { "Type", "Room" }, { "ElementIdObject", r.Id } });

            if (items.Count == 0) return Result.Succeeded;

            var chosen = CustomGUIs.DataGrid(items, new List<string> { "Name", "Type" }, false);
            if (chosen == null || chosen.Count == 0) return Result.Cancelled;

            foreach (var row in chosen)
            {
                if (!row.TryGetValue("ElementIdObject", out object idObj) || !(idObj is ElementId eid)) continue;
                Element elem = doc.GetElement(eid);
                if (elem is Area a2 && a2.Area > 0) areas.Add(a2);
                else if (elem is Room r2 && r2.Area > 0) rooms.Add(r2);
            }
            if (areas.Count == 0 && rooms.Count == 0) return Result.Succeeded;
        }

        List<ElementId> createdIds = new List<ElementId>();

        using (Transaction t = new Transaction(doc, "Create DirectShapes"))
        {
            t.Start();
            foreach (Area a in areas)
            {
                ElementId id = TryCreate(doc, a, null);
                if (id != null && id != ElementId.InvalidElementId) createdIds.Add(id);
            }
            foreach (Room r in rooms)
            {
                double? height = r.LookupParameter("Height")?.AsDouble();
                ElementId id = TryCreate(doc, r, height);
                if (id != null && id != ElementId.InvalidElementId) createdIds.Add(id);
            }

            if (createdIds.Count == 0)
            {
                message = "No DirectShapes could be created (invalid boundaries?).";
                t.RollBack();
                return Result.Failed;
            }
            t.Commit();
        }

        var currentSelection = uidoc.GetSelectionIds().ToList();
        currentSelection.AddRange(createdIds);
        uidoc.SetSelectionIds(currentSelection);

        return Result.Succeeded;
    }

    private ElementId TryCreate(Document doc, SpatialElement spatial, double? heightOverride)
    {
        var loops = spatial
            .GetBoundarySegments(new SpatialElementBoundaryOptions())?
            .Select(bl => bl.Select(s => s.GetCurve()).Where(c => c != null).ToList())
            .Where(cs => cs.Count > 2)
            .ToList();
        if (loops == null || loops.Count == 0) return null;

        CurveLoop loop;
        try { loop = CurveLoop.Create(loops.First()); } catch { return null; }
        if (loop.IsOpen()) return null;

        double height = heightOverride ?? UnitUtils.ConvertToInternalUnits(3.0, UnitTypeId.Meters);

        Solid solid;
        try { solid = GeometryCreationUtilities.CreateExtrusionGeometry(new[] { loop }, XYZ.BasisZ, height); }
        catch { return null; }

        RevitDirectShape ds = RevitDirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        if (ds == null) return null;

        try { ds.SetShape(new GeometryObject[] { solid }); }
        catch { doc.Delete(ds.Id); return null; }

        ds.Name = Sanitize(spatial.Name);

        Level level = doc.GetElement(spatial.LevelId) as Level;
        string levelName = level?.Name ?? "Unknown Level";
        string number = (spatial as Area)?.Number ?? (spatial as Room)?.Number ?? "";
        string stripped = StripNumber(spatial.Name, number);
        if (string.IsNullOrWhiteSpace(stripped)) stripped = spatial.Name;

        Parameter p = ds.LookupParameter("Comments") ??
                      ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (p != null && p.StorageType == StorageType.String)
            p.Set(Sanitize($"{number} - {stripped} - {levelName}"));

        return ds.Id;
    }

    private static string StripNumber(string name, string number)
    {
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(number)) return name ?? "";
        return Regex.Replace(name,
            $@"^\s*{Regex.Escape(number)}[\s\-\.:]*|\s*[\-\.:]*{Regex.Escape(number)}\s*$",
            "", RegexOptions.IgnoreCase).Trim();
    }

    private static string Sanitize(string raw)
    {
        string clean = new Regex(@"[<>:{}|;?*\\/\[\]]").Replace(raw ?? "", "_").Trim();
        return clean.Length > 250 ? clean.Substring(0, 250) : clean;
    }
}

} // namespace RevitCommands
#endif
