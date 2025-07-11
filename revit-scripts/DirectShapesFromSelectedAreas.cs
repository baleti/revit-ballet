using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

[Transaction(TransactionMode.Manual)]
public class DirectShapesFromSelectedAreas : IExternalCommand
{
    public Result Execute(ExternalCommandData data, ref string message, ElementSet elems)
    {
        UIDocument uidoc = data.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;

        // -------------------------------------------------- 1. collect selected areas
        IList<Area> areas = GetSelectedAreas(uidoc);
        if (areas.Count == 0)
        {
            message = "No areas selected. Please select one or more areas.";
            return Result.Failed;
        }

        // -------------------------------------------------- 2. create DirectShapes
        List<ElementId> createdIds = new List<ElementId>();
        
        using (Transaction t = new Transaction(doc, "Create DirectShapes for Selected Areas"))
        {
            t.Start();
            foreach (Area a in areas)
            {
                ElementId dsId = TryCreateDirectShape(doc, a);
                if (dsId != null && dsId != ElementId.InvalidElementId)
                    createdIds.Add(dsId);
            }

            if (createdIds.Count == 0)
            {
                message = "No DirectShapes could be created (invalid area boundaries?).";
                t.RollBack();
                return Result.Failed;
            }
            t.Commit();
        }
        
        // Add newly created DirectShapes to current selection
        var currentSelection = uidoc.GetSelectionIds().ToList();
        currentSelection.AddRange(createdIds);
        uidoc.SetSelectionIds(currentSelection);
        
        return Result.Succeeded;
    }

    // --------------------------------------------------------------------------
    IList<Area> GetSelectedAreas(UIDocument uidoc)
    {
        // Get currently selected elements
        var selection = uidoc.GetSelectionIds();
        
        // Filter to get only areas with valid area
        var areas = selection
            .Select(id => uidoc.Document.GetElement(id))
            .OfType<Area>()
            .Where(a => a.Area > 0)
            .ToList();

        return areas;
    }

    // --------------------------------------------------------------------------
    ElementId TryCreateDirectShape(Document doc, Area area)
    {
        var loops = area
            .GetBoundarySegments(new SpatialElementBoundaryOptions())?
            .Select(bl => bl.Select(s => s.GetCurve()).Where(c => c != null).ToList())
            .Where(cs => cs.Count > 2)
            .ToList();
        if (loops == null || loops.Count == 0) return null;

        CurveLoop loop;
        try { loop = CurveLoop.Create(loops.First()); } catch { return null; }
        if (loop.IsOpen()) return null;

        // Areas typically don't have a Height parameter, so use a default
        // You can adjust this value or add a parameter lookup if needed
        double height = UnitUtils.ConvertToInternalUnits(3.0, UnitTypeId.Meters);

        Solid solid;
        try
        {
            solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                        new[] { loop }, XYZ.BasisZ, height);
        }
        catch { return null; }

        DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        if (ds == null) return null;

        try { ds.SetShape(new GeometryObject[] { solid }); }
        catch { doc.Delete(ds.Id); return null; }

        // Name retains Revit's original (includes number)
        ds.Name = Sanitize(area.Name);

        // Get the level name
        Level level = doc.GetElement(area.LevelId) as Level;
        string levelName = level?.Name ?? "Unknown Level";

        // Comments: "<Number> - <stripped area name>"
        string stripped = StripNumber(area.Name, area.Number);
        if (string.IsNullOrWhiteSpace(stripped)) stripped = area.Name;
        string comment = $"{area.Number} - {stripped} - {levelName}";

        Parameter p = ds.LookupParameter("Comments") ??
                      ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (p != null && p.StorageType == StorageType.String)
            p.Set(Sanitize(comment));

        return ds.Id;
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
