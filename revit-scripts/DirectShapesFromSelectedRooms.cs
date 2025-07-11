using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

[Transaction(TransactionMode.Manual)]
public class DirectShapesFromSelectedRooms : IExternalCommand
{
    public Result Execute(ExternalCommandData data, ref string message, ElementSet elems)
    {
        UIDocument uidoc = data.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;

        // -------------------------------------------------- 1. collect selected rooms
        IList<Room> rooms = GetSelectedRooms(uidoc);
        if (rooms.Count == 0)
        {
            message = "No rooms selected. Please select one or more rooms.";
            return Result.Failed;
        }

        // -------------------------------------------------- 2. create DirectShapes
        List<ElementId> createdIds = new List<ElementId>();
        
        using (Transaction t = new Transaction(doc, "Create DirectShapes for Selected Rooms"))
        {
            t.Start();
            foreach (Room r in rooms)
            {
                ElementId dsId = TryCreateDirectShape(doc, r);
                if (dsId != null && dsId != ElementId.InvalidElementId)
                    createdIds.Add(dsId);
            }

            if (createdIds.Count == 0)
            {
                message = "No DirectShapes could be created (invalid room boundaries?).";
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
    IList<Room> GetSelectedRooms(UIDocument uidoc)
    {
        // Get currently selected elements
        var selection = uidoc.GetSelectionIds();
        
        // Filter to get only rooms with valid area
        var rooms = selection
            .Select(id => uidoc.Document.GetElement(id))
            .OfType<Room>()
            .Where(r => r.Area > 0)
            .ToList();

        return rooms;
    }

    // --------------------------------------------------------------------------
    ElementId TryCreateDirectShape(Document doc, Room room)
    {
        var loops = room
            .GetBoundarySegments(new SpatialElementBoundaryOptions())?
            .Select(bl => bl.Select(s => s.GetCurve()).Where(c => c != null).ToList())
            .Where(cs => cs.Count > 2)
            .ToList();
        if (loops == null || loops.Count == 0) return null;

        CurveLoop loop;
        try { loop = CurveLoop.Create(loops.First()); } catch { return null; }
        if (loop.IsOpen()) return null;

        double height = room.LookupParameter("Height")?.AsDouble()
            ?? UnitUtils.ConvertToInternalUnits(3.0, UnitTypeId.Meters);

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
        ds.Name = Sanitize(room.Name);

        // Get the level name
        Level level = doc.GetElement(room.LevelId) as Level;
        string levelName = level?.Name ?? "Unknown Level";

        // Comments: "<Number> - <stripped room name>"
        string stripped = StripNumber(room.Name, room.Number);
        if (string.IsNullOrWhiteSpace(stripped)) stripped = room.Name;
        string comment = $"{room.Number} - {stripped} - {levelName}";

        Parameter p = ds.LookupParameter("Comments") ??
                      ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (p != null && p.StorageType == StorageType.String)
            p.Set(Sanitize(comment));

        return ds.Id;
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
