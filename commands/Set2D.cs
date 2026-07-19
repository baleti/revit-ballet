using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
[CommandMeta("Grid, Level")]
public class Set2D : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;

        List<DatumPlane> selected = new List<DatumPlane>();
        foreach (ElementId id in uiDoc.GetSelectionIds())
        {
            Element elem = doc.GetElement(id);
            if (elem is Grid || elem is Level)
                selected.Add((DatumPlane)elem);
        }

        if (selected.Count == 0)
        {
            CustomGUIs.SetCurrentUIDocument(uiDoc);
            var items = new List<Dictionary<string, object>>();

            foreach (Grid g in new FilteredElementCollector(doc).OfClass(typeof(Grid)).Cast<Grid>())
                items.Add(new Dictionary<string, object> { { "Name", g.Name }, { "Type", "Grid" }, { "ElementIdObject", g.Id } });
            foreach (Level l in new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>())
                items.Add(new Dictionary<string, object> { { "Name", l.Name }, { "Type", "Level" }, { "ElementIdObject", l.Id } });

            var chosen = CustomGUIs.DataGrid(items, new List<string> { "Name", "Type" }, false);
            if (chosen == null || chosen.Count == 0) return Result.Cancelled;

            foreach (var row in chosen)
            {
                if (row.TryGetValue("ElementIdObject", out object idObj) && idObj is ElementId eid)
                {
                    if (doc.GetElement(eid) is DatumPlane dp)
                        selected.Add(dp);
                }
            }
            if (selected.Count == 0) return Result.Succeeded;
        }

        using (Transaction t = new Transaction(doc, "Set 2D Extents"))
        {
            t.Start();
            foreach (DatumPlane dp in selected)
            {
                dp.SetDatumExtentType(DatumEnds.End0, doc.ActiveView, DatumExtentType.ViewSpecific);
                dp.SetDatumExtentType(DatumEnds.End1, doc.ActiveView, DatumExtentType.ViewSpecific);
            }
            t.Commit();
        }

        return Result.Succeeded;
    }
}
