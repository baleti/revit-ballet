using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[CommandMeta("Room")]
public class TraceRooms : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData, 
      ref string message, 
      ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            // ── 1. Resolve rooms: selection-first, then DataGrid picker ──
            List<Room> selectedRooms = uidoc.GetSelectionIds()
                .Select(id => doc.GetElement(id))
                .OfType<Room>()
                .Where(r => r.Area > 0)
                .ToList();

            if (selectedRooms.Count == 0)
            {
                List<Room> allRooms = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .OfClass(typeof(SpatialElement))
                    .Cast<Room>().Where(r => r.Area > 0).ToList();

                List<Dictionary<string, object>> roomEntries = allRooms.Select(r => new Dictionary<string, object>
                {
                    { "Name",  r.Name },
                    { "Level", (doc.GetElement(r.LevelId) as Level)?.Name ?? "" },
                    { "__OriginalObject", r }
                }).ToList();

                var picked = CustomGUIs.DataGrid(roomEntries, new List<string> { "Name", "Level" }, false);
                if (picked == null || picked.Count == 0) return Result.Cancelled;

                selectedRooms = picked
                    .Select(d => d.ContainsKey("__OriginalObject") ? d["__OriginalObject"] as Room : null)
                    .Where(r => r != null)
                    .ToList();
                if (selectedRooms.Count == 0) return Result.Succeeded;
            }

            // ── 2. Get a filled region type ──
            FilledRegionType filledRegionType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType)).FirstElement() as FilledRegionType;
            if (filledRegionType == null)
            {
                message = "No filled region type found in the project.";
                return Result.Failed;
            }

            // ── 3. Create filled regions ──
            using (Transaction t = new Transaction(doc, "Create Filled Region for Selected Rooms"))
            {
                t.Start();

                foreach (Room room in selectedRooms)
                {
                    IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
                    if (boundaries == null || boundaries.Count == 0) continue;

                    List<CurveLoop> curveLoops = new List<CurveLoop>();
                    foreach (IList<BoundarySegment> boundary in boundaries)
                    {
                        CurveLoop curveLoop = new CurveLoop();
                        foreach (BoundarySegment segment in boundary)
                            curveLoop.Append(segment.GetCurve());
                        curveLoops.Add(curveLoop);
                    }

                    FilledRegion.Create(doc, filledRegionType.Id, doc.ActiveView.Id, curveLoops);
                }

                t.Commit();
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
