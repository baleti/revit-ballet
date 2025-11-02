using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
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
            // Start a transaction
            using (Transaction t = new Transaction(doc, "Create Filled Region for Selected Rooms"))
            {
                t.Start();

                // Get the filled region type
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(FilledRegionType));
                FilledRegionType filledRegionType = collector.FirstElement() as FilledRegionType;

                if (filledRegionType == null)
                {
                    message = "No filled region type found in the project.";
                    return Result.Failed;
                }

                // Get all rooms
                FilteredElementCollector roomCollector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .OfClass(typeof(SpatialElement));
                List<Room> rooms = roomCollector.Cast<Room>().ToList();

                // Prepare data for DataGridView
                List<Dictionary<string, object>> roomEntries = new List<Dictionary<string, object>>();
                foreach (Room room in rooms)
                {
                    Level roomLevel = doc.GetElement(room.LevelId) as Level;
                    Dictionary<string, object> entry = new Dictionary<string, object>
                    {
                        { "Name", room.Name },
                        { "Level", roomLevel.Name }
                    };
                    roomEntries.Add(entry);
                }

                // Define properties to display
                List<string> properties = new List<string> { "Name", "Level" };

                // Get selected rooms
                var selectedEntries = CustomGUIs.DataGrid(roomEntries, properties, false);

                // Filter rooms based on selection
                var selectedRooms = rooms.Where(r => selectedEntries.Any(e =>
                    e["Name"].ToString() == r.Name &&
                    e["Level"].ToString() == ((Level)doc.GetElement(r.LevelId)).Name)).ToList();

                foreach (Room room in selectedRooms)
                {
                    // Get the room boundaries
                    IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
                    if (boundaries == null || boundaries.Count == 0)
                    {
                        continue;
                    }

                    // Create curves for filled region boundaries
                    List<CurveLoop> curveLoops = new List<CurveLoop>();
                    foreach (IList<BoundarySegment> boundary in boundaries)
                    {
                        CurveLoop curveLoop = new CurveLoop();
                        foreach (BoundarySegment segment in boundary)
                        {
                            curveLoop.Append(segment.GetCurve());
                        }
                        curveLoops.Add(curveLoop);
                    }

                    // Create the filled region
                    FilledRegion filledRegion = FilledRegion.Create(doc, filledRegionType.Id, doc.ActiveView.Id, curveLoops);
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
