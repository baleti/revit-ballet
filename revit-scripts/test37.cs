using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;

namespace RoomGraphPlugin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DrawRoomGraph : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Stopwatch totalTimer = Stopwatch.StartNew();
                Dictionary<string, long> timings = new Dictionary<string, long>();
                Stopwatch stepTimer = new Stopwatch();
                
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;
                
                // Verify we have a 3D view
                stepTimer.Restart();
                View3D view3D = doc.ActiveView as View3D;
                if (view3D == null)
                {
                    TaskDialog.Show("Error", "Please run this command in a 3D view.");
                    return Result.Failed;
                }
                timings["View Check"] = stepTimer.ElapsedMilliseconds;

                // Get all rooms in the model
                stepTimer.Restart();
                FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
                List<Room> rooms = roomCollector
                    .OfClass(typeof(SpatialElement))
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .Cast<Room>()
                    .Where(r => r.Area > 0)
                    .ToList();
                timings["Room Collection"] = stepTimer.ElapsedMilliseconds;

                if (rooms.Count == 0)
                {
                    TaskDialog.Show("Info", "No rooms found in the model.");
                    return Result.Succeeded;
                }

                // Get all doors
                stepTimer.Restart();
                FilteredElementCollector doorCollector = new FilteredElementCollector(doc);
                List<FamilyInstance> doors = doorCollector
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_Doors)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .ToList();
                timings["Door Collection"] = stepTimer.ElapsedMilliseconds;

                // Get all stairs
                stepTimer.Restart();
                FilteredElementCollector stairCollector = new FilteredElementCollector(doc);
                List<Element> stairs = stairCollector
                    .OfCategory(BuiltInCategory.OST_Stairs)
                    .WhereElementIsNotElementType()
                    .ToList();
                timings["Stair Collection"] = stepTimer.ElapsedMilliseconds;

                // Build room connectivity graph
                stepTimer.Restart();
                var graphResult = BuildRoomConnectivityGraph(doc, rooms, doors, stairs, timings);
                Dictionary<ElementId, List<ElementId>> roomConnections = graphResult.Item1;
                Dictionary<string, bool> verticalConnections = graphResult.Item2;
                Dictionary<string, string> debugInfo = graphResult.Item3;
                timings["Build Graph Total"] = stepTimer.ElapsedMilliseconds;

                // Create model lines in transaction
                stepTimer.Restart();
                int horizontalLinesCreated = 0;
                int verticalLinesCreated = 0;
                int failedLines = 0;
                Dictionary<string, int> failureReasons = new Dictionary<string, int>();
                List<ElementId> createdLineIds = new List<ElementId>();
                
                using (Transaction trans = new Transaction(doc, "Create Room Connection Graph"))
                {
                    trans.Start();
                    
                    // Process connections
                    stepTimer.Restart();
                    HashSet<string> processedConnections = new HashSet<string>();
                    
                    foreach (var roomId in roomConnections.Keys)
                    {
                        Room room1 = doc.GetElement(roomId) as Room;
                        if (room1 == null) continue;
                        
                        XYZ center1 = GetRoomCenter(room1);
                        if (center1 == null) continue;
                        
                        foreach (var connectedRoomId in roomConnections[roomId])
                        {
                            // Avoid duplicate connections
                            string connectionKey = GetConnectionKey(roomId, connectedRoomId);
                            if (processedConnections.Contains(connectionKey)) continue;
                            processedConnections.Add(connectionKey);
                            
                            Room room2 = doc.GetElement(connectedRoomId) as Room;
                            if (room2 == null) continue;
                            
                            XYZ center2 = GetRoomCenter(room2);
                            if (center2 == null) continue;
                            
                            bool isVertical = verticalConnections.ContainsKey(connectionKey) && verticalConnections[connectionKey];
                            
                            try
                            {
                                // Create a direct line between room centers
                                Line line = Line.CreateBound(center1, center2);
                                
                                // Create appropriate sketch plane
                                SketchPlane lineSketchPlane = null;
                                
                                if (Math.Abs(center1.Z - center2.Z) < 0.01)
                                {
                                    // Horizontal line - use level plane
                                    Plane horizontalPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, center1);
                                    lineSketchPlane = SketchPlane.Create(doc, horizontalPlane);
                                }
                                else
                                {
                                    // Non-horizontal line - create plane containing the line
                                    XYZ lineDir = (center2 - center1).Normalize();
                                    XYZ perpDir = null;
                                    
                                    // Find a perpendicular direction
                                    if (Math.Abs(lineDir.DotProduct(XYZ.BasisZ)) < 0.9)
                                    {
                                        perpDir = lineDir.CrossProduct(XYZ.BasisZ).Normalize();
                                    }
                                    else
                                    {
                                        perpDir = lineDir.CrossProduct(XYZ.BasisX).Normalize();
                                    }
                                    
                                    Plane linePlane = Plane.CreateByNormalAndOrigin(perpDir, center1);
                                    lineSketchPlane = SketchPlane.Create(doc, linePlane);
                                }
                                
                                ModelLine modelLine = doc.Create.NewModelCurve(line, lineSketchPlane) as ModelLine;
                                
                                if (modelLine != null)
                                {
                                    createdLineIds.Add(modelLine.Id);
                                    
                                    // Set graphics overrides
                                    OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                                    ogs.SetProjectionLineWeight(isVertical ? 6 : 3);
                                    if (isVertical)
                                    {
                                        ogs.SetProjectionLineColor(new Color(255, 0, 128));
                                        verticalLinesCreated++;
                                    }
                                    else
                                    {
                                        ogs.SetProjectionLineColor(new Color(0, 128, 255));
                                        horizontalLinesCreated++;
                                    }
                                    view3D.SetElementOverrides(modelLine.Id, ogs);
                                }
                                else
                                {
                                    failedLines++;
                                    string reason = "ModelLine creation returned null";
                                    if (!failureReasons.ContainsKey(reason))
                                        failureReasons[reason] = 0;
                                    failureReasons[reason]++;
                                }
                            }
                            catch (Exception ex)
                            {
                                failedLines++;
                                string reason = ex.Message;
                                if (!failureReasons.ContainsKey(reason))
                                    failureReasons[reason] = 0;
                                failureReasons[reason]++;
                            }
                        }
                    }
                    timings["Line Creation"] = stepTimer.ElapsedMilliseconds;
                    
                    trans.Commit();
                }
                timings["Transaction Commit"] = stepTimer.ElapsedMilliseconds;
                
                totalTimer.Stop();

                // Build summary
                StringBuilder summary = new StringBuilder();
                summary.AppendLine($"Room Graph Generation Complete");
                summary.AppendLine($"==============================");
                summary.AppendLine($"Total Execution Time: {totalTimer.ElapsedMilliseconds}ms");
                summary.AppendLine();
                
                summary.AppendLine($"Performance Breakdown:");
                foreach (var timing in timings.OrderByDescending(t => t.Value))
                {
                    double percentage = (timing.Value * 100.0) / totalTimer.ElapsedMilliseconds;
                    summary.AppendLine($"- {timing.Key}: {timing.Value}ms ({percentage:F1}%)");
                }
                summary.AppendLine();
                
                summary.AppendLine($"Statistics:");
                summary.AppendLine($"- Rooms: {rooms.Count}");
                summary.AppendLine($"- Doors: {doors.Count}");
                summary.AppendLine($"- Stairs: {stairs.Count}");
                summary.AppendLine($"- Lines Created: {createdLineIds.Count}");
                summary.AppendLine($"  - Horizontal (blue): {horizontalLinesCreated}");
                summary.AppendLine($"  - Vertical (magenta): {verticalLinesCreated}");
                if (failedLines > 0)
                {
                    summary.AppendLine($"  - Failed: {failedLines}");
                    summary.AppendLine();
                    summary.AppendLine($"Failure Reasons:");
                    foreach (var reason in failureReasons.OrderByDescending(r => r.Value))
                    {
                        summary.AppendLine($"  - {reason.Key}: {reason.Value}");
                    }
                }
                
                if (debugInfo.Count > 0)
                {
                    summary.AppendLine();
                    summary.AppendLine($"Debug Information:");
                    foreach (var info in debugInfo.Take(20))
                    {
                        summary.AppendLine($"- {info.Key}: {info.Value}");
                    }
                }
                
                TaskDialog.Show("Room Graph Complete", summary.ToString());

                // Select created lines for easier inspection
                if (createdLineIds.Count > 0)
                {
                    uiDoc.Selection.SetElementIds(createdLineIds);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private Tuple<Dictionary<ElementId, List<ElementId>>, Dictionary<string, bool>, Dictionary<string, string>> 
            BuildRoomConnectivityGraph(Document doc, List<Room> rooms, List<FamilyInstance> doors, List<Element> stairs, Dictionary<string, long> timings)
        {
            Stopwatch stepTimer = new Stopwatch();
            Dictionary<ElementId, List<ElementId>> connections = new Dictionary<ElementId, List<ElementId>>();
            Dictionary<string, bool> verticalConnectionFlags = new Dictionary<string, bool>();
            Dictionary<string, string> debugInfo = new Dictionary<string, string>();
            
            // Initialize dictionary
            stepTimer.Restart();
            foreach (Room room in rooms)
            {
                connections[room.Id] = new List<ElementId>();
            }
            timings["Graph Init"] = stepTimer.ElapsedMilliseconds;
            
            // Get active phase
            stepTimer.Restart();
            Phase activePhase = GetActivePhase(doc);
            debugInfo["ActivePhase"] = activePhase?.Name ?? "None";
            timings["Phase Detection"] = stepTimer.ElapsedMilliseconds;
            
            // Process door connections (horizontal)
            stepTimer.Restart();
            int doorConnectionCount = ProcessDoorConnections(doc, doors, connections, activePhase, timings);
            debugInfo["DoorConnections"] = $"{doorConnectionCount} connections from {doors.Count} doors";
            long doorTotalTime = stepTimer.ElapsedMilliseconds;
            timings["Door Processing Total"] = doorTotalTime;
            
            // Process stair connections (vertical)
            stepTimer.Restart();
            int stairConnectionCount = ProcessStairConnections(doc, rooms, stairs, connections, verticalConnectionFlags, activePhase, debugInfo);
            debugInfo["StairConnections"] = $"{stairConnectionCount} vertical connections from {stairs.Count} stairs";
            timings["Stair Processing"] = stepTimer.ElapsedMilliseconds;
            
            return new Tuple<Dictionary<ElementId, List<ElementId>>, Dictionary<string, bool>, Dictionary<string, string>>(
                connections, verticalConnectionFlags, debugInfo);
        }

        private Phase GetActivePhase(Document doc)
        {
            View activeView = doc.ActiveView;
            if (activeView != null)
            {
                Parameter phaseParam = activeView.get_Parameter(BuiltInParameter.VIEW_PHASE);
                if (phaseParam != null)
                {
                    Phase phase = doc.GetElement(phaseParam.AsElementId()) as Phase;
                    if (phase != null) return phase;
                }
            }
            
            // Fallback to last phase
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Phase))
                .Cast<Phase>()
                .LastOrDefault();
        }

        private int ProcessDoorConnections(Document doc, List<FamilyInstance> doors, 
            Dictionary<ElementId, List<ElementId>> connections, Phase activePhase, Dictionary<string, long> timings)
        {
            Stopwatch timer = new Stopwatch();
            int connectionCount = 0;
            int doorsWithFromTo = 0;
            int doorsNeedingRoomAtPoint = 0;
            long fromToTime = 0;
            long roomAtPointTime = 0;
            
            timer.Restart();
            
            foreach (FamilyInstance door in doors)
            {
                // Try built-in FromRoom/ToRoom first (fastest)
                Stopwatch doorTimer = Stopwatch.StartNew();
                Room room1 = door.FromRoom;
                Room room2 = door.ToRoom;
                fromToTime += doorTimer.ElapsedMilliseconds;
                
                if (room1 != null && room2 != null)
                {
                    doorsWithFromTo++;
                }
                else
                {
                    // Only use GetRoomAtPoint if necessary
                    if (activePhase != null)
                    {
                        doorsNeedingRoomAtPoint++;
                        doorTimer.Restart();
                        
                        LocationPoint doorLoc = door.Location as LocationPoint;
                        if (doorLoc != null)
                        {
                            XYZ doorPoint = doorLoc.Point;
                            Transform transform = door.GetTransform();
                            XYZ doorDir = transform.BasisY;
                            
                            // Test points on either side
                            if (room1 == null)
                            {
                                XYZ testPoint = doorPoint + doorDir * 1.0;
                                room1 = doc.GetRoomAtPoint(testPoint, activePhase);
                            }
                            
                            if (room2 == null)
                            {
                                XYZ testPoint = doorPoint - doorDir * 1.0;
                                room2 = doc.GetRoomAtPoint(testPoint, activePhase);
                            }
                            
                            // Try perpendicular if still missing
                            if (room1 == null || room2 == null)
                            {
                                doorDir = transform.BasisX;
                                
                                if (room1 == null)
                                {
                                    XYZ testPoint = doorPoint + doorDir * 1.0;
                                    room1 = doc.GetRoomAtPoint(testPoint, activePhase);
                                }
                                
                                if (room2 == null)
                                {
                                    XYZ testPoint = doorPoint - doorDir * 1.0;
                                    room2 = doc.GetRoomAtPoint(testPoint, activePhase);
                                }
                            }
                        }
                        
                        roomAtPointTime += doorTimer.ElapsedMilliseconds;
                    }
                }
                
                // Add connection if valid
                if (room1 != null && room2 != null && room1.Id != room2.Id)
                {
                    connectionCount++;
                    
                    if (!connections[room1.Id].Contains(room2.Id))
                        connections[room1.Id].Add(room2.Id);
                    
                    if (!connections[room2.Id].Contains(room1.Id))
                        connections[room2.Id].Add(room1.Id);
                }
            }
            
            timings["Door FromTo Time"] = fromToTime;
            timings["Door RoomAtPoint Time"] = roomAtPointTime;
            timings["Doors with FromTo"] = doorsWithFromTo;
            timings["Doors needing RoomAtPoint"] = doorsNeedingRoomAtPoint;
            
            return connectionCount;
        }

        private int ProcessStairConnections(Document doc, List<Room> rooms, List<Element> stairs, 
            Dictionary<ElementId, List<ElementId>> connections, Dictionary<string, bool> verticalConnectionFlags,
            Phase activePhase, Dictionary<string, string> debugInfo)
        {
            int totalConnections = 0;
            
            // Group stairs by horizontal location (stair shafts)
            Dictionary<string, List<Element>> stairShafts = new Dictionary<string, List<Element>>();
            double groupingRadius = 10.0; // 10 feet radius to group stairs
            
            foreach (Element stair in stairs)
            {
                BoundingBoxXYZ bb = stair.get_BoundingBox(null);
                if (bb == null) continue;
                
                XYZ stairCenter = new XYZ((bb.Min.X + bb.Max.X) / 2, (bb.Min.Y + bb.Max.Y) / 2, 0);
                
                // Find existing shaft or create new one
                string shaftKey = null;
                foreach (var shaft in stairShafts)
                {
                    Element firstStair = shaft.Value.First();
                    BoundingBoxXYZ firstBB = firstStair.get_BoundingBox(null);
                    XYZ firstCenter = new XYZ((firstBB.Min.X + firstBB.Max.X) / 2, (firstBB.Min.Y + firstBB.Max.Y) / 2, 0);
                    
                    double distance = stairCenter.DistanceTo(firstCenter);
                    if (distance < groupingRadius)
                    {
                        shaftKey = shaft.Key;
                        break;
                    }
                }
                
                if (shaftKey == null)
                {
                    shaftKey = $"Shaft_{stairShafts.Count + 1}";
                    stairShafts[shaftKey] = new List<Element>();
                }
                
                stairShafts[shaftKey].Add(stair);
            }
            
            debugInfo["StairShafts"] = $"{stairShafts.Count} vertical circulation shafts detected";
            
            // Process each stair shaft
            int shaftIndex = 0;
            foreach (var shaft in stairShafts)
            {
                shaftIndex++;
                List<Room> shaftRooms = new List<Room>();
                Dictionary<Room, double> roomElevations = new Dictionary<Room, double>();
                
                // Find all rooms containing stairs in this shaft
                foreach (Element stair in shaft.Value)
                {
                    BoundingBoxXYZ bb = stair.get_BoundingBox(null);
                    if (bb == null) continue;
                    
                    // Test multiple points along the stair height
                    double bottomZ = bb.Min.Z;
                    double topZ = bb.Max.Z;
                    XYZ centerXY = new XYZ((bb.Min.X + bb.Max.X) / 2, (bb.Min.Y + bb.Max.Y) / 2, 0);
                    
                    // Test at various heights
                    int numTests = Math.Max(2, (int)((topZ - bottomZ) / 10) + 1);
                    for (int i = 0; i < numTests; i++)
                    {
                        double z = bottomZ + (topZ - bottomZ) * i / (numTests - 1);
                        if (i == 0) z += 1.0; // Offset from floor
                        if (i == numTests - 1) z -= 1.0; // Offset from ceiling
                        
                        XYZ testPoint = new XYZ(centerXY.X, centerXY.Y, z);
                        Room room = doc.GetRoomAtPoint(testPoint, activePhase);
                        
                        if (room != null && !shaftRooms.Any(r => r.Id == room.Id))
                        {
                            shaftRooms.Add(room);
                            roomElevations[room] = GetRoomCenter(room).Z;
                        }
                    }
                }
                
                // Sort rooms by elevation
                shaftRooms = shaftRooms.OrderBy(r => roomElevations[r]).ToList();
                
                // Connect adjacent rooms in the vertical chain
                int shaftConnections = 0;
                for (int i = 0; i < shaftRooms.Count - 1; i++)
                {
                    Room lowerRoom = shaftRooms[i];
                    Room upperRoom = shaftRooms[i + 1];
                    
                    double z1 = roomElevations[lowerRoom];
                    double z2 = roomElevations[upperRoom];
                    
                    if (Math.Abs(z2 - z1) > 1.0) // Different levels
                    {
                        shaftConnections++;
                        totalConnections++;
                        
                        if (!connections[lowerRoom.Id].Contains(upperRoom.Id))
                            connections[lowerRoom.Id].Add(upperRoom.Id);
                        
                        if (!connections[upperRoom.Id].Contains(lowerRoom.Id))
                            connections[upperRoom.Id].Add(lowerRoom.Id);
                        
                        string connectionKey = GetConnectionKey(lowerRoom.Id, upperRoom.Id);
                        verticalConnectionFlags[connectionKey] = true;
                    }
                }
                
                if (shaftRooms.Count > 0)
                {
                    string roomNames = string.Join(" -> ", shaftRooms.Take(5).Select(r => 
                        r.get_Parameter(BuiltInParameter.ROOM_NAME)?.AsString() ?? "Unnamed"));
                    debugInfo[$"Shaft{shaftIndex}"] = $"{shaftConnections} connections: {roomNames}";
                }
            }
            
            return totalConnections;
        }

        private XYZ GetRoomCenter(Room room)
        {
            if (room == null || room.Area <= 0) return null;
            
            LocationPoint locPoint = room.Location as LocationPoint;
            if (locPoint != null)
                return locPoint.Point;
            
            BoundingBoxXYZ bb = room.get_BoundingBox(null);
            if (bb != null)
                return (bb.Min + bb.Max) / 2.0;
            
            return null;
        }

        private string GetConnectionKey(ElementId id1, ElementId id2)
        {
            int val1 = id1.IntegerValue;
            int val2 = id2.IntegerValue;
            return val1 < val2 ? $"{val1}_{val2}" : $"{val2}_{val1}";
        }
    }
}
