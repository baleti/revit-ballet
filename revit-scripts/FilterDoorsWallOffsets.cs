using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace FilterDoorsWallOffsets
{
    [Transaction(TransactionMode.Manual)]
    public class FilterDoorsWallOffsets : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            try
            {
                // Get current selection
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                
                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("Error", "Please select one or more doors before running this command.");
                    return Result.Failed;
                }
                
                // Filter selection to get only doors
                List<FamilyInstance> selectedDoors = new List<FamilyInstance>();
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is FamilyInstance fi && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
                    {
                        selectedDoors.Add(fi);
                    }
                }
                
                if (selectedDoors.Count == 0)
                {
                    TaskDialog.Show("Error", "No doors found in the current selection. Please select one or more doors.");
                    return Result.Failed;
                }
                
                // Group doors by level for efficient wall caching
                var doorsByLevel = selectedDoors
                    .Where(d => d.LevelId != null && d.LevelId != ElementId.InvalidElementId)
                    .GroupBy(d => d.LevelId)
                    .ToDictionary(g => g.Key, g => g.ToList());
                
                // Cache walls by level
                Dictionary<ElementId, List<WallData>> wallsByLevel = new Dictionary<ElementId, List<WallData>>();
                
                foreach (var levelGroup in doorsByLevel)
                {
                    ElementLevelFilter levelFilter = new ElementLevelFilter(levelGroup.Key);
                    
                    var levelWalls = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .OfCategory(BuiltInCategory.OST_Walls)
                        .WherePasses(levelFilter)
                        .Cast<Wall>()
                        .Where(w => w.Location is LocationCurve)
                        .Select(w => new WallData
                        {
                            WallId = w.Id,
                            Wall = w,
                            Curve = (w.Location as LocationCurve).Curve,
                            BoundingBox = w.get_BoundingBox(null)
                        })
                        .ToList();
                    
                    wallsByLevel[levelGroup.Key] = levelWalls;
                }
                
                // Get walls without level assignment
                var wallsWithoutLevel = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .Cast<Wall>()
                    .Where(w => (w.LevelId == null || w.LevelId == ElementId.InvalidElementId) 
                            && w.Location is LocationCurve)
                    .Select(w => new WallData
                    {
                        WallId = w.Id,
                        Wall = w,
                        Curve = (w.Location as LocationCurve).Curve,
                        BoundingBox = w.get_BoundingBox(null)
                    })
                    .ToList();
                
                // Process doors in parallel
                ConcurrentBag<DoorProcessingResult> doorResults = new ConcurrentBag<DoorProcessingResult>();
                
                Parallel.ForEach(selectedDoors, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, doorInst =>
                {
                    DoorProcessingResult doorResult = new DoorProcessingResult
                    {
                        Door = doorInst,
                        AdjacentWalls = new List<AdjacentWallInfo>()
                    };
                    
                    try
                    {
                        // Get host wall info
                        Wall hostWall = doorInst.Host as Wall;
                        if (hostWall == null)
                        {
                            doorResult.NoHostWall = true;
                        }
                        else
                        {
                            // Get door orientation
                            DoorOrientation doorOrientation = GetDoorOrientation(doorInst, hostWall);
                            
                            // Get appropriate wall list
                            List<WallData> wallsToCheck = new List<WallData>();
                            if (wallsByLevel.ContainsKey(doorInst.LevelId))
                            {
                                wallsToCheck.AddRange(wallsByLevel[doorInst.LevelId]);
                            }
                            wallsToCheck.AddRange(wallsWithoutLevel);
                            
                            // Find adjacent walls using cached data
                            doorResult.AdjacentWalls = FindAdjacentWallsParallel(
                                doorInst, 
                                hostWall, 
                                doorOrientation, 
                                wallsToCheck,
                                doc);
                        }
                    }
                    catch (Exception ex)
                    {
                        doorResult.Error = ex.Message;
                    }
                    
                    doorResults.Add(doorResult);
                });
                
                // Create dimensions first and extract their values
                Dictionary<int, List<DimensionInfo>> doorDimensions = new Dictionary<int, List<DimensionInfo>>();
                
                using (Transaction tx = new Transaction(doc, "Create Door-Wall Dimensions"))
                {
                    tx.Start();
                    
                    foreach (var result in doorResults.Where(r => !r.NoHostWall && string.IsNullOrEmpty(r.Error) && r.AdjacentWalls.Any()))
                    {
                        try
                        {
                            List<DimensionInfo> dimensionsCreated = CreateDimensionsForDoor(doc, uidoc, result.Door, result.AdjacentWalls);
                            if (dimensionsCreated.Any())
                            {
                                doorDimensions[result.Door.Id.IntegerValue] = dimensionsCreated;
                            }
                        }
                        catch (Exception)
                        {
                            // Continue with other doors if this one fails
                            continue;
                        }
                    }
                    
                    tx.Commit();
                }
                
                // Determine the unique orientation labels for dynamic columns
                var allOrientationLabels = doorDimensions.Values
                    .SelectMany(dims => dims.Select(d => d.OrientationLabel))
                    .Distinct()
                    .OrderBy(label => label) // Order: Offset Left, Offset Left Reverse, Offset Right, Offset Right Reverse
                    .ToList();
                
                // Prepare data for DataGrid with orientation-based columns
                List<string> propertyNames = new List<string>
                {
                    "Family Name", 
                    "Type Name", 
                    "Level", 
                    "Mark",
                    "Width (mm)", 
                    "Height (mm)",
                    "Adjacent Walls Count"
                };
                
                // Add dynamic orientation-based value columns first
                foreach (string orientationLabel in allOrientationLabels)
                {
                    propertyNames.Add($"{orientationLabel} (mm)");
                }
                
                // Add ID columns
                foreach (string orientationLabel in allOrientationLabels)
                {
                    propertyNames.Add($"Wall {orientationLabel.Replace("Offset ", "")} ID");
                }
                
                // Add Door Element Id at the end
                propertyNames.Add("Door Element Id");
                
                List<Dictionary<string, object>> doorData = new List<Dictionary<string, object>>();
                
                foreach (var result in doorResults.OrderBy(r => r.Door.Id.IntegerValue))
                {
                    if (result.NoHostWall || !string.IsNullOrEmpty(result.Error))
                        continue;
                    
                    Dictionary<string, object> doorProperties = new Dictionary<string, object>();
                    ElementType doorType = doc.GetElement(result.Door.GetTypeId()) as ElementType;
                    
                    // Basic door properties
                    doorProperties["Family Name"] = doorType?.FamilyName ?? "";
                    doorProperties["Type Name"] = doorType?.Name ?? "";
                    doorProperties["Level"] = doc.GetElement(result.Door.LevelId)?.Name ?? "";
                    doorProperties["Mark"] = result.Door.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString() ?? "";
                    
                    double doorWidth = doorType?.get_Parameter(BuiltInParameter.DOOR_WIDTH)?.AsDouble() ?? 0.0;
                    double doorHeight = doorType?.get_Parameter(BuiltInParameter.DOOR_HEIGHT)?.AsDouble() ?? 0.0;
                    doorProperties["Width (mm)"] = Math.Round(doorWidth * 304.8);
                    doorProperties["Height (mm)"] = Math.Round(doorHeight * 304.8);
                    
                    // Adjacent wall information
                    doorProperties["Adjacent Walls Count"] = result.AdjacentWalls.Count;
                    
                    // Get dimensions for this door
                    List<DimensionInfo> dimensions = doorDimensions.ContainsKey(result.Door.Id.IntegerValue) 
                        ? doorDimensions[result.Door.Id.IntegerValue] 
                        : new List<DimensionInfo>();
                    
                    // Initialize all orientation value columns as empty
                    foreach (string orientationLabel in allOrientationLabels)
                    {
                        doorProperties[$"{orientationLabel} (mm)"] = "-";
                    }
                    
                    // Initialize all orientation ID columns as empty  
                    foreach (string orientationLabel in allOrientationLabels)
                    {
                        doorProperties[$"Wall {orientationLabel.Replace("Offset ", "")} ID"] = "-";
                    }
                    
                    // Populate orientation-based columns with actual dimension values
                    foreach (var dim in dimensions)
                    {
                        double distanceInMm = Math.Round(dim.Value * 304.8);
                        
                        doorProperties[$"{dim.OrientationLabel} (mm)"] = distanceInMm;
                        doorProperties[$"Wall {dim.OrientationLabel.Replace("Offset ", "")} ID"] = dim.WallId.IntegerValue;
                    }
                    
                    // Add Door Element Id at the end
                    doorProperties["Door Element Id"] = result.Door.Id.IntegerValue;
                    
                    doorData.Add(doorProperties);
                }
                
                if (doorData.Count == 0)
                {
                    TaskDialog.Show("Result", "No doors with adjacent walls found for wall offset analysis.");
                    return Result.Failed;
                }
                
                // Show DataGrid
                List<Dictionary<string, object>> selectedFromGrid = CustomGUIs.DataGrid(doorData, propertyNames, false);
                
                // Show results summary
                List<string> dimensionResults = new List<string>();
                foreach (var kvp in doorDimensions)
                {
                    var orientations = kvp.Value.Select(d => d.OrientationLabel).ToList();
                    var requiresBothSidesCount = kvp.Value.Count(d => d.RequiresBothSides);
                    string summary = $"Door {kvp.Key}: {kvp.Value.Count} dimension(s) created ({string.Join(", ", orientations)})";
                    if (requiresBothSidesCount > 0)
                    {
                        summary += $" - {requiresBothSidesCount} wall(s) span both sides";
                    }
                    dimensionResults.Add(summary);
                }
                
                // Update selection if user selected specific doors
                if (selectedFromGrid?.Any() == true)
                {
                    var finalSelection = selectedDoors
                        .Where(d => selectedFromGrid.Any(s => 
                            (int)s["Door Element Id"] == d.Id.IntegerValue))
                        .Select(d => d.Id)
                        .ToList();
                    
                    uidoc.Selection.SetElementIds(finalSelection);
                }
                
                // Show results
                string summaryMessage = $"Door wall offset analysis completed.\n\n";
                if (dimensionResults.Any())
                {
                    summaryMessage += "Dimensions created:\n" + string.Join("\n", dimensionResults);
                }
                else
                {
                    summaryMessage += "No dimensions were created (no valid door-wall pairs found).";
                }
                
                TaskDialog.Show("FilterDoorsWallOffsets", summaryMessage);
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
        
        /// <summary>
        /// Create dimensions between a door and its adjacent walls
        /// </summary>
        private List<DimensionInfo> CreateDimensionsForDoor(Document doc, UIDocument uidoc, FamilyInstance door, List<AdjacentWallInfo> adjacentWalls)
        {
            List<DimensionInfo> createdDimensions = new List<DimensionInfo>();
            
            // Get door references for dimensioning
            Reference doorLeft = door.GetReferences(FamilyInstanceReferenceType.Left)?.FirstOrDefault();
            Reference doorRight = door.GetReferences(FamilyInstanceReferenceType.Right)?.FirstOrDefault();
            
            // Fallback if edges are missing
            if (doorLeft == null || doorRight == null)
            {
                Reference centerLR = door.GetReferences(FamilyInstanceReferenceType.CenterLeftRight)?.FirstOrDefault();
                if (centerLR != null)
                {
                    doorLeft = centerLR;
                    doorRight = centerLR;
                }
                else
                {
                    return createdDimensions; // Cannot create dimensions without door references
                }
            }
            
            Wall hostWall = door.Host as Wall;
            if (hostWall == null) return createdDimensions;
            
            XYZ hostWallDir = GetWallDirection(hostWall);
            // Vector perpendicular to wall direction (for dimensioning parallel to wall)
            XYZ perpToWall = new XYZ(-hostWallDir.Y, hostWallDir.X, 0.0).Normalize();
            
            // Get door orientation data to ensure consistency with DataGrid calculations
            DoorOrientation doorOrientation = GetDoorOrientation(door, hostWall);
            
            // Create dimensions for each adjacent wall
            foreach (var wallInfo in adjacentWalls)
            {
                try
                {
                    Reference wallRef = wallInfo.ClosestFaceReference;
                    if (wallRef == null)
                    {
                        // Fallback if closest face wasn't determined earlier
                        wallRef = GetClosestFaceReference(wallInfo.Wall, doorOrientation.DoorPoint, doc);
                        if (wallRef == null) continue;
                    }
                    
                    // Check if wall is roughly perpendicular to host wall (for dimensioning)
                    XYZ wallDir = GetWallDirection(wallInfo.Wall);
                    if (!AreRoughlyPerpendicular(hostWallDir, wallDir)) continue;
                    
                    // Create dimensions based on wall side information
                    foreach (var sideInfo in wallInfo.WallSides)
                    {
                        // Use the position and side information from the wallInfo
                        Reference doorRef = null;
                        XYZ doorRefPoint = XYZ.Zero;
                        string orientationLabel = "";
                        
                        // Fixed: Swap left and right labels
                        switch (wallInfo.Position)
                        {
                            case WallPosition.Left:
                                doorRef = doorLeft;
                                doorRefPoint = doorOrientation.LeftEdge;
                                orientationLabel = sideInfo == WallSide.Front ? "Offset Right" : "Offset Right Reverse";
                                break;
                            case WallPosition.Right:
                                doorRef = doorRight;
                                doorRefPoint = doorOrientation.RightEdge;
                                orientationLabel = sideInfo == WallSide.Front ? "Offset Left" : "Offset Left Reverse";
                                break;
                            case WallPosition.Front:
                                // For front walls, choose the closer edge but use orientation data
                                XYZ wallPoint = GetReferencePoint(doc, wallRef);
                                double distToLeft = doorOrientation.LeftEdge.DistanceTo(wallPoint);
                                double distToRight = doorOrientation.RightEdge.DistanceTo(wallPoint);
                                
                                if (distToLeft < distToRight)
                                {
                                    doorRef = doorLeft;
                                    doorRefPoint = doorOrientation.LeftEdge;
                                    orientationLabel = sideInfo == WallSide.Front ? "Offset Right" : "Offset Right Reverse";
                                }
                                else
                                {
                                    doorRef = doorRight;
                                    doorRefPoint = doorOrientation.RightEdge;
                                    orientationLabel = sideInfo == WallSide.Front ? "Offset Left" : "Offset Left Reverse";
                                }
                                break;
                        }
                        
                        if (doorRef == null) continue;
                        
                        // Get wall reference point
                        XYZ wallRefPoint = GetReferencePoint(doc, wallRef);
                        
                        // Create a simple 2-point dimension
                        var refArray = new ReferenceArray();
                        refArray.Append(doorRef);
                        refArray.Append(wallRef);
                        
                        // Create dimension line parallel to host wall
                        // Project both points onto a line parallel to the host wall
                        XYZ midPoint = (doorRefPoint + wallRefPoint) * 0.5;
                        
                        // Offset the dimension line from the geometry
                        double offsetDistance = 3.0; // 3 feet offset
                        XYZ offsetVector = perpToWall * offsetDistance;
                        
                        // Determine offset direction based on which side we're dimensioning
                        if (sideInfo == WallSide.Back)
                            offsetVector = -offsetVector;
                        
                        XYZ offsetMidPoint = midPoint + offsetVector;
                        
                        // Create dimension line parallel to host wall, extending from offset midpoint
                        XYZ p0 = offsetMidPoint - hostWallDir * 10.0; // Extend 10 feet in each direction
                        XYZ p1 = offsetMidPoint + hostWallDir * 10.0;
                        
                        Line dimLine = Line.CreateBound(p0, p1);
                        
                        // Create the dimension
                        Dimension dimension = doc.Create.NewDimension(uidoc.ActiveView, dimLine, refArray);
                        
                        // Get the actual dimension value
                        double dimensionValue = 0.0;
                        if (dimension != null && dimension.Value.HasValue)
                        {
                            dimensionValue = dimension.Value.Value;
                        }
                        
                        // Store dimension info
                        createdDimensions.Add(new DimensionInfo
                        {
                            Dimension = dimension,
                            Value = dimensionValue,
                            OrientationLabel = orientationLabel,
                            WallId = wallInfo.WallId,
                            IsInFront = sideInfo == WallSide.Front,
                            RequiresBothSides = wallInfo.WallSides.Count > 1
                        });
                    }
                }
                catch (Exception)
                {
                    // Continue with next wall if this one fails
                    continue;
                }
            }
            
            return createdDimensions;
        }
        
        /// <summary>
        /// Get a point from a reference for dimension line calculation
        /// </summary>
        private XYZ GetReferencePoint(Document doc, Reference reference)
        {
            GeometryObject go = doc.GetElement(reference.ElementId).GetGeometryObjectFromReference(reference);
            
            if (go is Curve c)
                return c.Evaluate(0.5, true);
            else if (go is PlanarFace f)
                return f.Origin;
            else if (go is Solid s)
                return s.GetBoundingBox().Min;
            
            return XYZ.Zero;
        }
        
        /// <summary>
        /// Check if two wall directions are roughly perpendicular
        /// </summary>
        private bool AreRoughlyPerpendicular(XYZ dir1, XYZ dir2)
        {
            double dot = Math.Abs(dir1.DotProduct(dir2));
            return dot < 0.3; // Allow some tolerance
        }
        
        /// <summary>
        /// Get the face reference of a wall that is closest to a given point
        /// </summary>
        private static Reference GetClosestFaceReference(Wall wall, XYZ targetPoint, Document doc)
        {
            // Get both interior and exterior face references
            IList<Reference> exteriorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
            IList<Reference> interiorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior);
            
            Reference exteriorRef = exteriorFaces.FirstOrDefault();
            Reference interiorRef = interiorFaces.FirstOrDefault();
            
            // If only one face is available, return it
            if (exteriorRef == null) return interiorRef;
            if (interiorRef == null) return exteriorRef;
            
            // Get representative points from each face
            XYZ exteriorPoint = GetFaceReferencePoint(wall, exteriorRef, doc);
            XYZ interiorPoint = GetFaceReferencePoint(wall, interiorRef, doc);
            
            // Calculate distances to target point
            double distToExterior = targetPoint.DistanceTo(exteriorPoint);
            double distToInterior = targetPoint.DistanceTo(interiorPoint);
            
            // Return the closer face
            return distToInterior < distToExterior ? interiorRef : exteriorRef;
        }
        
        /// <summary>
        /// Get a representative point from a wall face reference
        /// </summary>
        private static XYZ GetFaceReferencePoint(Wall wall, Reference faceRef, Document doc)
        {
            try
            {
                // Get the face from the reference
                GeometryObject geomObj = wall.GetGeometryObjectFromReference(faceRef);
                if (geomObj is PlanarFace face)
                {
                    // Get the center of the face's bounding box
                    BoundingBoxUV bbox = face.GetBoundingBox();
                    UV center = new UV((bbox.Min.U + bbox.Max.U) / 2, (bbox.Min.V + bbox.Max.V) / 2);
                    return face.Evaluate(center);
                }
                else if (geomObj is Face genericFace)
                {
                    // Fallback for non-planar faces
                    BoundingBoxUV bbox = genericFace.GetBoundingBox();
                    UV center = new UV((bbox.Min.U + bbox.Max.U) / 2, (bbox.Min.V + bbox.Max.V) / 2);
                    return genericFace.Evaluate(center);
                }
            }
            catch
            {
                // Fallback: use wall location curve midpoint
                LocationCurve locCurve = wall.Location as LocationCurve;
                if (locCurve != null)
                {
                    return locCurve.Curve.Evaluate(0.5, true);
                }
            }
            
            // Last resort: use wall bounding box center
            BoundingBoxXYZ bb = wall.get_BoundingBox(null);
            return (bb.Min + bb.Max) * 0.5;
        }
        
        /// <summary>
        /// Get wall direction (adapted from DrawDimensions)
        /// </summary>
        private static XYZ GetWallDirection(Wall wall)
        {
            var lc = wall.Location as LocationCurve;
            return (lc.Curve.GetEndPoint(1) - lc.Curve.GetEndPoint(0)).Normalize();
        }
        
        /// <summary>
        /// Determines the orientation of the door relative to its host wall
        /// </summary>
        private DoorOrientation GetDoorOrientation(FamilyInstance doorInst, Wall hostWall)
        {
            DoorOrientation result = new DoorOrientation();
            
            LocationPoint doorLocation = doorInst.Location as LocationPoint;
            XYZ doorPoint = doorLocation.Point;
            
            result.DoorPoint = doorPoint;
            result.DoorFacing = doorInst.FacingOrientation;
            result.DoorHand = doorInst.HandOrientation;
            
            // Get wall direction
            Curve wallCurve = (hostWall.Location as LocationCurve).Curve;
            if (wallCurve is Line wallLine)
            {
                result.WallDirection = wallLine.Direction;
            }
            else
            {
                IntersectionResult projectResult = wallCurve.Project(doorPoint);
                double param = projectResult.Parameter;
                result.WallDirection = wallCurve.ComputeDerivatives(param, false).BasisX.Normalize();
            }
            
            double doorWidth = doorInst.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble();
            
            // Calculate door edges
            result.LeftEdge = doorPoint;
            result.RightEdge = doorPoint.Add(result.DoorHand.Multiply(doorWidth));
            
            return result;
        }
        
        /// <summary>
        /// Analyze which sides of the door a wall endpoint falls on
        /// </summary>
        private WallEndpointSide AnalyzeWallEndpointSide(XYZ endpoint, DoorOrientation doorOrientation)
        {
            // Project endpoint onto door position
            XYZ doorToEndpoint = endpoint - doorOrientation.DoorPoint;
            
            // Check if in front or back
            double facingDot = doorToEndpoint.DotProduct(doorOrientation.DoorFacing);
            bool isInFront = facingDot > 0.1; // Small tolerance
            bool isInBack = facingDot < -0.1;
            
            // Check if left or right
            double handDot = doorToEndpoint.DotProduct(doorOrientation.DoorHand);
            double doorWidth = doorOrientation.RightEdge.DistanceTo(doorOrientation.LeftEdge);
            
            bool isLeft = handDot < -0.1;
            bool isRight = handDot > doorWidth + 0.1;
            bool isCenter = !isLeft && !isRight;
            
            return new WallEndpointSide
            {
                IsInFront = isInFront,
                IsInBack = isInBack,
                IsLeft = isLeft,
                IsRight = isRight,
                IsCenter = isCenter
            };
        }
        
        /// <summary>
        /// Find adjacent walls using cached wall data (safe for parallel execution)
        /// </summary>
        private List<AdjacentWallInfo> FindAdjacentWallsParallel(FamilyInstance doorInst, Wall hostWall, DoorOrientation orientation, List<WallData> wallsToCheck, Document doc)
        {
            List<AdjacentWallInfo> adjacentWalls = new List<AdjacentWallInfo>();
            
            double searchDistance = 0.5 / 0.3048; // 500mm in feet
            BoundingBoxXYZ doorBB = doorInst.get_BoundingBox(null);
            
            XYZ min = new XYZ(
                doorBB.Min.X - searchDistance,
                doorBB.Min.Y - searchDistance,
                doorBB.Min.Z - 0.5);
                
            XYZ max = new XYZ(
                doorBB.Max.X + searchDistance,
                doorBB.Max.Y + searchDistance,
                doorBB.Max.Z + 0.5);
            
            double doorWidth = doorInst.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble();
            double extendedWidthSearch = doorWidth + searchDistance;
            
            // Get host wall curve for robust intersection checking
            Curve hostWallCurve = (hostWall.Location as LocationCurve).Curve;
            
            foreach (WallData wallData in wallsToCheck)
            {
                if (wallData.WallId == hostWall.Id)
                    continue;
                
                // Quick bounding box check
                if (wallData.BoundingBox != null)
                {
                    bool intersects = !(wallData.BoundingBox.Max.X < min.X || wallData.BoundingBox.Min.X > max.X ||
                                      wallData.BoundingBox.Max.Y < min.Y || wallData.BoundingBox.Min.Y > max.Y ||
                                      wallData.BoundingBox.Max.Z < min.Z || wallData.BoundingBox.Min.Z > max.Z);
                    
                    if (!intersects)
                        continue;
                }
                
                // Get wall direction at closest point to door
                IntersectionResult wallProjectResult = wallData.Curve.Project(orientation.DoorPoint);
                double param = wallProjectResult.Parameter;
                XYZ wallDir = wallData.Curve.ComputeDerivatives(param, false).BasisX.Normalize();
                
                // Check perpendicularity
                double dot = Math.Abs(wallDir.DotProduct(orientation.WallDirection));
                if (dot > 0.3)
                    continue;
                
                // Analyze wall endpoints to determine which sides they fall on
                XYZ wallStart = wallData.Curve.GetEndPoint(0);
                XYZ wallEnd = wallData.Curve.GetEndPoint(1);
                
                WallEndpointSide startSide = AnalyzeWallEndpointSide(wallStart, orientation);
                WallEndpointSide endSide = AnalyzeWallEndpointSide(wallEnd, orientation);
                
                // Determine which sides this wall needs dimensions for
                HashSet<WallSide> wallSides = new HashSet<WallSide>();
                
                // Check if wall endpoints are on different sides (front/back)
                bool requiresBothSides = (startSide.IsInFront && endSide.IsInBack) || 
                                       (startSide.IsInBack && endSide.IsInFront);
                
                if (requiresBothSides)
                {
                    // Wall spans both front and back
                    wallSides.Add(WallSide.Front);
                    wallSides.Add(WallSide.Back);
                }
                else if (startSide.IsInFront || endSide.IsInFront)
                {
                    wallSides.Add(WallSide.Front);
                }
                else if (startSide.IsInBack || endSide.IsInBack)
                {
                    wallSides.Add(WallSide.Back);
                }
                
                // Skip walls that are completely outside the search area
                if (wallSides.Count == 0)
                    continue;
                
                // Determine wall position (left/right/front)
                XYZ closestPointOnWall = wallProjectResult.XYZPoint;
                XYZ doorToWallVector = closestPointOnWall - orientation.DoorPoint;
                double projectionAlongDoorHand = doorToWallVector.DotProduct(orientation.DoorHand);
                
                WallPosition position;
                double distance;
                
                if (projectionAlongDoorHand < -0.1)
                {
                    position = WallPosition.Left;
                    distance = wallData.Curve.Distance(orientation.LeftEdge);
                }
                else if (projectionAlongDoorHand > doorWidth + 0.1)
                {
                    position = WallPosition.Right;
                    distance = wallData.Curve.Distance(orientation.RightEdge);
                }
                else
                {
                    position = WallPosition.Front;
                    // For front walls, use the minimum distance to either edge
                    double distToLeft = wallData.Curve.Distance(orientation.LeftEdge);
                    double distToRight = wallData.Curve.Distance(orientation.RightEdge);
                    distance = Math.Min(distToLeft, distToRight);
                }
                
                // Check if wall is within search distance
                if (distance > extendedWidthSearch)
                    continue;
                
                // Get the closest face reference for this wall
                Reference closestFaceRef = GetClosestFaceReference(wallData.Wall, orientation.DoorPoint, doc);
                
                // Recalculate distance based on the closest face for more accuracy
                if (closestFaceRef != null)
                {
                    XYZ facePoint = GetFaceReferencePoint(wallData.Wall, closestFaceRef, doc);
                    if (position == WallPosition.Left)
                    {
                        distance = orientation.LeftEdge.DistanceTo(facePoint);
                    }
                    else if (position == WallPosition.Right)
                    {
                        distance = orientation.RightEdge.DistanceTo(facePoint);
                    }
                    else // Front
                    {
                        double distToLeft = orientation.LeftEdge.DistanceTo(facePoint);
                        double distToRight = orientation.RightEdge.DistanceTo(facePoint);
                        distance = Math.Min(distToLeft, distToRight);
                    }
                }
                
                // Add wall info
                AdjacentWallInfo wallInfo = new AdjacentWallInfo
                {
                    WallId = wallData.WallId,
                    Wall = wallData.Wall,
                    Position = position,
                    Distance = distance,
                    WallSides = wallSides.ToList(),
                    RequiresBothSides = requiresBothSides,
                    ClosestFaceReference = closestFaceRef
                };
                
                adjacentWalls.Add(wallInfo);
            }
            
            return adjacentWalls;
        }
    }
    
    /// <summary>
    /// Data class to store door orientation information
    /// </summary>
    public class DoorOrientation
    {
        public XYZ DoorPoint { get; set; }
        public XYZ DoorFacing { get; set; }
        public XYZ DoorHand { get; set; }
        public XYZ WallDirection { get; set; }
        public XYZ LeftEdge { get; set; }
        public XYZ RightEdge { get; set; }
    }
    
    /// <summary>
    /// Information about which side(s) a wall endpoint falls on
    /// </summary>
    public class WallEndpointSide
    {
        public bool IsInFront { get; set; }
        public bool IsInBack { get; set; }
        public bool IsLeft { get; set; }
        public bool IsRight { get; set; }
        public bool IsCenter { get; set; }
    }
    
    /// <summary>
    /// Cached wall data for parallel processing
    /// </summary>
    public class WallData
    {
        public ElementId WallId { get; set; }
        public Wall Wall { get; set; }
        public Curve Curve { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
    }
    
    /// <summary>
    /// Result of processing a single door
    /// </summary>
    public class DoorProcessingResult
    {
        public FamilyInstance Door { get; set; }
        public List<AdjacentWallInfo> AdjacentWalls { get; set; }
        public bool NoHostWall { get; set; }
        public string Error { get; set; }
    }
    
    /// <summary>
    /// Information about an adjacent wall
    /// </summary>
    public class AdjacentWallInfo
    {
        public ElementId WallId { get; set; }
        public Wall Wall { get; set; }
        public WallPosition Position { get; set; }
        public double Distance { get; set; }
        public List<WallSide> WallSides { get; set; } // Which sides (front/back) this wall is on
        public bool RequiresBothSides { get; set; } // If wall needs dimensions on both sides
        public Reference ClosestFaceReference { get; set; } // The face reference closest to the door
        
        [Obsolete("Use RequiresBothSides instead")]
        public bool PassesThrough 
        { 
            get { return RequiresBothSides; } 
            set { RequiresBothSides = value; }
        }
    }
    
    /// <summary>
    /// Information about a created dimension
    /// </summary>
    public class DimensionInfo
    {
        public Dimension Dimension { get; set; }
        public double Value { get; set; }
        public string OrientationLabel { get; set; }
        public ElementId WallId { get; set; }
        public bool IsInFront { get; set; }
        public bool RequiresBothSides { get; set; }
        
        [Obsolete("Use RequiresBothSides instead")]
        public bool PassesThrough 
        { 
            get { return RequiresBothSides; } 
            set { RequiresBothSides = value; }
        }
    }
    
    /// <summary>
    /// Wall position relative to door
    /// </summary>
    public enum WallPosition
    {
        Left,
        Right,
        Front
    }
    
    /// <summary>
    /// Which side of the door the wall is on (front facing or back)
    /// </summary>
    public enum WallSide
    {
        Front,
        Back
    }
}
