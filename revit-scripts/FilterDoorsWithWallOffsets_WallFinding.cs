using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace FilterDoorsWithWallOffsets
{
    public static class WallFinding
    {
        public static DoorOrientation GetDoorOrientation(FamilyInstance doorInst, Wall hostWall)
        {
            DoorOrientation result = new DoorOrientation();
            LocationPoint doorLocation = doorInst.Location as LocationPoint;
            XYZ doorPoint = doorLocation.Point;

            result.DoorPoint = doorPoint;
            result.DoorFacing = doorInst.FacingOrientation;
            result.DoorHand = doorInst.HandOrientation;

            Curve wallCurve = (hostWall.Location as LocationCurve).Curve;
            result.HostWallCurve = wallCurve;

            if (wallCurve is Line wallLine)
            {
                result.WallDirection = wallLine.Direction.Normalize();
            }
            else
            {
                IntersectionResult projectResult = wallCurve.Project(doorPoint);
                double param = projectResult.Parameter;
                result.WallDirection = wallCurve.ComputeDerivatives(param, false).BasisX.Normalize();
            }

            if (!result.WallDirection.IsZeroLength() && !result.WallDirection.IsUnitLength())
            {
                result.WallDirection = result.WallDirection.Normalize();
            }

            double doorWidth = 0;
            Parameter doorWidthParam = doorInst.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH) ?? doorInst.get_Parameter(BuiltInParameter.DOOR_WIDTH);
            if (doorWidthParam != null && doorWidthParam.HasValue)
            {
                doorWidth = doorWidthParam.AsDouble();
            }
            result.DoorWidth = doorWidth;

            if (!result.DoorHand.IsZeroLength() && !result.DoorHand.IsUnitLength())
            {
                result.DoorHand = result.DoorHand.Normalize();
            }
            if (!result.DoorFacing.IsZeroLength() && !result.DoorFacing.IsUnitLength())
            {
                result.DoorFacing = result.DoorFacing.Normalize();
            }

            result.LeftEdge = doorPoint;
            result.RightEdge = doorPoint.Add(result.DoorHand.Multiply(doorWidth));

            IntersectionResult leftProj = wallCurve.Project(result.LeftEdge);
            IntersectionResult rightProj = wallCurve.Project(result.RightEdge);
            result.DoorStartParam = Math.Min(leftProj.Parameter, rightProj.Parameter);
            result.DoorEndParam = Math.Max(leftProj.Parameter, rightProj.Parameter);

            return result;
        }

        public static List<AdjacentWallInfo> FindAdjacentWallsWithDiagnostics(
            FamilyInstance doorInst,
            Wall hostWall,
            DoorOrientation orientation,
            List<WallData> wallsToCheck,
            Document doc,
            out List<WallDiagnosticInfo> diagnostics)
        {
            diagnostics = new List<WallDiagnosticInfo>();
            List<AdjacentWallInfo> adjacentWalls = new List<AdjacentWallInfo>();

            double searchDistance = 1.0 / 0.3048; // 1 meter in feet
            BoundingBoxXYZ doorBB = doorInst.get_BoundingBox(doc.ActiveView);
            if (doorBB == null) doorBB = doorInst.get_BoundingBox(null);

            XYZ min = new XYZ(
                doorBB.Min.X - searchDistance, doorBB.Min.Y - searchDistance, doorBB.Min.Z - 0.5);
            XYZ max = new XYZ(
                doorBB.Max.X + searchDistance, doorBB.Max.Y + searchDistance, doorBB.Max.Z + 0.5);

            double doorWidth = orientation.DoorWidth;
            double extendedWidthSearch = doorWidth + searchDistance;

            // STEP 1: Filter walls within search area
            List<WallData> wallsInSearchArea = new List<WallData>();
            foreach (WallData wallData in wallsToCheck)
            {
                if (wallData.WallId == hostWall.Id) 
                    continue; // Skip host wall, but don't add to diagnostics

                var diagnostic = Diagnostics.CreateWallDiagnosticInfo(wallData, doorInst);
                
                bool withinSearchArea = true;
                if (wallData.BoundingBox != null)
                {
                    withinSearchArea = !(wallData.BoundingBox.Max.X < min.X || wallData.BoundingBox.Min.X > max.X ||
                                        wallData.BoundingBox.Max.Y < min.Y || wallData.BoundingBox.Min.Y > max.Y ||
                                        wallData.BoundingBox.Max.Z < min.Z || wallData.BoundingBox.Min.Z > max.Z);
                }

                if (withinSearchArea)
                {
                    wallsInSearchArea.Add(wallData);
                    diagnostics.Add(diagnostic); // Only add walls within search area to diagnostics
                }
            }

            foreach (WallData wallData in wallsInSearchArea)
            {
                var diagnostic = diagnostics.FirstOrDefault(d => d.WallId == wallData.WallId);
                if (diagnostic == null) continue;

                // STEP 2: Check perpendicularity
                XYZ wallDir = GetWallDirection(wallData.Wall);
                double perpendicularityDotProduct = Math.Abs(wallDir.DotProduct(orientation.WallDirection));
                bool isPerpendicularToHost = perpendicularityDotProduct < 0.3;

                if (!isPerpendicularToHost)
                {
                    diagnostic.SetRejected(WallProcessingResult.RejectedNotPerpendicular, 
                        $"Dot product: {perpendicularityDotProduct:F3} (threshold: 0.3)");
                    continue;
                }

                // STEP 3: Check if wall endpoints are too far in front/behind the door
                if (AreWallEndpointsTooFarFromDoor(wallData, orientation))
                {
                    diagnostic.SetRejected(WallProcessingResult.RejectedEndpointsTooFar, 
                        "Both endpoints >500mm from door front/back");
                    continue;
                }

                // STEP 4: Check if wall is problematically positioned directly in front of or behind the door
                if (IsWallDirectlyInFrontOrBehindDoor(wallData, orientation))
                {
                    diagnostic.SetRejected(WallProcessingResult.RejectedDirectlyInFrontOrBehind, 
                        "Wall directly in front/behind door opening");
                    continue;
                }

                // STEP 5: OCCLUSION CHECK
                if (IsWallOccluded(wallData, orientation.DoorPoint, wallsInSearchArea, hostWall.Id, doc))
                {
                    diagnostic.SetRejected(WallProcessingResult.RejectedOccluded, 
                        "Wall hidden behind other walls");
                    continue;
                }

                // STEP 6: Analyze wall sides and position
                XYZ wallStart = wallData.Curve.GetEndPoint(0);
                XYZ wallEnd = wallData.Curve.GetEndPoint(1);
                WallEndpointSide startSideInfo = AnalyzeWallEndpointSide(wallStart, orientation);
                WallEndpointSide endSideInfo = AnalyzeWallEndpointSide(wallEnd, orientation);

                HashSet<WallSide> determinedWallSides = new HashSet<WallSide>();
                bool requiresBothSides = (startSideInfo.IsInFront && endSideInfo.IsInBack) || (startSideInfo.IsInBack && endSideInfo.IsInFront);

                if (requiresBothSides)
                {
                    determinedWallSides.Add(WallSide.Front);
                    determinedWallSides.Add(WallSide.Back);
                }
                else if (startSideInfo.IsInFront || endSideInfo.IsInFront)
                {
                    determinedWallSides.Add(WallSide.Front);
                }
                else if (startSideInfo.IsInBack || endSideInfo.IsInBack)
                {
                    determinedWallSides.Add(WallSide.Back);
                }

                if (!determinedWallSides.Any())
                {
                    diagnostic.SetRejected(WallProcessingResult.RejectedNoValidSides, 
                        "No valid wall sides determined");
                    continue;
                }

                // STEP 7: Calculate position and distance
                IntersectionResult wallProjectResult = wallData.Curve.Project(orientation.DoorPoint);
                XYZ closestPointOnWallCurve = wallProjectResult.XYZPoint;
                XYZ doorToWallVector = closestPointOnWallCurve - orientation.DoorPoint;

                XYZ normalizedDoorHand = orientation.DoorHand.IsUnitLength() ? orientation.DoorHand : orientation.DoorHand.Normalize();
                double projectionAlongDoorHand = doorToWallVector.DotProduct(normalizedDoorHand);

                WallPosition position;
                double distanceToEdge;
                double edgeTolerance = 0.1;

                if (projectionAlongDoorHand < -edgeTolerance)
                {
                    position = WallPosition.Left;
                    distanceToEdge = wallData.Curve.Distance(orientation.LeftEdge);
                }
                else if (projectionAlongDoorHand > (orientation.DoorWidth + edgeTolerance))
                {
                    position = WallPosition.Right;
                    distanceToEdge = wallData.Curve.Distance(orientation.RightEdge);
                }
                else
                {
                    position = WallPosition.Front;
                    double distToLeft = wallData.Curve.Distance(orientation.LeftEdge);
                    double distToRight = wallData.Curve.Distance(orientation.RightEdge);
                    distanceToEdge = Math.Min(distToLeft, distToRight);
                }

                // STEP 8: Check distance threshold
                if (distanceToEdge > extendedWidthSearch)
                {
                    diagnostic.SetRejected(WallProcessingResult.RejectedTooFarFromEdge, 
                        $"Distance: {distanceToEdge:F3} ft > threshold: {extendedWidthSearch:F3} ft");
                    continue;
                }

                // STEP 9: Get face reference
                Reference closestFaceRef = GetClosestFaceReference(wallData.Wall, orientation.DoorPoint, doc);
                if (closestFaceRef == null)
                {
                    diagnostic.SetRejected(WallProcessingResult.RejectedNoFaceReference, 
                        "Could not obtain wall face reference");
                    continue;
                }

                // Update distance with face reference if available
                XYZ facePoint = GetFaceReferencePoint(wallData.Wall, closestFaceRef, doc);
                if (facePoint != null)
                {
                    if (position == WallPosition.Left)
                    {
                        distanceToEdge = orientation.LeftEdge.DistanceTo(facePoint);
                    }
                    else if (position == WallPosition.Right)
                    {
                        distanceToEdge = orientation.RightEdge.DistanceTo(facePoint);
                    }
                    else // Front
                    {
                        distanceToEdge = Math.Min(orientation.LeftEdge.DistanceTo(facePoint), orientation.RightEdge.DistanceTo(facePoint));
                    }
                }

                // STEP 10: Accept the wall
                diagnostic.SetAccepted(position, determinedWallSides.ToList());

                AdjacentWallInfo wallInfo = new AdjacentWallInfo
                {
                    WallId = wallData.WallId,
                    Wall = wallData.Wall,
                    Position = position,
                    Distance = distanceToEdge,
                    WallSides = determinedWallSides.ToList(),
                    RequiresBothSides = requiresBothSides,
                    ClosestFaceReference = closestFaceRef
                };
                adjacentWalls.Add(wallInfo);
            }

            return adjacentWalls;
        }

        public static List<AdjacentWallInfo> FindAdjacentWallsParallel(
            FamilyInstance doorInst,
            Wall hostWall,
            DoorOrientation orientation,
            List<WallData> wallsToCheck,
            Document doc)
        {
            List<AdjacentWallInfo> adjacentWalls = new List<AdjacentWallInfo>();

            double searchDistance = 1.0 / 0.3048;
            BoundingBoxXYZ doorBB = doorInst.get_BoundingBox(doc.ActiveView);
            if (doorBB == null) doorBB = doorInst.get_BoundingBox(null);

            XYZ min = new XYZ(
                doorBB.Min.X - searchDistance, doorBB.Min.Y - searchDistance, doorBB.Min.Z - 0.5);
            XYZ max = new XYZ(
                doorBB.Max.X + searchDistance, doorBB.Max.Y + searchDistance, doorBB.Max.Z + 0.5);

            double doorWidth = orientation.DoorWidth;
            double extendedWidthSearch = doorWidth + searchDistance;

            // Pre-filter walls within search area for occlusion checking
            List<WallData> wallsInSearchArea = wallsToCheck.Where(wallData =>
            {
                if (wallData.WallId == hostWall.Id) return false;

                if (wallData.BoundingBox != null)
                {
                    return !(wallData.BoundingBox.Max.X < min.X || wallData.BoundingBox.Min.X > max.X ||
                            wallData.BoundingBox.Max.Y < min.Y || wallData.BoundingBox.Min.Y > max.Y ||
                            wallData.BoundingBox.Max.Z < min.Z || wallData.BoundingBox.Min.Z > max.Z);
                }
                return true;
            }).ToList();

            foreach (WallData wallData in wallsInSearchArea)
            {
                XYZ wallDir = GetWallDirection(wallData.Wall);
                double perpendicularityDotProduct = Math.Abs(wallDir.DotProduct(orientation.WallDirection));
                bool isPerpendicularToHost = perpendicularityDotProduct < 0.3;

                if (!isPerpendicularToHost)
                    continue;

                // Check if wall endpoints are too far in front/behind the door (more than 500mm)
                if (AreWallEndpointsTooFarFromDoor(wallData, orientation))
                    continue;

                // Check if wall is problematically positioned directly in front of or behind the door
                if (IsWallDirectlyInFrontOrBehindDoor(wallData, orientation))
                    continue;

                // OCCLUSION CHECK - Skip walls that are occluded by other walls
                if (IsWallOccluded(wallData, orientation.DoorPoint, wallsInSearchArea, hostWall.Id, doc))
                    continue;

                XYZ wallStart = wallData.Curve.GetEndPoint(0);
                XYZ wallEnd = wallData.Curve.GetEndPoint(1);
                WallEndpointSide startSideInfo = AnalyzeWallEndpointSide(wallStart, orientation);
                WallEndpointSide endSideInfo = AnalyzeWallEndpointSide(wallEnd, orientation);

                HashSet<WallSide> determinedWallSides = new HashSet<WallSide>();
                bool requiresBothSides = (startSideInfo.IsInFront && endSideInfo.IsInBack) || (startSideInfo.IsInBack && endSideInfo.IsInFront);

                if (requiresBothSides)
                {
                    determinedWallSides.Add(WallSide.Front);
                    determinedWallSides.Add(WallSide.Back);
                }
                else if (startSideInfo.IsInFront || endSideInfo.IsInFront)
                {
                    determinedWallSides.Add(WallSide.Front);
                }
                else if (startSideInfo.IsInBack || endSideInfo.IsInBack)
                {
                    determinedWallSides.Add(WallSide.Back);
                }

                if (!determinedWallSides.Any())
                    continue;

                IntersectionResult wallProjectResult = wallData.Curve.Project(orientation.DoorPoint);
                XYZ closestPointOnWallCurve = wallProjectResult.XYZPoint;
                XYZ doorToWallVector = closestPointOnWallCurve - orientation.DoorPoint;

                XYZ normalizedDoorHand = orientation.DoorHand.IsUnitLength() ? orientation.DoorHand : orientation.DoorHand.Normalize();
                double projectionAlongDoorHand = doorToWallVector.DotProduct(normalizedDoorHand);

                WallPosition position;
                double distanceToEdge;
                double edgeTolerance = 0.1;

                if (projectionAlongDoorHand < -edgeTolerance)
                {
                    position = WallPosition.Left;
                    distanceToEdge = wallData.Curve.Distance(orientation.LeftEdge);
                }
                else if (projectionAlongDoorHand > (orientation.DoorWidth + edgeTolerance))
                {
                    position = WallPosition.Right;
                    distanceToEdge = wallData.Curve.Distance(orientation.RightEdge);
                }
                else
                {
                    position = WallPosition.Front;
                    double distToLeft = wallData.Curve.Distance(orientation.LeftEdge);
                    double distToRight = wallData.Curve.Distance(orientation.RightEdge);
                    distanceToEdge = Math.Min(distToLeft, distToRight);
                }

                if (distanceToEdge > extendedWidthSearch)
                    continue;

                Reference closestFaceRef = GetClosestFaceReference(wallData.Wall, orientation.DoorPoint, doc);
                if (closestFaceRef != null)
                {
                    XYZ facePoint = GetFaceReferencePoint(wallData.Wall, closestFaceRef, doc);
                    if (facePoint != null)
                    {
                        if (position == WallPosition.Left)
                        {
                            distanceToEdge = orientation.LeftEdge.DistanceTo(facePoint);
                        }
                        else if (position == WallPosition.Right)
                        {
                            distanceToEdge = orientation.RightEdge.DistanceTo(facePoint);
                        }
                        else // Front
                        {
                            distanceToEdge = Math.Min(orientation.LeftEdge.DistanceTo(facePoint), orientation.RightEdge.DistanceTo(facePoint));
                        }
                    }
                }

                AdjacentWallInfo wallInfo = new AdjacentWallInfo
                {
                    WallId = wallData.WallId,
                    Wall = wallData.Wall,
                    Position = position,
                    Distance = distanceToEdge,
                    WallSides = determinedWallSides.ToList(),
                    RequiresBothSides = requiresBothSides,
                    ClosestFaceReference = closestFaceRef
                };
                adjacentWalls.Add(wallInfo);
            }
            return adjacentWalls;
        }

        /// <summary>
        /// Checks if a wall is occluded by other walls when viewed from the door point.
        /// Returns true if the wall is hidden behind other walls.
        /// </summary>
        private static bool IsWallOccluded(WallData targetWall, XYZ doorPoint, List<WallData> allWallsInArea, ElementId hostWallId, Document doc)
        {
            try
            {
                // Get target wall face reference for more accurate positioning
                Reference targetFaceRef = GetClosestFaceReference(targetWall.Wall, doorPoint, doc);
                if (targetFaceRef == null) return false;

                XYZ targetPoint = GetFaceReferencePoint(targetWall.Wall, targetFaceRef, doc);
                if (targetPoint == null) return false;

                // Create a ray from door to target wall face
                XYZ rayDirection = (targetPoint - doorPoint);
                double rayLength = rayDirection.GetLength();

                // If target wall is very close to door (less than 8 feet), don't consider it occluded
                // This handles cases where walls are adjacent to the door area
                if (rayLength < 8.0) return false;

                rayDirection = rayDirection.Normalize();

                // Use more realistic tolerances for wall thickness and clearance
                double wallThicknessTolerance = 1.0; // ~30cm tolerance for wall thickness variations
                double minDistanceForOcclusion = wallThicknessTolerance;
                double maxDistanceForOcclusion = rayLength - wallThicknessTolerance;

                if (maxDistanceForOcclusion <= minDistanceForOcclusion)
                    return false; // Target wall is too close to be occluded

                // Check each other wall for intersection with the ray
                foreach (WallData otherWall in allWallsInArea)
                {
                    // Skip the target wall itself and the host wall
                    if (otherWall.WallId == targetWall.WallId || otherWall.WallId == hostWallId)
                        continue;

                    // Skip walls that are very close to the target wall (likely adjacent/parallel)
                    double distanceBetweenWalls = GetDistanceBetweenWalls(targetWall, otherWall);
                    if (distanceBetweenWalls < 2.0) // If walls are less than 2 feet apart, they're likely adjacent
                        continue;

                    // Quick bounding box check first for performance
                    if (!DoesRayIntersectBoundingBox(doorPoint, rayDirection, rayLength, otherWall.BoundingBox))
                        continue;

                    // Check intersection with wall geometry (both faces)
                    double intersectionDistance = GetRayWallGeometryIntersectionDistance(otherWall.Wall, doorPoint, rayDirection, rayLength);

                    // If intersection is between door and target wall, then target is occluded
                    if (intersectionDistance > minDistanceForOcclusion && intersectionDistance < maxDistanceForOcclusion)
                    {
                        // Additional check: ensure the intersection is not just a grazing hit
                        // by verifying the intersection point is significantly inside the wall geometry
                        if (IsSignificantWallIntersection(otherWall.Wall, doorPoint, rayDirection, intersectionDistance, doc))
                        {
                            System.Diagnostics.Debug.WriteLine($"Wall {targetWall.WallId} is occluded by wall {otherWall.WallId} at distance {intersectionDistance:F3}");
                            return true; // Target wall is occluded by this other wall
                        }
                    }
                }

                return false; // No occlusion found
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in occlusion check for wall {targetWall.WallId}: {ex.Message}");
                return false; // If there's an error, assume not occluded to be safe
            }
        }

        /// <summary>
        /// Calculates the minimum distance between two walls
        /// </summary>
        private static double GetDistanceBetweenWalls(WallData wall1, WallData wall2)
        {
            try
            {
                // Sample points along both curves to find minimum distance
                Curve curve1 = wall1.Curve;
                Curve curve2 = wall2.Curve;

                double minDistance = double.MaxValue;

                // Sample multiple points along curve1 and find closest distance to curve2
                int sampleCount = 10;
                for (int i = 0; i <= sampleCount; i++)
                {
                    double param = i / (double)sampleCount;
                    
                    XYZ point1;
                    if (curve1.IsBound)
                    {
                        double actualParam = curve1.GetEndParameter(0) + param * (curve1.GetEndParameter(1) - curve1.GetEndParameter(0));
                        point1 = curve1.Evaluate(actualParam, false);
                    }
                    else
                    {
                        point1 = curve1.Evaluate(param, true);
                    }

                    double distanceToCurve2 = curve2.Distance(point1);
                    if (distanceToCurve2 < minDistance)
                    {
                        minDistance = distanceToCurve2;
                    }
                }

                // Also sample points along curve2 and find closest distance to curve1
                for (int i = 0; i <= sampleCount; i++)
                {
                    double param = i / (double)sampleCount;
                    
                    XYZ point2;
                    if (curve2.IsBound)
                    {
                        double actualParam = curve2.GetEndParameter(0) + param * (curve2.GetEndParameter(1) - curve2.GetEndParameter(0));
                        point2 = curve2.Evaluate(actualParam, false);
                    }
                    else
                    {
                        point2 = curve2.Evaluate(param, true);
                    }

                    double distanceToCurve1 = curve1.Distance(point2);
                    if (distanceToCurve1 < minDistance)
                    {
                        minDistance = distanceToCurve1;
                    }
                }

                // Also check endpoint distances for extra accuracy
                XYZ start1 = curve1.GetEndPoint(0);
                XYZ end1 = curve1.GetEndPoint(1);
                XYZ start2 = curve2.GetEndPoint(0);
                XYZ end2 = curve2.GetEndPoint(1);

                double[] endpointDistances = {
                    curve2.Distance(start1),
                    curve2.Distance(end1),
                    curve1.Distance(start2),
                    curve1.Distance(end2)
                };

                foreach (double dist in endpointDistances)
                {
                    if (dist < minDistance)
                    {
                        minDistance = dist;
                    }
                }

                return minDistance;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error calculating distance between walls: {ex.Message}");
                return double.MaxValue; // Return large value if calculation fails
            }
        }

        /// <summary>
        /// Checks if the ray intersection with a wall is significant (not just a grazing hit)
        /// </summary>
        private static bool IsSignificantWallIntersection(Wall wall, XYZ rayStart, XYZ rayDirection, double intersectionDistance, Document doc)
        {
            try
            {
                XYZ intersectionPoint = rayStart + rayDirection * intersectionDistance;

                // Get wall thickness for comparison
                Parameter widthParam = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                double wallThickness = widthParam?.AsDouble() ?? 0.5; // Default to 6 inches if not found

                // Check if intersection point is well within the wall geometry
                // by testing points slightly offset from the intersection
                XYZ testPoint1 = intersectionPoint + rayDirection * (wallThickness * 0.3);
                XYZ testPoint2 = intersectionPoint - rayDirection * (wallThickness * 0.3);

                // If both test points are still close to the wall, it's a significant intersection
                LocationCurve locCurve = wall.Location as LocationCurve;
                if (locCurve?.Curve != null)
                {
                    double dist1 = locCurve.Curve.Distance(testPoint1);
                    double dist2 = locCurve.Curve.Distance(testPoint2);

                    // If both points are within wall thickness distance, it's a real intersection
                    return (dist1 < wallThickness && dist2 < wallThickness);
                }

                return true; // Default to true if we can't verify
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking intersection significance: {ex.Message}");
                return true; // Default to true if calculation fails
            }
        }

        /// <summary>
        /// Quick check if a ray intersects with a bounding box for performance optimization
        /// </summary>
        private static bool DoesRayIntersectBoundingBox(XYZ rayStart, XYZ rayDirection, double rayLength, BoundingBoxXYZ bbox)
        {
            if (bbox == null) return true; // If no bounding box, assume potential intersection

            XYZ rayEnd = rayStart + rayDirection * rayLength;

            // Simple AABB vs line segment test
            XYZ min = bbox.Min;
            XYZ max = bbox.Max;

            // Check if either ray start or end is inside the bounding box
            if (IsPointInBoundingBox(rayStart, min, max) || IsPointInBoundingBox(rayEnd, min, max))
                return true;

            // Check if ray passes through the bounding box (simplified check)
            bool intersectsX = (rayStart.X <= max.X && rayEnd.X >= min.X) || (rayStart.X >= min.X && rayEnd.X <= max.X);
            bool intersectsY = (rayStart.Y <= max.Y && rayEnd.Y >= min.Y) || (rayStart.Y >= min.Y && rayEnd.Y <= max.Y);
            bool intersectsZ = (rayStart.Z <= max.Z && rayEnd.Z >= min.Z) || (rayStart.Z >= min.Z && rayEnd.Z <= max.Z);

            return intersectsX && intersectsY && intersectsZ;
        }

        private static bool IsPointInBoundingBox(XYZ point, XYZ min, XYZ max)
        {
            return point.X >= min.X && point.X <= max.X &&
                   point.Y >= min.Y && point.Y <= max.Y &&
                   point.Z >= min.Z && point.Z <= max.Z;
        }

        /// <summary>
        /// Calculates the distance along a ray where it intersects with wall geometry (faces).
        /// Returns -1 if no intersection found.
        /// </summary>
        private static double GetRayWallGeometryIntersectionDistance(Wall wall, XYZ rayStart, XYZ rayDirection, double maxRayLength)
        {
            try
            {
                // Get both exterior and interior faces of the wall
                IList<Reference> exteriorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
                IList<Reference> interiorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior);

                List<Reference> allFaces = new List<Reference>();
                if (exteriorFaces != null) allFaces.AddRange(exteriorFaces);
                if (interiorFaces != null) allFaces.AddRange(interiorFaces);

                double closestIntersection = -1;

                // Check intersection with each face
                foreach (Reference faceRef in allFaces)
                {
                    try
                    {
                        GeometryObject geomObj = wall.GetGeometryObjectFromReference(faceRef);
                        if (geomObj is PlanarFace face)
                        {
                            double distance = GetRayPlanarFaceIntersectionDistance(face, rayStart, rayDirection, maxRayLength);
                            if (distance > 0 && (closestIntersection < 0 || distance < closestIntersection))
                            {
                                closestIntersection = distance;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error checking face intersection: {ex.Message}");
                        continue;
                    }
                }

                // If no face intersection found, fall back to curve-based check with wall thickness
                if (closestIntersection < 0)
                {
                    LocationCurve locCurve = wall.Location as LocationCurve;
                    if (locCurve?.Curve != null)
                    {
                        double curveDistance = GetRayWallIntersectionDistance(rayStart, rayDirection, locCurve.Curve);
                        if (curveDistance > 0)
                        {
                            // Add half wall thickness as approximation
                            Parameter widthParam = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                            double wallThickness = widthParam?.AsDouble() ?? 0.5; // Default to 6 inches if not found
                            closestIntersection = Math.Max(0, curveDistance - wallThickness / 2.0);
                        }
                    }
                }

                return closestIntersection;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in GetRayWallGeometryIntersectionDistance: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// Calculates intersection of ray with a planar face
        /// </summary>
        private static double GetRayPlanarFaceIntersectionDistance(PlanarFace face, XYZ rayStart, XYZ rayDirection, double maxRayLength)
        {
            try
            {
                // Get face normal and origin
                XYZ faceNormal = face.FaceNormal;
                XYZ faceOrigin = face.Origin;

                // Calculate ray-plane intersection
                double denominator = rayDirection.DotProduct(faceNormal);
                if (Math.Abs(denominator) < 1e-10)
                    return -1; // Ray is parallel to face

                XYZ originToRayStart = rayStart - faceOrigin;
                double t = -originToRayStart.DotProduct(faceNormal) / denominator;

                // Check if intersection is in forward direction and within max distance
                if (t <= 0 || t > maxRayLength)
                    return -1;

                // Calculate intersection point
                XYZ intersectionPoint = rayStart + rayDirection * t;

                // Check if intersection point is within the face bounds
                IntersectionResult uvResult = face.Project(intersectionPoint);
                if (uvResult != null && uvResult.Distance < 0.1) // Small tolerance for point-on-face
                {
                    // Additional check: is the UV point within the face boundaries?
                    try
                    {
                        UV intersectionUV = uvResult.UVPoint;
                        BoundingBoxUV faceBounds = face.GetBoundingBox();

                        if (intersectionUV.U >= faceBounds.Min.U && intersectionUV.U <= faceBounds.Max.U &&
                            intersectionUV.V >= faceBounds.Min.V && intersectionUV.V <= faceBounds.Max.V)
                        {
                            return t;
                        }
                    }
                    catch
                    {
                        // If UV check fails, accept the intersection if it's close to the face
                        return t;
                    }
                }

                return -1;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in GetRayPlanarFaceIntersectionDistance: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// Fallback method: Calculates the distance along a ray where it intersects with a wall curve.
        /// Returns -1 if no intersection found.
        /// </summary>
        private static double GetRayWallIntersectionDistance(XYZ rayStart, XYZ rayDirection, Curve wallCurve)
        {
            try
            {
                // For 2D intersection (most walls are vertical), we'll work in the XY plane
                XYZ rayStart2D = new XYZ(rayStart.X, rayStart.Y, 0);
                XYZ rayDirection2D = new XYZ(rayDirection.X, rayDirection.Y, 0).Normalize();

                if (rayDirection2D.IsZeroLength())
                    return -1; // Ray is vertical, no 2D intersection

                // Convert wall curve to 2D
                XYZ wallStart = wallCurve.GetEndPoint(0);
                XYZ wallEnd = wallCurve.GetEndPoint(1);
                XYZ wallStart2D = new XYZ(wallStart.X, wallStart.Y, 0);
                XYZ wallEnd2D = new XYZ(wallEnd.X, wallEnd.Y, 0);

                // Line-line intersection in 2D
                XYZ wallDirection2D = (wallEnd2D - wallStart2D);
                if (wallDirection2D.IsZeroLength())
                    return -1; // Wall has no length

                wallDirection2D = wallDirection2D.Normalize();

                // Use parametric line intersection
                // Ray: P = rayStart2D + t * rayDirection2D
                // Wall: Q = wallStart2D + s * wallDirection2D
                // Solve: rayStart2D + t * rayDirection2D = wallStart2D + s * wallDirection2D

                double denominator = rayDirection2D.X * wallDirection2D.Y - rayDirection2D.Y * wallDirection2D.X;
                if (Math.Abs(denominator) < 1e-10)
                    return -1; // Lines are parallel

                XYZ startDiff = wallStart2D - rayStart2D;
                double t = (startDiff.X * wallDirection2D.Y - startDiff.Y * wallDirection2D.X) / denominator;
                double s = (startDiff.X * rayDirection2D.Y - startDiff.Y * rayDirection2D.X) / denominator;

                // Check if intersection is within the wall segment (s should be between 0 and wall length)
                double wallLength2D = wallStart2D.DistanceTo(wallEnd2D);
                if (s < 0 || s > wallLength2D)
                    return -1; // Intersection is outside wall segment

                // Check if intersection is in the forward direction of the ray
                if (t < 0)
                    return -1; // Intersection is behind the ray start

                return t; // Return distance along ray to intersection
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error calculating ray-wall intersection: {ex.Message}");
                return -1;
            }
        }

        // Simplified to return bool
        private static bool CheckIfWallIsInFrontOfDoor(WallData wallData, DoorOrientation orientation)
        {
            XYZ closestPointOnWallCurve = wallData.Curve.Project(orientation.DoorPoint).XYZPoint;
            XYZ doorToWallVec = closestPointOnWallCurve - orientation.DoorPoint;

            XYZ normalizedDoorHand = orientation.DoorHand.IsUnitLength() ? orientation.DoorHand : orientation.DoorHand.Normalize();
            XYZ normalizedDoorFacing = orientation.DoorFacing.IsUnitLength() ? orientation.DoorFacing : orientation.DoorFacing.Normalize();

            double projectionAlongHand = doorToWallVec.DotProduct(normalizedDoorHand);
            double tolerance = 0.1;
            bool isWithinDoorWidthAlongHost = projectionAlongHand >= -tolerance &&
                                            projectionAlongHand <= (orientation.DoorWidth + tolerance);

            bool finalIsInFrontDecision = false;

            if (isWithinDoorWidthAlongHost)
            {
                double projectionAlongFacing = doorToWallVec.DotProduct(normalizedDoorFacing);
                double minFrontOffsetForFiltering = 0.5;

                if (projectionAlongFacing > minFrontOffsetForFiltering)
                {
                    finalIsInFrontDecision = true;
                }
            }
            return finalIsInFrontDecision;
        }

        /// <summary>
        /// Checks if both wall endpoints are too far (>500mm) in front of or behind the door
        /// </summary>
        private static bool AreWallEndpointsTooFarFromDoor(WallData wallData, DoorOrientation orientation)
        {
            try
            {
                XYZ wallStart = wallData.Curve.GetEndPoint(0);
                XYZ wallEnd = wallData.Curve.GetEndPoint(1);

                XYZ normalizedDoorFacing = orientation.DoorFacing.IsUnitLength() ? orientation.DoorFacing : orientation.DoorFacing.Normalize();

                // Calculate projection of each endpoint along door facing direction
                XYZ doorToStart = wallStart - orientation.DoorPoint;
                XYZ doorToEnd = wallEnd - orientation.DoorPoint;

                double startProjection = doorToStart.DotProduct(normalizedDoorFacing);
                double endProjection = doorToEnd.DotProduct(normalizedDoorFacing);

                // Convert 500mm to feet
                double maxDistanceFeet = 500.0 / 304.8;

                // Check if both endpoints are too far in front (positive projection)
                bool bothTooFarInFront = startProjection > maxDistanceFeet && endProjection > maxDistanceFeet;

                // Check if both endpoints are too far behind (negative projection)
                bool bothTooFarBehind = startProjection < -maxDistanceFeet && endProjection < -maxDistanceFeet;

                return bothTooFarInFront || bothTooFarBehind;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in AreWallEndpointsTooFarFromDoor: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Checks if wall is directly in front of or behind the door (within door width) and problematic
        /// </summary>
        private static bool IsWallDirectlyInFrontOrBehindDoor(WallData wallData, DoorOrientation orientation)
        {
            try
            {
                XYZ closestPointOnWallCurve = wallData.Curve.Project(orientation.DoorPoint).XYZPoint;
                XYZ doorToWallVec = closestPointOnWallCurve - orientation.DoorPoint;

                XYZ normalizedDoorHand = orientation.DoorHand.IsUnitLength() ? orientation.DoorHand : orientation.DoorHand.Normalize();
                XYZ normalizedDoorFacing = orientation.DoorFacing.IsUnitLength() ? orientation.DoorFacing : orientation.DoorFacing.Normalize();

                double projectionAlongHand = doorToWallVec.DotProduct(normalizedDoorHand);
                double tolerance = 0.1;
                bool isWithinDoorWidthAlongHost = projectionAlongHand >= -tolerance &&
                                                projectionAlongHand <= (orientation.DoorWidth + tolerance);

                if (!isWithinDoorWidthAlongHost)
                    return false; // Not within door width, so not problematic

                double projectionAlongFacing = doorToWallVec.DotProduct(normalizedDoorFacing);
                double minOffsetForFiltering = 0.5; // ~150mm

                // Check if wall is directly in front (existing logic)
                bool isProblematicallyInFront = projectionAlongFacing > minOffsetForFiltering;

                // Check if wall is directly behind (new logic)
                bool isProblematicallyBehind = projectionAlongFacing < -minOffsetForFiltering;

                return isProblematicallyInFront || isProblematicallyBehind;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in IsWallDirectlyInFrontOrBehindDoor: {ex.Message}");
                return false;
            }
        }

        private static WallEndpointSide AnalyzeWallEndpointSide(XYZ endpoint, DoorOrientation doorOrientation)
        {
            XYZ doorToEndpoint = endpoint - doorOrientation.DoorPoint;

            XYZ normalizedDoorFacing = doorOrientation.DoorFacing.IsUnitLength() ? doorOrientation.DoorFacing : doorOrientation.DoorFacing.Normalize();
            XYZ normalizedDoorHand = doorOrientation.DoorHand.IsUnitLength() ? doorOrientation.DoorHand : doorOrientation.DoorHand.Normalize();

            double facingDot = doorToEndpoint.DotProduct(normalizedDoorFacing);
            bool isInFront = facingDot > 0.1;
            bool isInBack = facingDot < -0.1;

            double handDot = doorToEndpoint.DotProduct(normalizedDoorHand);
            double doorWidth = doorOrientation.DoorWidth;

            bool isLeft = handDot < -0.1;
            bool isRight = handDot > (doorWidth + 0.1);
            bool isCenter = !isLeft && !isRight;

            return new WallEndpointSide
            {
                IsInFront = isInFront, IsInBack = isInBack,
                IsLeft = isLeft, IsRight = isRight, IsCenter = isCenter
            };
        }

        private static Reference GetClosestFaceReference(Wall wall, XYZ targetPoint, Document doc)
        {
            IList<Reference> exteriorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
            IList<Reference> interiorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior);

            Reference exteriorRef = exteriorFaces.FirstOrDefault();
            Reference interiorRef = interiorFaces.FirstOrDefault();

            if (exteriorRef == null && interiorRef == null) return null;
            if (exteriorRef == null) return interiorRef;
            if (interiorRef == null) return exteriorRef;

            XYZ exteriorPoint = GetFaceReferencePoint(wall, exteriorRef, doc);
            XYZ interiorPoint = GetFaceReferencePoint(wall, interiorRef, doc);

            if (exteriorPoint == null && interiorPoint == null) return null;
            if (exteriorPoint == null) return interiorRef;
            if (interiorPoint == null) return exteriorRef;

            double distToExterior = targetPoint.DistanceTo(exteriorPoint);
            double distToInterior = targetPoint.DistanceTo(interiorPoint);

            return distToInterior < distToExterior ? interiorRef : exteriorRef;
        }

        private static XYZ GetFaceReferencePoint(Wall wall, Reference faceRef, Document doc)
        {
            try
            {
                GeometryObject geomObj = wall.GetGeometryObjectFromReference(faceRef);
                if (geomObj is PlanarFace pFace)
                {
                    BoundingBoxUV bboxUV = pFace.GetBoundingBox();
                    UV centerUV = (bboxUV.Min + bboxUV.Max) / 2.0;
                    return pFace.Evaluate(centerUV);
                }
                else if (geomObj is Face gFace)
                {
                    BoundingBoxUV bboxUV = gFace.GetBoundingBox();
                    UV centerUV = (bboxUV.Min + bboxUV.Max) / 2.0;
                    return gFace.Evaluate(centerUV);
                }
            }
            catch {}

            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve != null && locCurve.Curve != null)
            {
                try { return locCurve.Curve.Evaluate(0.5, true); } catch { }
            }
            return null;
        }

        public static XYZ GetWallDirection(Wall wall)
        {
            LocationCurve lc = wall.Location as LocationCurve;
            if (lc != null && lc.Curve != null)
            {
                Curve curve = lc.Curve;
                if (curve.IsBound)
                {
                    XYZ p0 = curve.GetEndPoint(0);
                    XYZ p1 = curve.GetEndPoint(1);
                    if (p0.IsAlmostEqualTo(p1))
                    {
                        try { return curve.ComputeDerivatives(0.5, true).BasisX.Normalize(); } catch { }
                    }
                    else
                    {
                        XYZ dir = (p1 - p0);
                        return dir.Normalize();
                    }
                }
                try { return curve.ComputeDerivatives(0.5, true).BasisX.Normalize(); } catch { }
            }
            BoundingBoxXYZ bb = wall.get_BoundingBox(null);
            if (bb != null)
            {
                XYZ extent = bb.Max - bb.Min;
                if (Math.Abs(extent.X) > Math.Abs(extent.Y) && Math.Abs(extent.X) > Math.Abs(extent.Z)) return (extent.X > 0 ? XYZ.BasisX : -XYZ.BasisX);
                if (Math.Abs(extent.Y) > Math.Abs(extent.X) && Math.Abs(extent.Y) > Math.Abs(extent.Z)) return (extent.Y > 0 ? XYZ.BasisY : -XYZ.BasisY);
                if (Math.Abs(extent.Z) > 0) return (extent.Z > 0 ? XYZ.BasisZ : -XYZ.BasisZ);
            }
            return XYZ.BasisX;
        }
    }
}
