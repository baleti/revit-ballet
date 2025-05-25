using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace FilterDoorsWallOffsets
{
    public static class WallFinding
    {
        // Static formatting helpers are no longer needed here if not used internally

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

        public static List<AdjacentWallInfo> FindAdjacentWallsParallel(
            FamilyInstance doorInst,
            Wall hostWall,
            DoorOrientation orientation,
            List<WallData> wallsToCheck,
            Document doc) // Removed DoorProcessingResult currentDoorResult parameter
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

            foreach (WallData wallData in wallsToCheck)
            {
                if (wallData.WallId == hostWall.Id)
                    continue;

                if (wallData.BoundingBox != null)
                {
                    bool intersectsBB = !(wallData.BoundingBox.Max.X < min.X || wallData.BoundingBox.Min.X > max.X ||
                                        wallData.BoundingBox.Max.Y < min.Y || wallData.BoundingBox.Min.Y > max.Y ||
                                        wallData.BoundingBox.Max.Z < min.Z || wallData.BoundingBox.Min.Z > max.Z);
                    if (!intersectsBB)
                        continue;
                }

                XYZ wallDir = GetWallDirection(wallData.Wall);
                double perpendicularityDotProduct = Math.Abs(wallDir.DotProduct(orientation.WallDirection));
                bool isPerpendicularToHost = perpendicularityDotProduct < 0.3;

                if (!isPerpendicularToHost)
                    continue;

                bool isInFrontAndProblematic = CheckIfWallIsInFrontOfDoor(wallData, orientation);
                if (isInFrontAndProblematic)
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
            // No detailed string preparation
            return finalIsInFrontDecision;
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
