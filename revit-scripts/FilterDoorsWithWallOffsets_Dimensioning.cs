using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

namespace FilterDoorsWithWallOffsets
{
    public static class Dimensioning
    {
        /// <summary>
        /// Calculates distances between door and adjacent walls by creating temporary dimensions.
        /// Must be called within a transaction.
        /// </summary>
        public static List<DimensionInfo> CalculateDistancesForDoor(
            Document doc,
            UIDocument uidoc,
            FamilyInstance door,
            List<AdjacentWallInfo> adjacentWalls)
        {
            List<DimensionInfo> calculatedDistances = new List<DimensionInfo>();

            Reference doorLeft = door.GetReferences(FamilyInstanceReferenceType.Left)?.FirstOrDefault();
            Reference doorRight = door.GetReferences(FamilyInstanceReferenceType.Right)?.FirstOrDefault();

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
                    return calculatedDistances;
                }
            }

            Wall hostWall = door.Host as Wall;
            if (hostWall == null) return calculatedDistances;

            // Get door orientation data which includes DoorFacing
            DoorOrientation doorOrientation = WallFinding.GetDoorOrientation(door, hostWall);
            // Get host wall direction for orienting the dimension line itself (parallel to host wall)
            XYZ hostWallDir = doorOrientation.WallDirection;

            // This is the primary direction for offsetting the dimension line "in front" of the door.
            XYZ primaryOffsetDirection = doorOrientation.DoorFacing.Normalize();

            foreach (var wallInfo in adjacentWalls)
            {
                try
                {
                    Reference wallRef = wallInfo.ClosestFaceReference;
                    if (wallRef == null)
                    {
                        continue;
                    }

                    XYZ adjWallDir = WallFinding.GetWallDirection(wallInfo.Wall);
                    if (!AreRoughlyPerpendicular(hostWallDir, adjWallDir)) continue;

                    foreach (var sideInfo in wallInfo.WallSides) // sideInfo is WallSide.Front or WallSide.Back
                    {
                        Reference doorRefToUse = null;
                        XYZ doorRefPointForDimLine = XYZ.Zero; // Point on door for dimension line calculation
                        string orientationLabel = "";

                        switch (wallInfo.Position)
                        {
                            case WallPosition.Left:
                                doorRefToUse = doorLeft;
                                doorRefPointForDimLine = doorOrientation.LeftEdge;
                                orientationLabel = sideInfo == WallSide.Front ? "Offset Left" : "Offset Left Reverse";
                                break;
                            case WallPosition.Right:
                                doorRefToUse = doorRight;
                                doorRefPointForDimLine = doorOrientation.RightEdge;
                                orientationLabel = sideInfo == WallSide.Front ? "Offset Right" : "Offset Right Reverse";
                                break;
                            case WallPosition.Front:
                                XYZ wallActualPoint = GetReferencePoint(doc, wallRef);
                                if (wallActualPoint == null) continue;

                                double distToDoorLeftEdge = doorOrientation.LeftEdge.DistanceTo(wallActualPoint);
                                double distToDoorRightEdge = doorOrientation.RightEdge.DistanceTo(wallActualPoint);

                                if (distToDoorLeftEdge < distToDoorRightEdge)
                                {
                                    doorRefToUse = doorLeft;
                                    doorRefPointForDimLine = doorOrientation.LeftEdge;
                                    orientationLabel = sideInfo == WallSide.Front ? "Offset Left" : "Offset Left Reverse";
                                }
                                else
                                {
                                    doorRefToUse = doorRight;
                                    doorRefPointForDimLine = doorOrientation.RightEdge;
                                    orientationLabel = sideInfo == WallSide.Front ? "Offset Right" : "Offset Right Reverse";
                                }
                                break;
                        }

                        if (doorRefToUse == null) continue;

                        XYZ wallRefPointForDimLine = GetReferencePoint(doc, wallRef);
                        if (wallRefPointForDimLine == null) continue;

                        var refArray = new ReferenceArray();
                        refArray.Append(doorRefToUse);
                        refArray.Append(wallRef);

                        XYZ midPointBetweenRefs = (doorRefPointForDimLine + wallRefPointForDimLine) * 0.5;
                        double offsetMagnitude = 3.0; // Desired offset magnitude for the dimension line (e.g., 3 feet)

                        XYZ dimensionLineOffsetVector;
                        if (sideInfo == WallSide.Front)
                        {
                            // For "Front" dimensions (e.g., "Offset Left"), place dim line in primary offset direction
                            dimensionLineOffsetVector = primaryOffsetDirection * offsetMagnitude;
                        }
                        else // sideInfo == WallSide.Back, corresponding to "Reverse" labels
                        {
                            // For "Back" dimensions (e.g., "Offset Left Reverse"), place dim line in opposite direction
                            dimensionLineOffsetVector = -primaryOffsetDirection * offsetMagnitude;
                        }

                        XYZ dimensionLineMidPoint = midPointBetweenRefs + dimensionLineOffsetVector;

                        // Dimension line extends parallel to the host wall
                        // Ensure hostWallDir is normalized for consistent extension length
                        XYZ normalizedHostWallDir = hostWallDir.IsUnitLength() ? hostWallDir : hostWallDir.Normalize();
                        double dimLineHalfLength = 5.0; // Extend 5 feet in each direction from midpoint

                        Line dimLine = Line.CreateBound(
                            dimensionLineMidPoint - normalizedHostWallDir * dimLineHalfLength,
                            dimensionLineMidPoint + normalizedHostWallDir * dimLineHalfLength);

                        // Create a temporary dimension to get the accurate value
                        double dimensionValue = 0.0;
                        try
                        {
                            if (uidoc.ActiveView != null)
                            {
                                Dimension tempDim = doc.Create.NewDimension(uidoc.ActiveView, dimLine, refArray);
                                if (tempDim != null)
                                {
                                    if (tempDim.Value.HasValue)
                                    {
                                        dimensionValue = tempDim.Value.Value;
                                    }
                                    else if (tempDim.Segments.Size > 0 && (tempDim.Segments.get_Item(0) is DimensionSegment segment) && segment.Value.HasValue)
                                    {
                                        dimensionValue = segment.Value.Value;
                                    }
                                    // Delete the temporary dimension
                                    doc.Delete(tempDim.Id);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error creating temp dimension: {ex.Message}");
                            continue;
                        }

                        if (dimensionValue > 0)
                        {
                            calculatedDistances.Add(new DimensionInfo
                            {
                                Dimension = null, // No dimension element created yet
                                Value = dimensionValue,
                                OrientationLabel = orientationLabel,
                                WallId = wallInfo.WallId,
                                IsInFront = sideInfo == WallSide.Front,
                                RequiresBothSides = wallInfo.RequiresBothSides,
                                DoorReference = doorRefToUse,
                                WallReference = wallRef,
                                DoorPoint = doorRefPointForDimLine,
                                WallPoint = wallRefPointForDimLine,
                                HostWallDirection = hostWallDir,
                                DoorFacing = doorOrientation.DoorFacing
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error processing wall {wallInfo.WallId} for door {door.Id}: {ex.Message}");
                    continue;
                }
            }
            return calculatedDistances;
        }

        /// <summary>
        /// Creates dimension elements based on pre-calculated distances
        /// </summary>
        public static void CreateDimensionsFromCalculatedDistances(
            Document doc,
            UIDocument uidoc,
            List<DimensionInfo> calculatedDistances)
        {
            foreach (var dimInfo in calculatedDistances)
            {
                try
                {
                    if (dimInfo.DoorReference == null || dimInfo.WallReference == null)
                        continue;

                    var refArray = new ReferenceArray();
                    refArray.Append(dimInfo.DoorReference);
                    refArray.Append(dimInfo.WallReference);

                    XYZ midPointBetweenRefs = (dimInfo.DoorPoint + dimInfo.WallPoint) * 0.5;
                    double offsetMagnitude = 3.0; // Desired offset magnitude for the dimension line (e.g., 3 feet)

                    XYZ primaryOffsetDirection = dimInfo.DoorFacing.Normalize();
                    XYZ dimensionLineOffsetVector;
                    
                    if (dimInfo.IsInFront)
                    {
                        // For "Front" dimensions (e.g., "Offset Left"), place dim line in primary offset direction
                        dimensionLineOffsetVector = primaryOffsetDirection * offsetMagnitude;
                    }
                    else // Back dimensions
                    {
                        // For "Back" dimensions (e.g., "Offset Left Reverse"), place dim line in opposite direction
                        dimensionLineOffsetVector = -primaryOffsetDirection * offsetMagnitude;
                    }

                    XYZ dimensionLineMidPoint = midPointBetweenRefs + dimensionLineOffsetVector;

                    // Dimension line extends parallel to the host wall
                    // Ensure hostWallDir is normalized for consistent extension length
                    XYZ normalizedHostWallDir = dimInfo.HostWallDirection.IsUnitLength() ? dimInfo.HostWallDirection : dimInfo.HostWallDirection.Normalize();
                    double dimLineHalfLength = 5.0; // Extend 5 feet in each direction from midpoint

                    Line dimLine = Line.CreateBound(
                        dimensionLineMidPoint - normalizedHostWallDir * dimLineHalfLength,
                        dimensionLineMidPoint + normalizedHostWallDir * dimLineHalfLength);

                    Dimension dimension = null;
                    try
                    {
                        if (uidoc.ActiveView == null) continue; // Should not happen in a command
                        dimension = doc.Create.NewDimension(uidoc.ActiveView, dimLine, refArray);
                        dimInfo.Dimension = dimension; // Store the created dimension
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Dimension creation failed for {dimInfo.OrientationLabel}: {ex.Message}");
                        continue;
                    }
                    catch (Exception exAll)
                    {
                        System.Diagnostics.Debug.WriteLine($"Generic error creating dimension for {dimInfo.OrientationLabel}: {exAll.Message}");
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating dimension: {ex.Message}");
                    continue;
                }
            }
        }

        [Obsolete("Use CalculateDistancesForDoor and CreateDimensionsFromCalculatedDistances instead. Note: Must be called within a transaction.")]
        public static List<DimensionInfo> CreateDimensionsForDoor(
            Document doc,
            UIDocument uidoc,
            FamilyInstance door,
            List<AdjacentWallInfo> adjacentWalls)
        {
            // Calculate distances first
            var distances = CalculateDistancesForDoor(doc, uidoc, door, adjacentWalls);
            
            // Then create dimensions
            CreateDimensionsFromCalculatedDistances(doc, uidoc, distances);
            
            return distances;
        }

        private static XYZ GetReferencePoint(Document doc, Reference reference)
        {
            try
            {
                Element referencedElement = doc.GetElement(reference.ElementId);
                if (referencedElement == null) return null;

                GeometryObject go = referencedElement.GetGeometryObjectFromReference(reference);
                if (go == null) return null;

                if (go is Curve c)
                {
                    // For unbound curves, taking midpoint of segment might be problematic if not truly representative.
                    // Using an endpoint or a known parameter might be safer if curve isn't simple.
                    return c.IsBound ? c.Evaluate(0.5, true) : null; // Avoid unbound for now
                }
                else if (go is PlanarFace f)
                {
                    BoundingBoxUV bboxUV = f.GetBoundingBox();
                    if (bboxUV == null || bboxUV.Min == null || bboxUV.Max == null) return f.Origin; // Fallback to origin
                    UV centerUV = (bboxUV.Min + bboxUV.Max) / 2.0;
                    return f.Evaluate(centerUV);
                }
                else if (go is Solid s && !s.Faces.IsEmpty)
                {
                    // Fallback for solid: center of the first face's bounding box (very approximate)
                    Face firstFace = s.Faces.get_Item(0);
                    BoundingBoxUV bboxUV = firstFace.GetBoundingBox();
                    if (bboxUV == null || bboxUV.Min == null || bboxUV.Max == null) return null;
                    UV centerUV = (bboxUV.Min + bboxUV.Max) / 2.0;
                    return firstFace.Evaluate(centerUV);
                }
                else if (go is Edge e)
                {
                    return e.AsCurve().Evaluate(0.5, true);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in GetReferencePoint: {ex.Message}");
            }
            return null;
        }

        private static bool AreRoughlyPerpendicular(XYZ dir1, XYZ dir2)
        {
            if (dir1 == null || dir2 == null || dir1.IsZeroLength() || dir2.IsZeroLength()) return false;
            // Ensure vectors are normalized for consistent dot product interpretation
            XYZ normDir1 = dir1.IsUnitLength() ? dir1 : dir1.Normalize();
            XYZ normDir2 = dir2.IsUnitLength() ? dir2 : dir2.Normalize();
            double dot = Math.Abs(normDir1.DotProduct(normDir2));
            return dot < 0.3;
        }
    }
}
