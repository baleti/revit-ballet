using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

namespace FilterDoorsWallOffsets
{
    public static class Dimensioning
    {
        public static List<DimensionInfo> CreateDimensionsForDoor(
            Document doc,
            UIDocument uidoc,
            FamilyInstance door,
            List<AdjacentWallInfo> adjacentWalls)
        {
            List<DimensionInfo> createdDimensions = new List<DimensionInfo>();

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
                    return createdDimensions;
                }
            }

            Wall hostWall = door.Host as Wall;
            if (hostWall == null) return createdDimensions;

            // Get door orientation data which includes DoorFacing
            DoorOrientation doorOrientation = WallFinding.GetDoorOrientation(door, hostWall);
            // Get host wall direction for orienting the dimension line itself (parallel to host wall)
            XYZ hostWallDir = doorOrientation.WallDirection; // Or WallFinding.GetWallDirection(hostWall); they should be consistent

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

                        Dimension dimension = null;
                        try
                        {
                            if (uidoc.ActiveView == null) continue; // Should not happen in a command
                            dimension = doc.Create.NewDimension(uidoc.ActiveView, dimLine, refArray);
                        }
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Dimension creation failed for door {door.Id}, wall {wallInfo.WallId} ({orientationLabel}): {ex.Message}");
                            continue;
                        }
                        catch (Exception exAll) // Catch any other potential errors during dimension creation
                        {
                            System.Diagnostics.Debug.WriteLine($"Generic error creating dimension for door {door.Id}, wall {wallInfo.WallId} ({orientationLabel}): {exAll.Message}");
                            continue;
                        }

                        double dimensionValue = 0.0;
                        if (dimension != null)
                        {
                            if (dimension.Value.HasValue)
                            {
                                dimensionValue = dimension.Value.Value;
                            }
                            else if (dimension.Segments.Size > 0 && (dimension.Segments.get_Item(0) is DimensionSegment segment) && segment.Value.HasValue)
                            {
                                dimensionValue = segment.Value.Value;
                            }
                        }
                        else // If dimension creation failed and dimension is null
                        {
                            // We continued, so this part might not be reached unless we change the continue logic.
                            // For now, if dimension is null, we don't add to createdDimensions.
                            continue;
                        }

                        createdDimensions.Add(new DimensionInfo
                        {
                            Dimension = dimension,
                            Value = dimensionValue,
                            OrientationLabel = orientationLabel,
                            WallId = wallInfo.WallId,
                            IsInFront = sideInfo == WallSide.Front, // This reflects the wall feature's side
                            RequiresBothSides = wallInfo.RequiresBothSides
                        });
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error processing wall {wallInfo.WallId} for door {door.Id}: {ex.Message}");
                    continue;
                }
            }
            return createdDimensions;
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
