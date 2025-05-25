using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace FilterDoorsWithWallOffsets
{
    public static class Diagnostics
    {
        public static void ShowDiagnosticsDialog(List<WallDiagnosticInfo> diagnostics, FamilyInstance door)
        {
            if (diagnostics == null || !diagnostics.Any())
            {
                TaskDialog.Show("Diagnostics", $"No walls found within search area for door {door.Id}.");
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"WALL PROCESSING DIAGNOSTICS FOR DOOR {door.Id}");
            sb.AppendLine($"Door Location: {FormatXYZ(GetDoorLocation(door))}");
            sb.AppendLine();

            // Group by processing result
            var accepted = diagnostics.Where(d => d.ProcessingResult == WallProcessingResult.Accepted).ToList();
            var rejected = diagnostics.Where(d => d.ProcessingResult != WallProcessingResult.Accepted).ToList();

            sb.AppendLine($"SUMMARY:");
            sb.AppendLine($"  Total walls in search area: {diagnostics.Count}");
            sb.AppendLine($"  Accepted walls: {accepted.Count}");
            sb.AppendLine($"  Rejected walls: {rejected.Count}");
            sb.AppendLine();

            if (accepted.Any())
            {
                sb.AppendLine("✓ ACCEPTED WALLS:");
                foreach (var wall in accepted.OrderBy(w => w.WallId.IntegerValue))
                {
                    sb.AppendLine($"  Wall {wall.WallId} - {wall.WallName}");
                    sb.AppendLine($"    Start: {FormatXYZ(wall.WallStartPoint)} End: {FormatXYZ(wall.WallEndPoint)}");
                    sb.AppendLine($"    Distance to door: {wall.DistanceToDoor:F3} ft");
                    sb.AppendLine($"    Position: {wall.Position}, Sides: [{string.Join(", ", wall.WallSides)}]");
                    sb.AppendLine();
                }
            }

            if (rejected.Any())
            {
                sb.AppendLine("✗ REJECTED WALLS:");
                
                // Group by rejection reason
                var rejectionGroups = rejected.GroupBy(w => w.ProcessingResult);
                
                foreach (var group in rejectionGroups.OrderBy(g => g.Key.ToString()))
                {
                    sb.AppendLine($"  Reason: {GetRejectionReasonDescription(group.Key)}");
                    
                    foreach (var wall in group.OrderBy(w => w.WallId.IntegerValue))
                    {
                        sb.AppendLine($"    Wall {wall.WallId} - {wall.WallName}");
                        sb.AppendLine($"      Start: {FormatXYZ(wall.WallStartPoint)} End: {FormatXYZ(wall.WallEndPoint)}");
                        sb.AppendLine($"      Distance to door: {wall.DistanceToDoor:F3} ft");
                        
                        if (!string.IsNullOrEmpty(wall.RejectionDetails))
                        {
                            sb.AppendLine($"      Details: {wall.RejectionDetails}");
                        }
                    }
                    sb.AppendLine();
                }
            }

            // Show the dialog with the diagnostic information
            TaskDialog dialog = new TaskDialog("Wall Processing Diagnostics");
            dialog.MainContent = sb.ToString();
            dialog.CommonButtons = TaskDialogCommonButtons.Ok;
            dialog.DefaultButton = TaskDialogResult.Ok;
            dialog.Show();
        }

        private static string GetRejectionReasonDescription(WallProcessingResult result)
        {
            switch (result)
            {
                case WallProcessingResult.RejectedNotPerpendicular:
                    return "Not perpendicular to host wall";
                case WallProcessingResult.RejectedEndpointsTooFar:
                    return "Wall endpoints too far from door (>500mm)";
                case WallProcessingResult.RejectedDirectlyInFrontOrBehind:
                    return "Wall directly in front of or behind door";
                case WallProcessingResult.RejectedOccluded:
                    return "Wall occluded by other walls";
                case WallProcessingResult.RejectedTooFarFromEdge:
                    return "Wall too far from door edge";
                case WallProcessingResult.RejectedNoValidSides:
                    return "No valid wall sides determined";
                case WallProcessingResult.RejectedNoFaceReference:
                    return "Could not get wall face reference";
                default:
                    return result.ToString();
            }
        }

        private static string FormatXYZ(XYZ point)
        {
            if (point == null) return "null";
            return $"({point.X:F2}, {point.Y:F2}, {point.Z:F2})";
        }

        private static XYZ GetDoorLocation(FamilyInstance door)
        {
            if (door?.Location is LocationPoint locationPoint)
            {
                return locationPoint.Point;
            }
            return null;
        }

        public static WallDiagnosticInfo CreateWallDiagnosticInfo(WallData wallData, FamilyInstance door)
        {
            var info = new WallDiagnosticInfo
            {
                WallId = wallData.WallId,
                WallName = wallData.Wall.Name ?? "Unnamed Wall",
                WallStartPoint = wallData.Curve.GetEndPoint(0),
                WallEndPoint = wallData.Curve.GetEndPoint(1),
                ProcessingResult = WallProcessingResult.InProgress
            };

            // Calculate distance to door
            XYZ doorLocation = GetDoorLocation(door);
            if (doorLocation != null)
            {
                info.DistanceToDoor = wallData.Curve.Distance(doorLocation);
            }

            return info;
        }
    }

    public class WallDiagnosticInfo
    {
        public ElementId WallId { get; set; }
        public string WallName { get; set; }
        public XYZ WallStartPoint { get; set; }
        public XYZ WallEndPoint { get; set; }
        public double DistanceToDoor { get; set; }
        public WallProcessingResult ProcessingResult { get; set; }
        public string RejectionDetails { get; set; }
        public WallPosition Position { get; set; }
        public List<WallSide> WallSides { get; set; } = new List<WallSide>();

        public void SetRejected(WallProcessingResult reason, string details = null)
        {
            ProcessingResult = reason;
            RejectionDetails = details;
        }

        public void SetAccepted(WallPosition position, List<WallSide> sides)
        {
            ProcessingResult = WallProcessingResult.Accepted;
            Position = position;
            WallSides = sides?.ToList() ?? new List<WallSide>();
        }
    }

    public enum WallProcessingResult
    {
        InProgress,
        Accepted,
        RejectedNotPerpendicular,
        RejectedEndpointsTooFar,
        RejectedDirectlyInFrontOrBehind,
        RejectedOccluded,
        RejectedTooFarFromEdge,
        RejectedNoValidSides,
        RejectedNoFaceReference
    }
}