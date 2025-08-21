using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RevitCommands
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class ShowTransforms : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                // Get UI document and document
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                
                // Get selected elements
                Selection selection = uidoc.Selection;
                ICollection<ElementId> selectedIds = selection.GetElementIds();
                
                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("Show Transforms", 
                        "No elements selected. Please select one or more elements.");
                    return Result.Cancelled;
                }
                
                // Build transformation information
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"Transformation data for {selectedIds.Count} selected element(s):");
                sb.AppendLine();
                
                int elementIndex = 1;
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem == null) continue;
                    
                    sb.AppendLine($"--- Element {elementIndex} ---");
                    sb.AppendLine($"ID: {elem.Id.IntegerValue}");
                    sb.AppendLine($"Name: {elem.Name}");
                    sb.AppendLine($"Category: {elem.Category?.Name ?? "None"}");
                    sb.AppendLine($"Type: {elem.GetType().Name}");
                    
                    // Get location information
                    Location loc = elem.Location;
                    if (loc != null)
                    {
                        if (loc is LocationPoint locPoint)
                        {
                            XYZ point = locPoint.Point;
                            double rotation = locPoint.Rotation;
                            
                            sb.AppendLine("Location Type: Point");
                            sb.AppendLine($"  Position: ({point.X:F3}, {point.Y:F3}, {point.Z:F3})");
                            sb.AppendLine($"  Rotation: {rotation:F3} radians ({RadiansToDegrees(rotation):F2}°)");
                        }
                        else if (loc is LocationCurve locCurve)
                        {
                            Curve curve = locCurve.Curve;
                            XYZ startPoint = curve.GetEndPoint(0);
                            XYZ endPoint = curve.GetEndPoint(1);
                            
                            sb.AppendLine("Location Type: Curve");
                            sb.AppendLine($"  Start: ({startPoint.X:F3}, {startPoint.Y:F3}, {startPoint.Z:F3})");
                            sb.AppendLine($"  End: ({endPoint.X:F3}, {endPoint.Y:F3}, {endPoint.Z:F3})");
                            sb.AppendLine($"  Length: {curve.Length:F3}");
                        }
                    }
                    else
                    {
                        sb.AppendLine("Location Type: None");
                    }
                    
                    // For family instances, get transform
                    if (elem is FamilyInstance famInst)
                    {
                        Transform transform = famInst.GetTransform();
                        sb.AppendLine("Transform Matrix:");
                        sb.AppendLine($"  Origin: ({transform.Origin.X:F3}, {transform.Origin.Y:F3}, {transform.Origin.Z:F3})");
                        sb.AppendLine($"  BasisX: ({transform.BasisX.X:F3}, {transform.BasisX.Y:F3}, {transform.BasisX.Z:F3})");
                        sb.AppendLine($"  BasisY: ({transform.BasisY.X:F3}, {transform.BasisY.Y:F3}, {transform.BasisY.Z:F3})");
                        sb.AppendLine($"  BasisZ: ({transform.BasisZ.X:F3}, {transform.BasisZ.Y:F3}, {transform.BasisZ.Z:F3})");
                        
                        // Check if mirrored by calculating determinant
                        double determinant = transform.Determinant;
                        bool isMirrored = IsTransformMirrored(transform);
                        sb.AppendLine($"  Determinant: {determinant:F3} (Mirrored: {isMirrored})");
                        
                        // Get facing orientation if available
                        if (famInst.CanFlipFacing)
                        {
                            sb.AppendLine($"  Facing Flipped: {famInst.FacingFlipped}");
                        }
                        if (famInst.CanFlipHand)
                        {
                            sb.AppendLine($"  Hand Flipped: {famInst.HandFlipped}");
                        }
                    }
                    
                    // Get bounding box
                    BoundingBoxXYZ bbox = elem.get_BoundingBox(doc.ActiveView);
                    if (bbox != null)
                    {
                        XYZ min = bbox.Min;
                        XYZ max = bbox.Max;
                        XYZ center = (min + max) * 0.5;
                        XYZ size = max - min;
                        
                        sb.AppendLine("Bounding Box:");
                        sb.AppendLine($"  Min: ({min.X:F3}, {min.Y:F3}, {min.Z:F3})");
                        sb.AppendLine($"  Max: ({max.X:F3}, {max.Y:F3}, {max.Z:F3})");
                        sb.AppendLine($"  Center: ({center.X:F3}, {center.Y:F3}, {center.Z:F3})");
                        sb.AppendLine($"  Size: ({size.X:F3}, {size.Y:F3}, {size.Z:F3})");
                    }
                    
                    // For elements with level
                    Parameter levelParam = elem.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) 
                        ?? elem.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
                        ?? elem.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                        
                    if (levelParam != null && levelParam.HasValue)
                    {
                        ElementId levelId = levelParam.AsElementId();
                        if (levelId != ElementId.InvalidElementId)
                        {
                            Level level = doc.GetElement(levelId) as Level;
                            if (level != null)
                            {
                                sb.AppendLine($"Level: {level.Name} (Elevation: {level.Elevation:F3})");
                            }
                        }
                    }
                    
                    // Get offset parameters if available
                    Parameter offsetParam = elem.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                        ?? elem.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                        ?? elem.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                        
                    if (offsetParam != null && offsetParam.HasValue)
                    {
                        double offset = offsetParam.AsDouble();
                        sb.AppendLine($"Offset from Level: {offset:F3}");
                    }
                    
                    sb.AppendLine();
                    elementIndex++;
                }
                
                // Show results in expanded TaskDialog
                TaskDialog dialog = new TaskDialog("Show Transforms")
                {
                    MainContent = sb.ToString(),
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                    CommonButtons = TaskDialogCommonButtons.Ok,
                    DefaultButton = TaskDialogResult.Ok,
                    ExpandedContent = "Note: All coordinates are in project units (feet by default).\n" +
                                     "Rotation values are shown in both radians and degrees.\n" +
                                     "A negative determinant indicates the element is mirrored.",
                    FooterText = $"Total elements analyzed: {selectedIds.Count}"
                };
                
                dialog.Show();
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"An error occurred:\n{ex.Message}");
                return Result.Failed;
            }
        }
        
        /// <summary>
        /// Convert radians to degrees
        /// </summary>
        private double RadiansToDegrees(double radians)
        {
            return radians * (180.0 / Math.PI);
        }
        
        /// <summary>
        /// Check if transform represents a mirrored state
        /// A negative determinant indicates mirroring
        /// </summary>
        private bool IsTransformMirrored(Transform transform)
        {
            // Calculate determinant - negative means mirrored
            double det = transform.Determinant;
            return det < 0;
        }
    }
}
