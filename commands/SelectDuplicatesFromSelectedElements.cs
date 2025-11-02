using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectDuplicatesFromSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get selected elements
                ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
                
                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("No Selection", "Please select elements before running this command.");
                    return Result.Cancelled;
                }

                // Find duplicates
                ICollection<ElementId> duplicateIds = FindDuplicates(doc, selectedIds);
                
                if (duplicateIds.Count > 0)
                {
                    // Select the duplicate elements
                    uidoc.SetSelectionIds(duplicateIds);
                }
                else
                {
                    TaskDialog.Show("No Duplicates", 
                        "No duplicate elements found in the selection.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private ICollection<ElementId> FindDuplicates(Document doc, ICollection<ElementId> elementIds)
        {
            List<ElementId> duplicates = new List<ElementId>();

            // Group elements by their duplicate signature
            Dictionary<string, List<Element>> duplicateGroups = new Dictionary<string, List<Element>>();
            
            // Debug info
            StringBuilder debugInfo = new StringBuilder();
            debugInfo.AppendLine("=== Duplicate Detection Debug Info ===");

            foreach (ElementId id in elementIds)
            {
                Element elem = doc.GetElement(id);
                if (elem == null) continue;

                // Generate a unique signature for the element
                string signature = GetElementSignature(elem);
                
                // Debug: log signatures for sprinklers
                if (elem.Category != null && elem.Category.Name == "Sprinklers")
                {
                    debugInfo.AppendLine($"Element {elem.Id}: {signature}");
                }
                
                if (!string.IsNullOrEmpty(signature))
                {
                    if (!duplicateGroups.ContainsKey(signature))
                    {
                        duplicateGroups[signature] = new List<Element>();
                    }
                    duplicateGroups[signature].Add(elem);
                }
            }

            // Show debug info for sprinklers if no duplicates found
            bool foundDuplicates = false;
            
            // For each group with duplicates, add all duplicates (except the first) to the list
            foreach (var group in duplicateGroups.Values)
            {
                if (group.Count > 1)
                {
                    foundDuplicates = true;
                    // Skip the first element (original), add rest as duplicates
                    for (int i = 1; i < group.Count; i++)
                    {
                        duplicates.Add(group[i].Id);
                    }
                }
            }
            
            // If no duplicates found and we have sprinklers, show debug info
            if (!foundDuplicates && debugInfo.Length > 50)
            {
                TaskDialog.Show("Debug Info", debugInfo.ToString());
            }

            return duplicates;
        }

        private string GetElementSignature(Element elem)
        {
            if (elem == null) return string.Empty;

            StringBuilder signature = new StringBuilder();
            
            // Add category
            Category category = elem.Category;
            if (category != null)
            {
                signature.Append($"CAT:{category.Id.IntegerValue}|");
            }

            // Add element type
            ElementId typeId = elem.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                signature.Append($"TYPE:{typeId.IntegerValue}|");
            }

            // Add class type
            signature.Append($"CLASS:{elem.GetType().FullName}|");

            // Handle different element types
            if (elem is FamilyInstance fi)
            {
                signature.Append(GetFamilyInstanceSignature(fi));
            }
            else if (elem is Wall wall)
            {
                signature.Append(GetWallSignature(wall));
            }
            else if (elem is Floor floor)
            {
                signature.Append(GetFloorSignature(floor));
            }
            else if (elem is Room room)
            {
                signature.Append(GetRoomSignature(room));
            }
            else if (elem is TextNote textNote)
            {
                signature.Append(GetTextNoteSignature(textNote));
            }
            else if (elem is Dimension dimension)
            {
                signature.Append(GetDimensionSignature(dimension));
            }
            else if (elem is DetailLine detailLine)
            {
                signature.Append(GetDetailLineSignature(detailLine));
            }
            else if (elem is ModelLine modelLine)
            {
                signature.Append(GetModelLineSignature(modelLine));
            }
            else
            {
                // For other element types, use parameter-based comparison
                signature.Append(GetParameterSignature(elem));
            }

            return signature.ToString();
        }

        private string GetFamilyInstanceSignature(FamilyInstance fi)
        {
            StringBuilder sig = new StringBuilder();
            
            // Family and symbol
            sig.Append($"FAM:{fi.Symbol.Family.Name}|");
            sig.Append($"SYM:{fi.Symbol.Name}|");
            
            // Location - using less precision for better duplicate detection
            LocationPoint locPoint = fi.Location as LocationPoint;
            if (locPoint != null)
            {
                XYZ point = locPoint.Point;
                // Round to 2 decimal places (about 0.3mm precision in mm units)
                sig.Append($"LOC:{Math.Round(point.X, 2)},{Math.Round(point.Y, 2)},{Math.Round(point.Z, 2)}|");
                
                // Normalize rotation to 0-2PI range and round
                double rotation = locPoint.Rotation;
                while (rotation < 0) rotation += 2 * Math.PI;
                while (rotation >= 2 * Math.PI) rotation -= 2 * Math.PI;
                sig.Append($"ROT:{Math.Round(rotation, 2)}|");
            }
            
            // Level
            Parameter levelParam = fi.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
            if (levelParam == null)
                levelParam = fi.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
            if (levelParam == null)
                levelParam = fi.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
            
            if (levelParam != null && levelParam.HasValue)
            {
                sig.Append($"LEVEL:{levelParam.AsElementId().IntegerValue}|");
            }
            
            // Host
            if (fi.Host != null)
            {
                sig.Append($"HOST:{fi.Host.Id.IntegerValue}|");
            }
            
            // Facing and hand flipped for doors/windows
            if (fi.CanFlipFacing)
            {
                sig.Append($"FACE:{fi.FacingFlipped}|");
            }
            if (fi.CanFlipHand)
            {
                sig.Append($"HAND:{fi.HandFlipped}|");
            }
            
            // MEP specific parameters
            if (fi.Symbol.Family.FamilyCategory != null)
            {
                string catName = fi.Symbol.Family.FamilyCategory.Name;
                if (catName.Contains("Sprinkler") || catName.Contains("Mechanical") || 
                    catName.Contains("Electrical") || catName.Contains("Plumbing"))
                {
                    
                    // Check offset
                    Parameter offset = fi.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
                    if (offset == null)
                        offset = fi.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                    if (offset != null && offset.HasValue)
                    {
                        sig.Append($"OFFSET:{Math.Round(offset.AsDouble(), 2)}|");
                    }
                }
            }
            
            // For more precise duplicate detection, don't include variable parameters
            // like Mark, Comments, etc. unless specifically needed
            
            return sig.ToString();
        }

        private string GetWallSignature(Wall wall)
        {
            StringBuilder sig = new StringBuilder();
            
            // Wall type
            WallType wallType = wall.WallType;
            sig.Append($"WTYPE:{wallType.Name}|");
            
            // Location curve
            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve != null && locCurve.Curve != null)
            {
                Curve curve = locCurve.Curve;
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                sig.Append($"START:{Math.Round(start.X, 4)},{Math.Round(start.Y, 4)},{Math.Round(start.Z, 4)}|");
                sig.Append($"END:{Math.Round(end.X, 4)},{Math.Round(end.Y, 4)},{Math.Round(end.Z, 4)}|");
            }
            
            // Height
            Parameter heightParam = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
            if (heightParam != null && heightParam.HasValue)
            {
                sig.Append($"HEIGHT:{Math.Round(heightParam.AsDouble(), 4)}|");
            }
            
            return sig.ToString();
        }

        private string GetFloorSignature(Floor floor)
        {
            StringBuilder sig = new StringBuilder();
            
            // Floor type
            FloorType floorType = floor.FloorType;
            sig.Append($"FTYPE:{floorType.Name}|");
            
            // Level
            Parameter levelParam = floor.get_Parameter(BuiltInParameter.LEVEL_PARAM);
            if (levelParam != null && levelParam.HasValue)
            {
                sig.Append($"LEVEL:{levelParam.AsElementId().IntegerValue}|");
            }
            
            // Get geometry to calculate area/perimeter
            Options geomOptions = new Options();
            geomOptions.ComputeReferences = false;
            geomOptions.DetailLevel = ViewDetailLevel.Coarse;
            
            GeometryElement geomElement = floor.get_Geometry(geomOptions);
            if (geomElement != null)
            {
                double totalArea = 0;
                foreach (GeometryObject geomObj in geomElement)
                {
                    if (geomObj is Solid solid)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            totalArea += face.Area;
                        }
                    }
                }
                sig.Append($"AREA:{Math.Round(totalArea, 4)}|");
            }
            
            // Offset from level
            Parameter offsetParam = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
            if (offsetParam != null && offsetParam.HasValue)
            {
                sig.Append($"OFFSET:{Math.Round(offsetParam.AsDouble(), 4)}|");
            }
            
            return sig.ToString();
        }

        private string GetRoomSignature(Room room)
        {
            StringBuilder sig = new StringBuilder();
            
            // Room name and number
            sig.Append($"NAME:{room.Name}|");
            sig.Append($"NUM:{room.Number}|");
            
            // Level
            sig.Append($"LEVEL:{room.LevelId.IntegerValue}|");
            
            // Area
            sig.Append($"AREA:{Math.Round(room.Area, 4)}|");
            
            return sig.ToString();
        }

        private string GetTextNoteSignature(TextNote textNote)
        {
            StringBuilder sig = new StringBuilder();
            
            // Text content
            sig.Append($"TEXT:{textNote.Text}|");
            
            // Text type
            sig.Append($"TTYPE:{textNote.TextNoteType.Name}|");
            
            // Position
            XYZ coord = textNote.Coord;
            sig.Append($"POS:{Math.Round(coord.X, 4)},{Math.Round(coord.Y, 4)},{Math.Round(coord.Z, 4)}|");
            
            return sig.ToString();
        }

        private string GetDimensionSignature(Dimension dimension)
        {
            StringBuilder sig = new StringBuilder();
            
            // Dimension type
            sig.Append($"DTYPE:{dimension.DimensionType.Name}|");
            
            // References count
            ReferenceArray refs = dimension.References;
            sig.Append($"REFS:{refs.Size}|");
            
            // Dimension line
            Line dimLine = dimension.Curve as Line;
            if (dimLine != null)
            {
                XYZ origin = dimLine.Origin;
                sig.Append($"ORIGIN:{Math.Round(origin.X, 4)},{Math.Round(origin.Y, 4)},{Math.Round(origin.Z, 4)}|");
                XYZ direction = dimLine.Direction;
                sig.Append($"DIR:{Math.Round(direction.X, 4)},{Math.Round(direction.Y, 4)},{Math.Round(direction.Z, 4)}|");
            }
            
            return sig.ToString();
        }

        private string GetDetailLineSignature(DetailLine detailLine)
        {
            StringBuilder sig = new StringBuilder();
            
            // Line style
            GraphicsStyle lineStyle = detailLine.LineStyle as GraphicsStyle;
            if (lineStyle != null)
            {
                sig.Append($"STYLE:{lineStyle.Name}|");
            }
            
            // Geometry
            Curve curve = detailLine.GeometryCurve;
            if (curve != null)
            {
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                sig.Append($"START:{Math.Round(start.X, 4)},{Math.Round(start.Y, 4)},{Math.Round(start.Z, 4)}|");
                sig.Append($"END:{Math.Round(end.X, 4)},{Math.Round(end.Y, 4)},{Math.Round(end.Z, 4)}|");
            }
            
            return sig.ToString();
        }

        private string GetModelLineSignature(ModelLine modelLine)
        {
            StringBuilder sig = new StringBuilder();
            
            // Line style
            GraphicsStyle lineStyle = modelLine.LineStyle as GraphicsStyle;
            if (lineStyle != null)
            {
                sig.Append($"STYLE:{lineStyle.Name}|");
            }
            
            // Sketch plane
            sig.Append($"SKETCH:{modelLine.SketchPlane.Id.IntegerValue}|");
            
            // Geometry
            Curve curve = modelLine.GeometryCurve;
            if (curve != null)
            {
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                sig.Append($"START:{Math.Round(start.X, 4)},{Math.Round(start.Y, 4)},{Math.Round(start.Z, 4)}|");
                sig.Append($"END:{Math.Round(end.X, 4)},{Math.Round(end.Y, 4)},{Math.Round(end.Z, 4)}|");
            }
            
            return sig.ToString();
        }

        private string GetKeyParametersSignature(Element elem)
        {
            StringBuilder sig = new StringBuilder();
            
            // Important instance parameters that define uniqueness
            // Using only parameters that are confirmed to exist in Revit API
            BuiltInParameter[] keyParams = new[]
            {
                BuiltInParameter.ALL_MODEL_MARK,
                BuiltInParameter.DOOR_NUMBER,
                BuiltInParameter.ROOM_NUMBER,
                BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                BuiltInParameter.INSTANCE_LENGTH_PARAM
            };
            
            foreach (BuiltInParameter bip in keyParams)
            {
                try
                {
                    Parameter param = elem.get_Parameter(bip);
                    if (param != null && param.HasValue)
                    {
                        if (param.StorageType == StorageType.String)
                        {
                            string value = param.AsString();
                            if (!string.IsNullOrEmpty(value))
                            {
                                sig.Append($"{bip}:{value}|");
                            }
                        }
                        else if (param.StorageType == StorageType.Double)
                        {
                            sig.Append($"{bip}:{Math.Round(param.AsDouble(), 4)}|");
                        }
                        else if (param.StorageType == StorageType.Integer)
                        {
                            sig.Append($"{bip}:{param.AsInteger()}|");
                        }
                        else if (param.StorageType == StorageType.ElementId)
                        {
                            sig.Append($"{bip}:{param.AsElementId().IntegerValue}|");
                        }
                    }
                }
                catch
                {
                    // Skip parameters that don't exist for this element type
                    continue;
                }
            }
            
            // Also check for common instance parameters by name
            string[] paramNames = new[] { "Width", "Height", "Length", "Depth", "Thickness" };
            foreach (string paramName in paramNames)
            {
                Parameter param = elem.LookupParameter(paramName);
                if (param != null && param.HasValue && param.StorageType == StorageType.Double)
                {
                    sig.Append($"{paramName}:{Math.Round(param.AsDouble(), 4)}|");
                }
            }
            
            return sig.ToString();
        }

        private string GetParameterSignature(Element elem)
        {
            StringBuilder sig = new StringBuilder();
            
            // Get all parameters
            foreach (Parameter param in elem.Parameters)
            {
                if (param.IsReadOnly || !param.HasValue) continue;
                
                // Skip some parameters that shouldn't affect duplicate detection
                if (param.Definition.Name.Contains("Volume") ||
                    param.Definition.Name.Contains("Area") ||
                    param.Definition.Name.Contains("Perimeter"))
                    continue;
                
                if (param.StorageType == StorageType.String)
                {
                    string value = param.AsString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        sig.Append($"{param.Definition.Name}:{value}|");
                    }
                }
                else if (param.StorageType == StorageType.Double)
                {
                    sig.Append($"{param.Definition.Name}:{Math.Round(param.AsDouble(), 4)}|");
                }
                else if (param.StorageType == StorageType.Integer)
                {
                    sig.Append($"{param.Definition.Name}:{param.AsInteger()}|");
                }
                else if (param.StorageType == StorageType.ElementId)
                {
                    ElementId id = param.AsElementId();
                    if (id != ElementId.InvalidElementId)
                    {
                        sig.Append($"{param.Definition.Name}:{id.IntegerValue}|");
                    }
                }
            }
            
            return sig.ToString();
        }
    }
}
