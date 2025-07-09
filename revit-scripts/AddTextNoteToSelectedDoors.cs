using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

[Transaction(TransactionMode.Manual)]
public class AddTextNoteToSelectedDoors : IExternalCommand
{
    // Class to track text note placement information
    private class TextNoteInfo
    {
        public TextNote Note { get; set; }
        public XYZ OriginalPosition { get; set; }
        public XYZ CurrentPosition { get; set; }
        public Element SourceElement { get; set; }
        public string TextContent { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public bool WasMoved { get; set; }
    }
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get and filter selected elements
        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
        if (!selectedIds.Any())
        {
            TaskDialog.Show("Warning", "Please select doors or windows before running the command.");
            return Result.Failed;
        }

        List<Element> selectedOpenings = selectedIds
            .Select(id => doc.GetElement(id))
            .Where(e => e?.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_Doors ||
                       e?.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
            .ToList();

        if (!selectedOpenings.Any())
        {
            TaskDialog.Show("Warning", "No doors or windows in current selection.");
            return Result.Failed;
        }

        try
        {
            // Collect all parameters once before processing
            var allParams = selectedOpenings
                .SelectMany(d => d.Parameters.OfType<Parameter>())
                .Select(p => p.Definition.Name)
                .Distinct()
                .OrderBy(n => n)
                .ToList();

            // Prepare parameter selection data
            List<string> coreProperties = new List<string>
            {
                "Element Id", "Category", "Family Name", "Instance Name", "Level", "Mark",
                "Head Height", "Comments", "Group", "FacingFlipped", "HandFlipped", "Wall Direction",
                "Room From", "Room To", "Width", "Height"
            };

            // Combine core properties with all other parameters
            var allProperties = coreProperties.Concat(allParams.Except(coreProperties)).ToList();

            // Create parameter selection data for grid
            var parameterSelectionData = allProperties
                .Select(p => new Dictionary<string, object>
                {
                    ["Parameter Name"] = p,
                    ["Parameter Type"] = coreProperties.Contains(p) ? "Core" : "Additional"
                })
                .ToList();

            // Show parameter selection dialog
            var chosenParameters = CustomGUIs.DataGrid(
                parameterSelectionData,
                new List<string> { "Parameter Name", "Parameter Type" },
                false);

            if (!chosenParameters.Any())
            {
                return Result.Cancelled;
            }

            // Extract selected parameter names
            var selectedParameterNames = chosenParameters
                .Select(dict => dict["Parameter Name"] as string)
                .Where(name => !string.IsNullOrEmpty(name))
                .ToList();

            // Get or create text note type adjusted for current view scale
            ElementId textNoteTypeId = GetOrCreateTextNoteType(doc, doc.ActiveView);
            if (textNoteTypeId == null || textNoteTypeId == ElementId.InvalidElementId)
            {
                TaskDialog.Show("Error", "Failed to create text note type.");
                return Result.Failed;
            }

            // Generate UUID for this batch of text notes
            string uuid = Guid.NewGuid().ToString();
            string commentPrefix = $"door-window parameter note {uuid}";
            
            // Get the actual type name for display
            TextNoteType actualType = doc.GetElement(textNoteTypeId) as TextNoteType;
            string actualTypeName = actualType?.Name ?? "Unknown";

            using (Transaction trans = new Transaction(doc, "Add Text Notes to Doors/Windows"))
            {
                trans.Start();

                try
                {
                    View activeView = doc.ActiveView;
                    int viewScale = activeView.Scale;
                    
                    // Get text note type to retrieve its properties
                    TextNoteType textType = doc.GetElement(textNoteTypeId) as TextNoteType;
                    if (textType == null)
                    {
                        trans.RollBack();
                        return Result.Failed;
                    }
                    
                    // Get text height from the type
                    Parameter textSizeParam = textType.get_Parameter(BuiltInParameter.TEXT_SIZE);
                    double textHeight = textSizeParam?.AsDouble() ?? (3.0 / 32.0 / 12.0 * viewScale);
                    
                    // Estimate character width (typically 60-70% of height for proportional fonts)
                    double avgCharWidth = textHeight * 0.6;
                    
                    // List to track all text note info for collision detection
                    var textNoteInfos = new List<TextNoteInfo>();

                    // First pass: Calculate positions and dimensions without creating notes
                    foreach (Element opening in selectedOpenings)
                    {
                        if (!(opening is FamilyInstance openingInst)) continue;

                        // Build text content from selected parameters
                        var textContent = BuildTextContent(opening, openingInst, selectedParameterNames, allParams, doc);
                        if (string.IsNullOrEmpty(textContent))
                            continue;

                        // Get opening position
                        LocationPoint locPoint = opening.Location as LocationPoint;
                        if (locPoint == null) continue;

                        XYZ position = locPoint.Point;

                        // Calculate text note position with offset
                        BoundingBoxXYZ boundingBox = opening.get_BoundingBox(null);
                        if (boundingBox != null)
                        {
                            // Offset based on text height plus gap
                            double offsetY = textHeight * 3; // More space for readability
                            
                            XYZ textPosition = new XYZ(position.X, boundingBox.Min.Y - offsetY, position.Z);
                            
                            // Calculate text dimensions
                            string[] lines = textContent.Split('\n');
                            int maxLineLength = lines.Max(l => l.Length);
                            double estimatedWidth = maxLineLength * avgCharWidth;
                            double estimatedHeight = lines.Length * textHeight * 1.2; // 1.2 for line spacing
                            
                            var noteInfo = new TextNoteInfo
                            {
                                OriginalPosition = textPosition,
                                CurrentPosition = textPosition,
                                SourceElement = opening,
                                TextContent = textContent,
                                Width = estimatedWidth,
                                Height = estimatedHeight,
                                WasMoved = false
                            };
                            
                            textNoteInfos.Add(noteInfo);
                        }
                    }
                    
                    // Resolve overlaps before creating any text notes
                    ResolveOverlapsBeforeCreation(textNoteInfos, textHeight);
                    
                    // Second pass: Create text notes at their final positions
                    foreach (var noteInfo in textNoteInfos)
                    {
                        // Create text note at final position
                        TextNote textNote = TextNote.Create(
                            doc,
                            activeView.Id,
                            noteInfo.CurrentPosition,
                            noteInfo.TextContent,
                            textNoteTypeId);

                        if (textNote != null)
                        {
                            noteInfo.Note = textNote;
                            
                            // Set comment parameter
                            Parameter commentParam = textNote.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                            if (commentParam != null && !commentParam.IsReadOnly)
                            {
                                commentParam.Set(commentPrefix);
                            }
                            
                            // Create leader line if the note was moved
                            if (noteInfo.WasMoved)
                            {
                                CreateLeaderLine(doc, activeView, noteInfo);
                            }
                        }
                    }

                    trans.Commit();

                    int movedCount = textNoteInfos.Count(n => n.WasMoved);
                    TaskDialog.Show("Success",
                        $"Created text notes for {selectedOpenings.Count} doors/windows.\n" +
                        $"Text Note Type: {actualTypeName}\n" +
                        $"Notes repositioned: {movedCount}\n" +
                        $"UUID: {uuid}");
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    trans.RollBack();
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Unexpected error: {ex.Message}";
            return Result.Failed;
        }
    }

    // NOTE: In Revit API, you cannot create a TextNoteType from scratch.
    // You must duplicate an existing type and modify its properties.
    // This method creates a highly customized type specifically for our use.
    private ElementId GetOrCreateTextNoteType(Document doc, View view)
    {
        // Calculate appropriate text height based on view scale
        // Standard text height is typically 3/32" on paper, adjust for scale
        double paperTextHeight = 3.0 / 32.0 / 12.0; // Convert to feet
        double modelTextHeight = paperTextHeight * view.Scale;
        
        // Round to nearest 1/32" for cleaner values
        modelTextHeight = Math.Round(modelTextHeight * 32 * 12) / (32 * 12);
        
        // Create a unique name for our custom text note type
        // Note: Revit doesn't allow certain characters in type names including : { } [ ] | ; < > ? ` ~
        string typeName = $"Door-Window Parameter Notes - 1-{view.Scale}";
        
        // Check if this type already exists
        var existingType = new FilteredElementCollector(doc)
            .OfClass(typeof(TextNoteType))
            .Cast<TextNoteType>()
            .FirstOrDefault(t => t.Name == typeName);
            
        if (existingType != null)
            return existingType.Id;
        
        // Find the simplest text note type to use as a base
        // Prefer types without leaders or special formatting
        var baseType = new FilteredElementCollector(doc)
            .OfClass(typeof(TextNoteType))
            .Cast<TextNoteType>()
            .Where(t => t.Name != null)
            .OrderBy(t => t.Name.Length) // Prefer shorter names (usually simpler types)
            .FirstOrDefault(t => 
                !t.Name.Contains("Title") && 
                !t.Name.Contains("Large") &&
                !t.Name.Contains("Leader") &&
                !t.Name.Contains("Arrow"));
            
        if (baseType == null)
        {
            // Fallback to any text note type
            baseType = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .FirstElement() as TextNoteType;
        }
        
        if (baseType == null)
            return ElementId.InvalidElementId;
        
        try
        {
            // Create a duplicate name that's unique
            string baseDuplicateName = $"{baseType.Name} - temporary note";
            string duplicateName = baseDuplicateName;
            int counter = 1;
            
            // Check if duplicate name already exists and make it unique
            while (new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .Any(t => t.Name == duplicateName))
            {
                duplicateName = $"{baseDuplicateName} {counter}";
                counter++;
            }
            
            TextNoteType newType = baseType.Duplicate(duplicateName) as TextNoteType;
            
            // If duplication worked, try to rename to our desired name
            if (newType != null)
            {
                try 
                {
                    // Only rename if our desired name doesn't already exist
                    if (!new FilteredElementCollector(doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .Any(t => t.Name == typeName))
                    {
                        newType.Name = typeName;
                    }
                }
                catch
                {
                    // If rename fails, keep the duplicate name
                }
                
                // Configure all available parameters to make this type distinct
                
                // Text size based on scale
                Parameter textSizeParam = newType.get_Parameter(BuiltInParameter.TEXT_SIZE);
                if (textSizeParam != null && !textSizeParam.IsReadOnly)
                {
                    textSizeParam.Set(modelTextHeight);
                }
                
                // Background - opaque for better readability
                Parameter backgroundParam = newType.get_Parameter(BuiltInParameter.TEXT_BACKGROUND);
                if (backgroundParam != null && !backgroundParam.IsReadOnly)
                {
                    backgroundParam.Set(1); // Opaque
                }
                
                // Set text font
                Parameter fontParam = newType.get_Parameter(BuiltInParameter.TEXT_FONT);
                if (fontParam != null && !fontParam.IsReadOnly)
                {
                    fontParam.Set("Arial Narrow"); // Compact font for parameter lists
                }
                
                // Text style - not bold, not italic
                Parameter boldParam = newType.get_Parameter(BuiltInParameter.TEXT_STYLE_BOLD);
                if (boldParam != null && !boldParam.IsReadOnly)
                {
                    boldParam.Set(0);
                }
                
                Parameter italicParam = newType.get_Parameter(BuiltInParameter.TEXT_STYLE_ITALIC);
                if (italicParam != null && !italicParam.IsReadOnly)
                {
                    italicParam.Set(0);
                }
                
                Parameter underlineParam = newType.get_Parameter(BuiltInParameter.TEXT_STYLE_UNDERLINE);
                if (underlineParam != null && !underlineParam.IsReadOnly)
                {
                    underlineParam.Set(0);
                }
                
                // Width factor - slightly condensed for compact display
                Parameter widthParam = newType.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE);
                if (widthParam != null && !widthParam.IsReadOnly)
                {
                    widthParam.Set(0.9); // 90% width
                }
                
                // Tab size for alignment
                Parameter tabSizeParam = newType.get_Parameter(BuiltInParameter.TEXT_TAB_SIZE);
                if (tabSizeParam != null && !tabSizeParam.IsReadOnly)
                {
                    tabSizeParam.Set(modelTextHeight * 4); // Tab size proportional to text
                }
                
                // Color - dark gray for less intrusive appearance
                Parameter colorParam = newType.get_Parameter(BuiltInParameter.LINE_COLOR);
                if (colorParam != null && !colorParam.IsReadOnly)
                {
                    // Set to dark gray (RGB: 64, 64, 64)
                    int colorValue = 64 + (64 << 8) + (64 << 16);
                    colorParam.Set(colorValue);
                }
                
                return newType.Id;
            }
        }
        catch (Exception)
        {
            // If duplication fails, return the base type
            return baseType?.Id ?? ElementId.InvalidElementId;
        }
        
        return ElementId.InvalidElementId;
    }

    private string BuildTextContent(Element opening, FamilyInstance openingInst, 
        List<string> selectedParameterNames, List<string> allParams, Document doc)
    {
        var sb = new StringBuilder();
        ElementType openingType = doc.GetElement(opening.GetTypeId()) as ElementType;

        foreach (string paramName in selectedParameterNames)
        {
            try
            {
                string value = "";

                // Handle core properties
                switch (paramName)
                {
                    case "Element Id":
                        value = opening.Id.IntegerValue.ToString();
                        break;
                    case "Category":
                        value = opening.Category.Name;
                        break;
                    case "Family Name":
                        value = openingType?.FamilyName ?? "";
                        break;
                    case "Instance Name":
                        value = opening.Name;
                        break;
                    case "Level":
                        value = doc.GetElement(opening.LevelId)?.Name ?? "";
                        break;
                    case "Mark":
                        value = opening.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString() ?? "";
                        break;
                    case "Head Height":
                        Parameter headHeight = opening.get_Parameter(BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM);
                        value = headHeight?.AsValueString() ?? "";
                        break;
                    case "Comments":
                        value = opening.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.AsString() ?? "";
                        break;
                    case "Group":
                        value = opening.GroupId != ElementId.InvalidElementId ? doc.GetElement(opening.GroupId)?.Name ?? "" : "";
                        break;
                    case "FacingFlipped":
                        value = openingInst.FacingFlipped.ToString();
                        break;
                    case "HandFlipped":
                        value = openingInst.HandFlipped.ToString();
                        break;
                    case "Wall Direction":
                        value = GetWallDirection(openingInst);
                        break;
                    case "Room From":
                        value = openingInst.FromRoom?.Name ?? "";
                        break;
                    case "Room To":
                        value = openingInst.ToRoom?.Name ?? "";
                        break;
                    case "Width":
                        if (opening.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
                        {
                            Parameter width = openingType?.get_Parameter(BuiltInParameter.DOOR_WIDTH);
                            value = width?.AsValueString() ?? "";
                        }
                        else if (opening.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
                        {
                            Parameter width = openingType?.get_Parameter(BuiltInParameter.WINDOW_WIDTH);
                            value = width?.AsValueString() ?? "";
                        }
                        break;
                    case "Height":
                        if (opening.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
                        {
                            Parameter height = openingType?.get_Parameter(BuiltInParameter.DOOR_HEIGHT);
                            value = height?.AsValueString() ?? "";
                        }
                        else if (opening.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
                        {
                            Parameter height = openingType?.get_Parameter(BuiltInParameter.WINDOW_HEIGHT);
                            value = height?.AsValueString() ?? "";
                        }
                        break;
                    default:
                        // Handle all other parameters
                        if (allParams.Contains(paramName))
                        {
                            var param = opening.LookupParameter(paramName);
                            if (param != null)
                            {
                                value = param.AsValueString() ?? param.AsString() ?? "None";
                            }
                        }
                        break;
                }

                if (!string.IsNullOrEmpty(value))
                {
                    // Add parameter with consistent formatting
                    sb.AppendLine($"{paramName}: {value}");
                }
            }
            catch { }
        }

        return sb.ToString().TrimEnd();
    }

    private string GetWallDirection(FamilyInstance opening)
    {
        try
        {
            // Get the host wall
            Element host = opening.Host;
            if (!(host is Wall wall)) return "";

            // Get wall location curve
            LocationCurve wallLocation = wall.Location as LocationCurve;
            if (wallLocation == null) return "";

            Curve wallCurve = wallLocation.Curve;
            if (wallCurve == null) return "";

            // Get opening location
            LocationPoint openingLocation = opening.Location as LocationPoint;
            if (openingLocation == null) return "";

            XYZ openingPoint = openingLocation.Point;

            // Project opening point onto wall curve to get parameter
            IntersectionResult projection = wallCurve.Project(openingPoint);
            if (projection == null) return "";

            double parameter = projection.Parameter;

            // Get tangent at the parameter
            Transform transform = wallCurve.ComputeDerivatives(parameter, false);
            XYZ tangent = transform.BasisX.Normalize(); // Tangent is the first derivative

            // Use global coordinate system for consistent reference
            // Assume looking in positive Y direction (north), so right is positive X
            XYZ globalRight = XYZ.BasisX; // (1, 0, 0)

            // Check if wall tangent aligns with global right direction
            double dotProduct = tangent.DotProduct(globalRight);

            // If dot product is positive, wall extends to the right (positive X)
            // If dot product is negative, wall extends to the left (negative X)
            return dotProduct > 0 ? "Right" : "Left";
        }
        catch
        {
            return "";
        }
    }
    
    private void ResolveOverlapsBeforeCreation(List<TextNoteInfo> textNoteInfos, double textHeight)
    {
        if (textNoteInfos.Count < 2) return;
        
        // Sort by Y position (top to bottom) then X (left to right)
        textNoteInfos.Sort((a, b) =>
        {
            int yCompare = b.OriginalPosition.Y.CompareTo(a.OriginalPosition.Y);
            return yCompare != 0 ? yCompare : a.OriginalPosition.X.CompareTo(b.OriginalPosition.X);
        });
        
        // Minimum spacing between notes
        double minHorizontalSpacing = textHeight * 0.5;
        double minVerticalSpacing = textHeight * 0.3;
        
        // Process each note
        for (int i = 1; i < textNoteInfos.Count; i++)
        {
            var currentNote = textNoteInfos[i];
            
            // Find best position that doesn't overlap with previous notes
            XYZ bestPosition = FindNonOverlappingPosition(
                currentNote, 
                textNoteInfos.Take(i).ToList(), 
                minHorizontalSpacing, 
                minVerticalSpacing,
                textHeight);
            
            if (!bestPosition.IsAlmostEqualTo(currentNote.OriginalPosition))
            {
                currentNote.CurrentPosition = bestPosition;
                currentNote.WasMoved = true;
            }
        }
    }
    
    private XYZ FindNonOverlappingPosition(
        TextNoteInfo noteToPlace,
        List<TextNoteInfo> placedNotes,
        double minHSpacing,
        double minVSpacing,
        double textHeight)
    {
        // Start with original position
        XYZ testPosition = noteToPlace.OriginalPosition;
        
        // Check if original position works
        if (!HasOverlap(noteToPlace, testPosition, placedNotes, minHSpacing, minVSpacing))
        {
            return testPosition;
        }
        
        // Try positions in a smart pattern: right, left, down-right, down-left, etc.
        double[] xOffsets = { 1.5, -1.5, 1.5, -1.5, 2.5, -2.5, 0, 0 };
        double[] yOffsets = { 0, 0, -1.2, -1.2, -1.2, -1.2, -2.4, -3.6 };
        
        for (int i = 0; i < xOffsets.Length; i++)
        {
            testPosition = new XYZ(
                noteToPlace.OriginalPosition.X + (noteToPlace.Width * xOffsets[i]),
                noteToPlace.OriginalPosition.Y + (noteToPlace.Height * yOffsets[i]),
                noteToPlace.OriginalPosition.Z);
            
            if (!HasOverlap(noteToPlace, testPosition, placedNotes, minHSpacing, minVSpacing))
            {
                return testPosition;
            }
        }
        
        // If no good position found, place below all existing notes
        double lowestY = placedNotes.Min(n => n.CurrentPosition.Y - n.Height);
        return new XYZ(
            noteToPlace.OriginalPosition.X,
            lowestY - minVSpacing - noteToPlace.Height,
            noteToPlace.OriginalPosition.Z);
    }
    
    private bool HasOverlap(
        TextNoteInfo noteToCheck,
        XYZ position,
        List<TextNoteInfo> existingNotes,
        double minHSpacing,
        double minVSpacing)
    {
        // Calculate bounds for note at test position
        double left1 = position.X - minHSpacing;
        double right1 = position.X + noteToCheck.Width + minHSpacing;
        double top1 = position.Y + minVSpacing;
        double bottom1 = position.Y - noteToCheck.Height - minVSpacing;
        
        foreach (var existingNote in existingNotes)
        {
            double left2 = existingNote.CurrentPosition.X - minHSpacing;
            double right2 = existingNote.CurrentPosition.X + existingNote.Width + minHSpacing;
            double top2 = existingNote.CurrentPosition.Y + minVSpacing;
            double bottom2 = existingNote.CurrentPosition.Y - existingNote.Height - minVSpacing;
            
            // Check if rectangles overlap
            bool overlapX = left1 < right2 && right1 > left2;
            bool overlapY = bottom1 < top2 && top1 > bottom2;
            
            if (overlapX && overlapY)
            {
                return true;
            }
        }
        
        return false;
    }
    
    private void CreateLeaderLine(Document doc, View view, TextNoteInfo noteInfo)
    {
        // Get element base point
        BoundingBoxXYZ elementBB = noteInfo.SourceElement.get_BoundingBox(null);
        if (elementBB == null) return;
        
        // Leader starts from left edge of text note
        XYZ textStart = new XYZ(
            noteInfo.CurrentPosition.X,
            noteInfo.CurrentPosition.Y - noteInfo.Height / 2,
            noteInfo.CurrentPosition.Z);
        
        // Leader ends at bottom center of element
        XYZ elementEnd = new XYZ(
            (elementBB.Min.X + elementBB.Max.X) / 2,
            elementBB.Min.Y,
            (elementBB.Min.Z + elementBB.Max.Z) / 2);
        
        try
        {
            Line leaderLine = Line.CreateBound(textStart, elementEnd);
            DetailCurve detailLine = doc.Create.NewDetailCurve(view, leaderLine);
            
            // Set line style to thin
            var lineStyles = new FilteredElementCollector(doc)
                .OfClass(typeof(GraphicsStyle))
                .Cast<GraphicsStyle>()
                .Where(gs => gs.GraphicsStyleType == GraphicsStyleType.Projection)
                .ToList();
            
            var thinStyle = lineStyles.FirstOrDefault(s => 
                s.Name.Contains("Thin") || 
                s.Name.Contains("Fine") || 
                s.Name.Contains("2"));
                
            if (thinStyle != null)
            {
                detailLine.LineStyle = thinStyle;
            }
        }
        catch
        {
            // If leader creation fails, continue without it
        }
    }
}
