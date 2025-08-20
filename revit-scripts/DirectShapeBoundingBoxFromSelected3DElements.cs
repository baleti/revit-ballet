using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

[Transaction(TransactionMode.Manual)]
public class DirectShapeBoundingBoxFromSelected3DElements : IExternalCommand
{
    public Result Execute(ExternalCommandData data, ref string message, ElementSet elems)
    {
        UIDocument uidoc = data.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;
        View       activeView = doc.ActiveView;

        // -------------------------------------------------- 1. collect selected 3D model elements
        IList<Element> elements3D = GetSelected3DModelElements(uidoc);
        if (elements3D.Count == 0)
        {
            message = "No 3D model elements selected. Please select one or more 3D model elements (2D elements and annotations are ignored).";
            return Result.Failed;
        }

        // -------------------------------------------------- 2. create DirectShapes
        List<ElementId> createdIds = new List<ElementId>();
        int skippedCount = 0;
        bool showDiagnosticsForSuccess = elements3D.Count == 1; // Only show for single selection
        
        using (Transaction t = new Transaction(doc, "Create DirectShape Bounding Boxes for 3D Elements"))
        {
            t.Start();
            foreach (Element elem in elements3D)
            {
                ElementId dsId = TryCreate3DBoundingBoxDirectShape(doc, elem, activeView, showDiagnosticsForSuccess);
                if (dsId != null && dsId != ElementId.InvalidElementId)
                    createdIds.Add(dsId);
                else
                    skippedCount++;
            }

            if (createdIds.Count == 0)
            {
                message = "No DirectShapes could be created. Elements may not have valid 3D geometry.";
                t.RollBack();
                return Result.Failed;
            }
            t.Commit();
        }
        
        // Add newly created DirectShapes to current selection
        var currentSelection = uidoc.GetSelectionIds().ToList();
        currentSelection.AddRange(createdIds);
        uidoc.SetSelectionIds(currentSelection);
        
        // Report if some elements were skipped
        if (skippedCount > 0)
        {
            TaskDialog.Show("Info", 
                $"Created {createdIds.Count} bounding box(es). " +
                $"Skipped {skippedCount} element(s) without valid 3D geometry.");
        }
        
        return Result.Succeeded;
    }

    // --------------------------------------------------------------------------
    IList<Element> GetSelected3DModelElements(UIDocument uidoc)
    {
        // Get currently selected elements
        var selection = uidoc.GetSelectionIds();
        Document doc = uidoc.Document;
        
        // Filter to get only 3D model elements
        var elements3D = new List<Element>();
        
        foreach (ElementId id in selection)
        {
            Element elem = doc.GetElement(id);
            if (elem == null) continue;
            
            // Skip if it's a view-specific element
            if (elem.ViewSpecific) continue;
            
            // Skip if it's an annotation element
            if (elem is AnnotationSymbol || 
                elem is TextNote || 
                elem is Dimension ||
                elem is DetailLine ||
                elem is DetailArc ||
                elem is DetailCurve ||
                elem is DetailNurbSpline ||
                elem is FilledRegion ||
                elem is IndependentTag ||
                elem is Grid ||
                elem is Level)
                continue;
            
            // Check if element belongs to a model category
            Category cat = elem.Category;
            if (cat == null) continue;
            
            // Skip annotation categories
            if (cat.CategoryType == CategoryType.Annotation) continue;
            
            // Skip certain known 2D/view categories
            if (IsNon3DCategory(cat)) continue;
            
            // Check if element has valid 3D geometry
            Options geomOptions = new Options();
            geomOptions.ComputeReferences = false;
            geomOptions.DetailLevel = ViewDetailLevel.Coarse;
            geomOptions.IncludeNonVisibleObjects = false;
            
            GeometryElement geomElem = elem.get_Geometry(geomOptions);
            if (geomElem == null) continue;
            
            // Check if geometry contains any 3D solids or meshes
            bool has3DGeometry = Has3DGeometry(geomElem);
            if (!has3DGeometry) continue;
            
            // This appears to be a valid 3D model element
            elements3D.Add(elem);
        }

        return elements3D;
    }

    // --------------------------------------------------------------------------
    bool Has3DGeometry(GeometryElement geomElem)
    {
        foreach (GeometryObject geomObj in geomElem)
        {
            if (geomObj is Solid solid && solid.Volume > 0)
                return true;
                
            if (geomObj is Mesh mesh && mesh.NumTriangles > 0)
                return true;
                
            if (geomObj is GeometryInstance instance)
            {
                GeometryElement instGeom = instance.GetInstanceGeometry();
                if (instGeom != null && Has3DGeometry(instGeom))
                    return true;
            }
        }
        return false;
    }

    // --------------------------------------------------------------------------
    bool IsNon3DCategory(Category cat)
    {
        // List of known non-3D categories to skip
        BuiltInCategory bic = (BuiltInCategory)cat.Id.IntegerValue;
        
        switch (bic)
        {
            case BuiltInCategory.OST_Lines:
            case BuiltInCategory.OST_SketchLines:
            case BuiltInCategory.OST_DetailComponents:
            case BuiltInCategory.OST_Dimensions:
            case BuiltInCategory.OST_TextNotes:
            case BuiltInCategory.OST_Tags:
            case BuiltInCategory.OST_GenericAnnotation:
            case BuiltInCategory.OST_Grids:
            case BuiltInCategory.OST_Levels:
            case BuiltInCategory.OST_Views:
            case BuiltInCategory.OST_Sheets:
            case BuiltInCategory.OST_ReferenceLines:
            case BuiltInCategory.OST_CLines:
            case BuiltInCategory.OST_Constraints:
            case BuiltInCategory.OST_SectionBox:
            case BuiltInCategory.OST_Cameras:
                return true;
            default:
                return false;
        }
    }

    // --------------------------------------------------------------------------
    ElementId TryCreate3DBoundingBoxDirectShape(Document doc, Element sourceElement, View view, bool showSuccessDiagnostics = true)
    {
        // Diagnostic data collection
        List<string> diagnostics = new List<string>();
        
        // Get original element's location
        LocationPoint locPoint = sourceElement.Location as LocationPoint;
        XYZ elementLocation = locPoint?.Point;
        if (elementLocation != null)
            diagnostics.Add($"Element Location: ({elementLocation.X:F3}, {elementLocation.Y:F3}, {elementLocation.Z:F3})");
        else
            diagnostics.Add("Element Location: Not a point (possibly line-based)");
            
        // Get standard bounding box for comparison
        BoundingBoxXYZ standardBBox = sourceElement.get_BoundingBox(null);
        if (standardBBox != null)
        {
            diagnostics.Add($"Standard BBox Min: ({standardBBox.Min.X:F3}, {standardBBox.Min.Y:F3}, {standardBBox.Min.Z:F3})");
            diagnostics.Add($"Standard BBox Max: ({standardBBox.Max.X:F3}, {standardBBox.Max.Y:F3}, {standardBBox.Max.Z:F3})");
        }
        
        // Get the actual 3D solid geometry bounding box
        BoundingBoxXYZ bbox = GetTrue3DSolidBoundingBox(sourceElement, doc, ref diagnostics);
        
        if (bbox == null)
        {
            diagnostics.Add("No 3D solid geometry found - falling back to standard bbox");
            // Fallback to standard bounding box if no solid geometry found
            bbox = sourceElement.get_BoundingBox(null);
            if (bbox == null || !bbox.Enabled)
                bbox = sourceElement.get_BoundingBox(view);
        }
        else
        {
            diagnostics.Add($"Calculated BBox Min: ({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3})");
            diagnostics.Add($"Calculated BBox Max: ({bbox.Max.X:F3}, {bbox.Max.Y:F3}, {bbox.Max.Z:F3})");
            
            // Check if we need to adjust Z position based on standard bbox
            // This handles families where geometry is centered at origin rather than based at it
            if (standardBBox != null && Math.Abs(bbox.Min.Z - standardBBox.Min.Z) > 0.01)
            {
                double zAdjustment = standardBBox.Min.Z - bbox.Min.Z;
                diagnostics.Add($"Z Adjustment needed: {zAdjustment:F3} feet (geometry appears centered, not based)");
                
                // Adjust the bounding box to match the standard floor level
                XYZ adjustedMin = new XYZ(bbox.Min.X, bbox.Min.Y, standardBBox.Min.Z);
                XYZ adjustedMax = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z + zAdjustment);
                
                bbox = new BoundingBoxXYZ();
                bbox.Min = adjustedMin;
                bbox.Max = adjustedMax;
                
                diagnostics.Add($"Adjusted BBox Min: ({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3})");
                diagnostics.Add($"Adjusted BBox Max: ({bbox.Max.X:F3}, {bbox.Max.Y:F3}, {bbox.Max.Z:F3})");
            }
            
            // Check if this is a significantly rotated element where our tight box might look wrong
            // Compare aspect ratios to detect if the bounding box seems rotated compared to standard
            if (standardBBox != null)
            {
                double stdWidth = standardBBox.Max.X - standardBBox.Min.X;
                double stdDepth = standardBBox.Max.Y - standardBBox.Min.Y;
                double ourWidth = bbox.Max.X - bbox.Min.X;
                double ourDepth = bbox.Max.Y - bbox.Min.Y;
                
                double stdAspect = stdWidth / stdDepth;
                double ourAspect = ourWidth / ourDepth;
                
                diagnostics.Add($"Standard BBox (includes 2D): W={stdWidth:F3} x D={stdDepth:F3}");
                diagnostics.Add($"Our 3D-only BBox: W={ourWidth:F3} x D={ourDepth:F3}");
                diagnostics.Add($"Standard Aspect (W/D): {stdAspect:F2}");
                diagnostics.Add($"Our Aspect (W/D): {ourAspect:F2}");
                
                // Let's check if the door's geometry is built sideways in the family
                // by seeing if our bbox dimensions match the standard bbox when swapped
                bool dimensionsMatchIfSwapped = 
                    (Math.Abs(ourWidth - stdDepth) < 0.5 && Math.Abs(ourDepth - stdWidth) < 0.5);
                    
                if (dimensionsMatchIfSwapped)
                {
                    diagnostics.Add("!!! DOOR FAMILY GEOMETRY ISSUE DETECTED !!!");
                    diagnostics.Add("The door family appears to be built with geometry rotated 90° in the Family Editor.");
                    diagnostics.Add($"Our W={ourWidth:F3} ≈ Standard D={stdDepth:F3}");
                    diagnostics.Add($"Our D={ourDepth:F3} ≈ Standard W={stdWidth:F3}");
                    diagnostics.Add("This is a FAMILY AUTHORING issue, not a placement issue.");
                    
                    // We could potentially swap the dimensions here, but that would be incorrect
                    // The bounding box we're creating IS correct for the actual 3D geometry
                    diagnostics.Add("The bounding box correctly represents the 3D geometry orientation.");
                }
            }
        }
        
        if (bbox == null || !bbox.Enabled)
        {
            ShowDiagnostics(sourceElement, diagnostics, "Failed: No valid bounding box");
            return ElementId.InvalidElementId;
        }

        // Get min and max points
        XYZ min = bbox.Min;
        XYZ max = bbox.Max;

        // Check if bounding box has valid dimensions
        double width = max.X - min.X;
        double depth = max.Y - min.Y;
        double height = max.Z - min.Z;
        
        diagnostics.Add($"BBox Dimensions: W={width:F3} x D={depth:F3} x H={height:F3} (feet)");
        diagnostics.Add($"BBox Center: ({(min.X + max.X)/2:F3}, {(min.Y + max.Y)/2:F3}, {(min.Z + max.Z)/2:F3})");
        
        // Allow very thin elements but require some minimum dimension
        double minDimension = 0.001; // About 0.3mm
        if (width < minDimension || depth < minDimension || height < minDimension)
        {
            ShowDiagnostics(sourceElement, diagnostics, "Failed: Dimensions too small");
            return ElementId.InvalidElementId;
        }

        // Create a solid box from the bounding box
        Solid boxSolid = null;
        try
        {
            // Create the bottom face profile (rectangle at min Z)
            XYZ p0 = new XYZ(min.X, min.Y, min.Z);
            XYZ p1 = new XYZ(max.X, min.Y, min.Z);
            XYZ p2 = new XYZ(max.X, max.Y, min.Z);
            XYZ p3 = new XYZ(min.X, max.Y, min.Z);
            
            diagnostics.Add($"Creating box corners:");
            diagnostics.Add($"  P0: ({p0.X:F3}, {p0.Y:F3}, {p0.Z:F3})");
            diagnostics.Add($"  P1: ({p1.X:F3}, {p1.Y:F3}, {p1.Z:F3})");
            diagnostics.Add($"  P2: ({p2.X:F3}, {p2.Y:F3}, {p2.Z:F3})");
            diagnostics.Add($"  P3: ({p3.X:F3}, {p3.Y:F3}, {p3.Z:F3})");

            // Create curves for the rectangle
            Line line0 = Line.CreateBound(p0, p1);
            Line line1 = Line.CreateBound(p1, p2);
            Line line2 = Line.CreateBound(p2, p3);
            Line line3 = Line.CreateBound(p3, p0);

            // Create curve loop
            List<Curve> curves = new List<Curve> { line0, line1, line2, line3 };
            CurveLoop curveLoop = CurveLoop.Create(curves);

            // Create solid by extrusion
            boxSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                new[] { curveLoop }, 
                XYZ.BasisZ, 
                height);
                
            diagnostics.Add($"Solid created successfully, extruded {height:F3} feet upward");
        }
        catch (Exception ex)
        {
            diagnostics.Add($"Solid creation failed: {ex.Message}");
            ShowDiagnostics(sourceElement, diagnostics, "Failed: Could not create solid");
            return ElementId.InvalidElementId;
        }

        if (boxSolid == null)
        {
            ShowDiagnostics(sourceElement, diagnostics, "Failed: Solid is null");
            return ElementId.InvalidElementId;
        }

        // Create DirectShape
        DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        if (ds == null)
        {
            ShowDiagnostics(sourceElement, diagnostics, "Failed: Could not create DirectShape");
            return ElementId.InvalidElementId;
        }

        try 
        { 
            ds.SetShape(new GeometryObject[] { boxSolid });
            
            // Verify DirectShape location
            BoundingBoxXYZ dsBBox = ds.get_BoundingBox(null);
            if (dsBBox != null)
            {
                diagnostics.Add($"DirectShape BBox Min: ({dsBBox.Min.X:F3}, {dsBBox.Min.Y:F3}, {dsBBox.Min.Z:F3})");
                diagnostics.Add($"DirectShape BBox Max: ({dsBBox.Max.X:F3}, {dsBBox.Max.Y:F3}, {dsBBox.Max.Z:F3})");
                
                // Check offset
                double offsetX = dsBBox.Min.X - bbox.Min.X;
                double offsetY = dsBBox.Min.Y - bbox.Min.Y;
                double offsetZ = dsBBox.Min.Z - bbox.Min.Z;
                diagnostics.Add($"Offset from intended position: ({offsetX:F3}, {offsetY:F3}, {offsetZ:F3})");
            }
        }
        catch (Exception ex)
        { 
            diagnostics.Add($"SetShape failed: {ex.Message}");
            doc.Delete(ds.Id); 
            ShowDiagnostics(sourceElement, diagnostics, "Failed: Could not set shape");
            return ElementId.InvalidElementId; 
        }

        // Set name based on source element
        string elementName = GetElementName(sourceElement);
        ds.Name = $"BBox3D_{elementName}";

        // Set Comments with Type and Id information as key-value pairs
        string elementTypeName = GetElementTypeName(doc, sourceElement);
        string comments = $"Type: {elementTypeName}, Id: {sourceElement.Id.IntegerValue}";
        
        // Add bounding box dimensions to comments
        string dimensions = $", Dims: {width:F3} x {depth:F3} x {height:F3}";
        comments += dimensions;
        
        Parameter commentsParam = ds.LookupParameter("Comments") ??
                                 ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (commentsParam != null && commentsParam.StorageType == StorageType.String)
        {
            commentsParam.Set(comments);
        }

        // Show diagnostics for successful creation (only if requested)
        if (showSuccessDiagnostics)
            ShowDiagnostics(sourceElement, diagnostics, "SUCCESS: DirectShape created");

        return ds.Id;
    }

    // --------------------------------------------------------------------------
    BoundingBoxXYZ Get3DGeometryBoundingBox(Element elem)
    {
        // Options for getting geometry - not view dependent
        Options geomOptions = new Options();
        geomOptions.ComputeReferences = false;
        geomOptions.DetailLevel = ViewDetailLevel.Medium;
        geomOptions.IncludeNonVisibleObjects = false;
        
        GeometryElement geomElem = elem.get_Geometry(geomOptions);
        if (geomElem == null)
            return null;

        // Calculate bounding box from actual 3D geometry
        XYZ minPoint = null;
        XYZ maxPoint = null;
        
        foreach (GeometryObject geomObj in geomElem)
        {
            BoundingBoxXYZ geomBBox = null;
            
            if (geomObj is Solid solid && solid.Volume > 0)
            {
                geomBBox = solid.GetBoundingBox();
            }
            else if (geomObj is Mesh mesh && mesh.NumTriangles > 0)
            {
                // Calculate bounding box from mesh vertices
                geomBBox = GetMeshBoundingBox(mesh);
            }
            else if (geomObj is GeometryInstance instance)
            {
                GeometryElement instGeom = instance.GetInstanceGeometry();
                if (instGeom != null)
                {
                    BoundingBoxXYZ instBBox = GetGeometryElementBoundingBox(instGeom);
                    if (instBBox != null)
                    {
                        geomBBox = instBBox;
                    }
                }
            }
            
            if (geomBBox != null)
            {
                if (minPoint == null)
                {
                    minPoint = geomBBox.Min;
                    maxPoint = geomBBox.Max;
                }
                else
                {
                    minPoint = new XYZ(
                        Math.Min(minPoint.X, geomBBox.Min.X),
                        Math.Min(minPoint.Y, geomBBox.Min.Y),
                        Math.Min(minPoint.Z, geomBBox.Min.Z));
                    maxPoint = new XYZ(
                        Math.Max(maxPoint.X, geomBBox.Max.X),
                        Math.Max(maxPoint.Y, geomBBox.Max.Y),
                        Math.Max(maxPoint.Z, geomBBox.Max.Z));
                }
            }
        }
        
        if (minPoint == null || maxPoint == null)
            return null;
        
        BoundingBoxXYZ result = new BoundingBoxXYZ();
        result.Min = minPoint;
        result.Max = maxPoint;
        return result;
    }

    // --------------------------------------------------------------------------
    BoundingBoxXYZ GetGeometryElementBoundingBox(GeometryElement geomElem)
    {
        XYZ minPoint = null;
        XYZ maxPoint = null;
        
        foreach (GeometryObject geomObj in geomElem)
        {
            BoundingBoxXYZ geomBBox = null;
            
            if (geomObj is Solid solid && solid.Volume > 0)
            {
                geomBBox = solid.GetBoundingBox();
            }
            else if (geomObj is Mesh mesh && mesh.NumTriangles > 0)
            {
                // Calculate bounding box from mesh vertices
                geomBBox = GetMeshBoundingBox(mesh);
            }
            
            if (geomBBox != null)
            {
                if (minPoint == null)
                {
                    minPoint = geomBBox.Min;
                    maxPoint = geomBBox.Max;
                }
                else
                {
                    minPoint = new XYZ(
                        Math.Min(minPoint.X, geomBBox.Min.X),
                        Math.Min(minPoint.Y, geomBBox.Min.Y),
                        Math.Min(minPoint.Z, geomBBox.Min.Z));
                    maxPoint = new XYZ(
                        Math.Max(maxPoint.X, geomBBox.Max.X),
                        Math.Max(maxPoint.Y, geomBBox.Max.Y),
                        Math.Max(maxPoint.Z, geomBBox.Max.Z));
                }
            }
        }
        
        if (minPoint == null || maxPoint == null)
            return null;
        
        BoundingBoxXYZ result = new BoundingBoxXYZ();
        result.Min = minPoint;
        result.Max = maxPoint;
        return result;
    }

    // --------------------------------------------------------------------------
    void ShowDiagnostics(Element elem, List<string> diagnostics, string status)
    {
        // Save to file on desktop
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string fileName = $"BBox_Diagnostics_{elem.Category?.Name}_{elem.Id.IntegerValue}_{timestamp}.txt";
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string filePath = Path.Combine(desktopPath, fileName);
        
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"=== BOUNDING BOX DIAGNOSTICS ===");
        sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"Element: {elem.Name} (Id: {elem.Id.IntegerValue})");
        sb.AppendLine($"Category: {elem.Category?.Name}");
        sb.AppendLine($"Status: {status}");
        sb.AppendLine();
        
        // Add comprehensive element analysis
        sb.AppendLine("=== ELEMENT PROPERTIES ===");
        CollectElementProperties(elem, sb);
        sb.AppendLine();
        
        sb.AppendLine("=== ELEMENT PARAMETERS ===");
        CollectElementParameters(elem, sb);
        sb.AppendLine();
        
        sb.AppendLine("=== GEOMETRY DIAGNOSTICS ===");
        foreach (string line in diagnostics)
        {
            sb.AppendLine(line);
        }
        
        // Write to file
        File.WriteAllText(filePath, sb.ToString());
        
        // Show simple dialog with file location
        TaskDialog.Show("Diagnostics Saved", $"Diagnostics saved to:\n{filePath}\n\nStatus: {status}");
    }
    
    // --------------------------------------------------------------------------
    void CollectElementProperties(Element elem, StringBuilder sb)
    {
        try
        {
            // Basic properties
            sb.AppendLine($"Class Type: {elem.GetType().Name}");
            sb.AppendLine($"Element Id: {elem.Id.IntegerValue}");
            sb.AppendLine($"UniqueId: {elem.UniqueId}");
            sb.AppendLine($"Name: {elem.Name}");
            sb.AppendLine($"Category: {elem.Category?.Name} (Id: {elem.Category?.Id.IntegerValue})");
            
            // Type information
            ElementId typeId = elem.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                Element elemType = elem.Document.GetElement(typeId);
                sb.AppendLine($"Type Name: {elemType?.Name}");
                sb.AppendLine($"Type Id: {typeId.IntegerValue}");
            }
            
            // Location information
            Location loc = elem.Location;
            if (loc is LocationPoint locPoint)
            {
                sb.AppendLine($"Location Type: Point");
                sb.AppendLine($"Location Point: ({locPoint.Point.X:F6}, {locPoint.Point.Y:F6}, {locPoint.Point.Z:F6})");
                sb.AppendLine($"Rotation: {locPoint.Rotation:F6} radians ({locPoint.Rotation * 180 / Math.PI:F2} degrees)");
            }
            else if (loc is LocationCurve locCurve)
            {
                sb.AppendLine($"Location Type: Curve");
                sb.AppendLine($"Curve Start: ({locCurve.Curve.GetEndPoint(0).X:F6}, {locCurve.Curve.GetEndPoint(0).Y:F6}, {locCurve.Curve.GetEndPoint(0).Z:F6})");
                sb.AppendLine($"Curve End: ({locCurve.Curve.GetEndPoint(1).X:F6}, {locCurve.Curve.GetEndPoint(1).Y:F6}, {locCurve.Curve.GetEndPoint(1).Z:F6})");
            }
            else
            {
                sb.AppendLine($"Location Type: {loc?.GetType().Name ?? "None"}");
            }
            
            // Level information
            if (elem.LevelId != null && elem.LevelId != ElementId.InvalidElementId)
            {
                Level level = elem.Document.GetElement(elem.LevelId) as Level;
                sb.AppendLine($"Level: {level?.Name} (Elevation: {level?.Elevation:F6})");
            }
            
            // Family instance specific
            if (elem is FamilyInstance fi)
            {
                sb.AppendLine($"Is FamilyInstance: Yes");
                sb.AppendLine($"Family Name: {fi.Symbol?.Family?.Name}");
                sb.AppendLine($"Symbol Name: {fi.Symbol?.Name}");
                sb.AppendLine($"Mirrored: {fi.Mirrored}");
                sb.AppendLine($"Hand Flipped: {fi.HandFlipped}");
                sb.AppendLine($"Facing Flipped: {fi.FacingFlipped}");
                sb.AppendLine($"Can Flip Facing: {fi.CanFlipFacing}");
                sb.AppendLine($"Can Flip Hand: {fi.CanFlipHand}");
                
                // Get facing orientation if available
                try
                {
                    XYZ facing = fi.FacingOrientation;
                    sb.AppendLine($"Facing Orientation: ({facing.X:F6}, {facing.Y:F6}, {facing.Z:F6})");
                }
                catch { }
                
                try
                {
                    XYZ hand = fi.HandOrientation;
                    sb.AppendLine($"Hand Orientation: ({hand.X:F6}, {hand.Y:F6}, {hand.Z:F6})");
                }
                catch { }
            }
            
            // View specific
            sb.AppendLine($"View Specific: {elem.ViewSpecific}");
            sb.AppendLine($"Pinned: {elem.Pinned}");
        }
        catch (Exception ex)
        {
            sb.AppendLine($"Error collecting properties: {ex.Message}");
        }
    }
    
    // --------------------------------------------------------------------------
    void CollectElementParameters(Element elem, StringBuilder sb)
    {
        try
        {
            // Get all parameters
            var parameters = elem.Parameters.Cast<Parameter>()
                .OrderBy(p => p.Definition.Name)
                .ToList();
            
            sb.AppendLine($"Total Parameters: {parameters.Count}");
            sb.AppendLine();
            
            // List key parameters
            foreach (Parameter param in parameters)
            {
                if (param == null || param.Definition == null) continue;
                
                string name = param.Definition.Name;
                string value = "N/A";
                
                try
                {
                    switch (param.StorageType)
                    {
                        case StorageType.String:
                            value = param.AsString() ?? "null";
                            break;
                        case StorageType.Integer:
                            value = param.AsInteger().ToString();
                            break;
                        case StorageType.Double:
                            double d = param.AsDouble();
                            value = $"{d:F6}";
                            // Try to get display value too
                            string displayValue = param.AsValueString();
                            if (!string.IsNullOrEmpty(displayValue))
                                value += $" ({displayValue})";
                            break;
                        case StorageType.ElementId:
                            ElementId id = param.AsElementId();
                            if (id != null && id != ElementId.InvalidElementId)
                            {
                                Element e = elem.Document.GetElement(id);
                                value = $"Id: {id.IntegerValue} ({e?.Name ?? "Not found"})";
                            }
                            else
                            {
                                value = "InvalidElementId";
                            }
                            break;
                    }
                }
                catch { }
                
                // Only include parameters with values
                if (value != "N/A" && value != "null" && value != "0" && value != "0.000000")
                {
                    sb.AppendLine($"{name}: {value}");
                }
            }
        }
        catch (Exception ex)
        {
            sb.AppendLine($"Error collecting parameters: {ex.Message}");
        }
    }
    
    // --------------------------------------------------------------------------
    BoundingBoxXYZ GetTrue3DSolidBoundingBox(Element elem, Document doc, ref List<string> diagnostics)
    {
        // Options for getting geometry - include only visible objects
        Options geomOptions = new Options();
        geomOptions.ComputeReferences = false;
        geomOptions.DetailLevel = ViewDetailLevel.Fine;
        geomOptions.IncludeNonVisibleObjects = false; // This is key - exclude invisible geometry
        
        GeometryElement geomElem = elem.get_Geometry(geomOptions);
        if (geomElem == null)
        {
            diagnostics.Add("No geometry element found");
            return null;
        }

        // Set to false to hide detailed geometry info
        bool showDetailedGeomInfo = true;
        
        // Collect debug information
        List<string> geometryInfo = new List<string>();
        int solidCount = 0;
        int curveCount = 0;
        int meshCount = 0;
        int instanceCount = 0;
        double totalVolume = 0;
        
        // Calculate bounding box from actual visible 3D solids only
        XYZ minPoint = null;
        XYZ maxPoint = null;
        
        foreach (GeometryObject geomObj in geomElem)
        {
            if (geomObj is Solid solid)
            {
                if (solid.Volume > 0)
                {
                    solidCount++;
                    totalVolume += solid.Volume;
                    BoundingBoxXYZ solidBBox = solid.GetBoundingBox();
                    
                    if (solidBBox != null)
                    {
                        UpdateBoundingBox(ref minPoint, ref maxPoint, solidBBox);
                        if (showDetailedGeomInfo)
                            geometryInfo.Add($"Solid {solidCount}: Vol={solid.Volume:F4}, Min=({solidBBox.Min.X:F3},{solidBBox.Min.Y:F3},{solidBBox.Min.Z:F3})");
                    }
                }
            }
            else if (geomObj is Curve curve)
            {
                curveCount++;
            }
            else if (geomObj is Mesh mesh)
            {
                meshCount++;
            }
            else if (geomObj is GeometryInstance instance)
            {
                instanceCount++;
                
                // Get transform details
                Transform transform = instance.Transform;
                diagnostics.Add($"GeometryInstance {instanceCount} Transform:");
                diagnostics.Add($"  Origin: ({transform.Origin.X:F3}, {transform.Origin.Y:F3}, {transform.Origin.Z:F3})");
                diagnostics.Add($"  BasisX: ({transform.BasisX.X:F3}, {transform.BasisX.Y:F3}, {transform.BasisX.Z:F3})");
                diagnostics.Add($"  BasisY: ({transform.BasisY.X:F3}, {transform.BasisY.Y:F3}, {transform.BasisY.Z:F3})");
                diagnostics.Add($"  BasisZ: ({transform.BasisZ.X:F3}, {transform.BasisZ.Y:F3}, {transform.BasisZ.Z:F3})");
                diagnostics.Add($"  Scale: {transform.Scale:F3}");
                
                // Check if element is rotated
                bool isRotated = Math.Abs(transform.BasisX.X - 1.0) > 0.01 || 
                                Math.Abs(transform.BasisY.Y - 1.0) > 0.01;
                if (isRotated)
                {
                    double rotationAngle = Math.Atan2(transform.BasisX.Y, transform.BasisX.X) * 180.0 / Math.PI;
                    diagnostics.Add($"  Element is ROTATED: {rotationAngle:F1} degrees from X-axis");
                }
                
                // Process instance geometry WITHOUT transform first
                GeometryElement instGeom = instance.GetInstanceGeometry();
                if (instGeom != null)
                {
                    // Collect all solids from the instance for detailed analysis
                    List<Solid> instanceSolids = new List<Solid>();
                    foreach (GeometryObject obj in instGeom)
                    {
                        if (obj is Solid instSolid && instSolid.Volume > 0)
                            instanceSolids.Add(instSolid);
                    }
                    
                    diagnostics.Add($"  Instance contains {instanceSolids.Count} solid(s)");
                    
                    // Analyze each solid
                    for (int i = 0; i < instanceSolids.Count && i < 5; i++) // Limit to first 5 solids
                    {
                        Solid s = instanceSolids[i];
                        BoundingBoxXYZ sBBox = s.GetBoundingBox();
                        diagnostics.Add($"  Solid {i+1}: Vol={s.Volume:F4}, Faces={s.Faces.Size}");
                        diagnostics.Add($"    BBox: ({sBBox.Min.X:F3},{sBBox.Min.Y:F3},{sBBox.Min.Z:F3}) to ({sBBox.Max.X:F3},{sBBox.Max.Y:F3},{sBBox.Max.Z:F3})");
                        diagnostics.Add($"    Dims: {sBBox.Max.X-sBBox.Min.X:F3} x {sBBox.Max.Y-sBBox.Min.Y:F3} x {sBBox.Max.Z-sBBox.Min.Z:F3}");
                    }
                    
                    BoundingBoxXYZ instBBox = GetInstanceSolidBoundingBox(instGeom, ref geometryInfo, showDetailedGeomInfo);
                    if (instBBox != null)
                    {
                        diagnostics.Add($"  Combined Local BBox: Min=({instBBox.Min.X:F3},{instBBox.Min.Y:F3},{instBBox.Min.Z:F3}) Max=({instBBox.Max.X:F3},{instBBox.Max.Y:F3},{instBBox.Max.Z:F3})");
                        diagnostics.Add($"  Local Dimensions: W={instBBox.Max.X - instBBox.Min.X:F3} x D={instBBox.Max.Y - instBBox.Min.Y:F3} x H={instBBox.Max.Z - instBBox.Min.Z:F3}");
                        
                        // Check if the family geometry seems to be authored sideways
                        double localWidth = instBBox.Max.X - instBBox.Min.X;
                        double localDepth = instBBox.Max.Y - instBBox.Min.Y;
                        
                        // Check if this is a door/window with geometry authored sideways
                        bool isDoorOrWindow = elem.Category != null && 
                            (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors ||
                             elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows);
                        
                        bool geometryAuthoredSideways = isDoorOrWindow && localDepth > localWidth * 1.5;
                        
                        // Check if element is mirrored or flipped
                        bool isMirroredOrFlipped = false;
                        if (elem is FamilyInstance fi)
                        {
                            isMirroredOrFlipped = fi.Mirrored || fi.HandFlipped;
                            diagnostics.Add($"  Mirrored: {fi.Mirrored}, HandFlipped: {fi.HandFlipped}, FacingFlipped: {fi.FacingFlipped}");
                        }
                        
                        // If geometry is authored sideways AND element is mirrored/flipped, we need to swap dimensions
                        if (geometryAuthoredSideways && isMirroredOrFlipped)
                        {
                            diagnostics.Add("  >>> CORRECTION NEEDED: Family authored sideways + mirrored/flipped!");
                            diagnostics.Add($"  >>> Swapping local X and Y before transformation");
                            
                            // Swap X and Y in the local bounding box
                            XYZ correctedMin = new XYZ(instBBox.Min.Y, instBBox.Min.X, instBBox.Min.Z);
                            XYZ correctedMax = new XYZ(instBBox.Max.Y, instBBox.Max.X, instBBox.Max.Z);
                            
                            instBBox = new BoundingBoxXYZ();
                            instBBox.Min = correctedMin;
                            instBBox.Max = correctedMax;
                            
                            diagnostics.Add($"  Corrected Local BBox: Min=({instBBox.Min.X:F3},{instBBox.Min.Y:F3},{instBBox.Min.Z:F3}) Max=({instBBox.Max.X:F3},{instBBox.Max.Y:F3},{instBBox.Max.Z:F3})");
                        }
                        else if (geometryAuthoredSideways)
                        {
                            diagnostics.Add("  >>> Family appears to be authored with geometry rotated in Family Editor!");
                            diagnostics.Add($"  >>> Local geometry has depth ({localDepth:F3}) > width ({localWidth:F3})");
                        }
                        
                        // Show all 8 corners before and after transformation
                        diagnostics.Add("  === Corner Transformations ===");
                        XYZ[] localCorners = new XYZ[]
                        {
                            new XYZ(instBBox.Min.X, instBBox.Min.Y, instBBox.Min.Z),
                            new XYZ(instBBox.Max.X, instBBox.Min.Y, instBBox.Min.Z),
                            new XYZ(instBBox.Min.X, instBBox.Max.Y, instBBox.Min.Z),
                            new XYZ(instBBox.Max.X, instBBox.Max.Y, instBBox.Min.Z),
                            new XYZ(instBBox.Min.X, instBBox.Min.Y, instBBox.Max.Z),
                            new XYZ(instBBox.Max.X, instBBox.Min.Y, instBBox.Max.Z),
                            new XYZ(instBBox.Min.X, instBBox.Max.Y, instBBox.Max.Z),
                            new XYZ(instBBox.Max.X, instBBox.Max.Y, instBBox.Max.Z)
                        };
                        
                        XYZ[] worldCorners = new XYZ[8];
                        for (int i = 0; i < 8; i++)
                        {
                            worldCorners[i] = transform.OfPoint(localCorners[i]);
                            if (i < 4) // Show first 4 corners
                            {
                                diagnostics.Add($"  C{i}: ({localCorners[i].X:F3},{localCorners[i].Y:F3},{localCorners[i].Z:F3}) -> ({worldCorners[i].X:F3},{worldCorners[i].Y:F3},{worldCorners[i].Z:F3})");
                            }
                        }
                        
                        // Find the actual min and max after transformation
                        XYZ worldMin = worldCorners[0];
                        XYZ worldMax = worldCorners[0];
                        foreach (XYZ corner in worldCorners)
                        {
                            worldMin = new XYZ(
                                Math.Min(worldMin.X, corner.X),
                                Math.Min(worldMin.Y, corner.Y),
                                Math.Min(worldMin.Z, corner.Z));
                            worldMax = new XYZ(
                                Math.Max(worldMax.X, corner.X),
                                Math.Max(worldMax.Y, corner.Y),
                                Math.Max(worldMax.Z, corner.Z));
                        }
                        
                        // Create properly positioned bounding box
                        BoundingBoxXYZ worldBBox = new BoundingBoxXYZ();
                        worldBBox.Min = worldMin;
                        worldBBox.Max = worldMax;
                        
                        diagnostics.Add($"  World BBox: Min=({worldBBox.Min.X:F3},{worldBBox.Min.Y:F3},{worldBBox.Min.Z:F3}) Max=({worldBBox.Max.X:F3},{worldBBox.Max.Y:F3},{worldBBox.Max.Z:F3})");
                        diagnostics.Add($"  World Dimensions: W={worldBBox.Max.X - worldBBox.Min.X:F3} x D={worldBBox.Max.Y - worldBBox.Min.Y:F3} x H={worldBBox.Max.Z - worldBBox.Min.Z:F3}");
                        
                        UpdateBoundingBox(ref minPoint, ref maxPoint, worldBBox);
                    }
                }
            }
        }
        
        diagnostics.Add($"Geometry Summary: Solids={solidCount}, Curves={curveCount}, Meshes={meshCount}, Instances={instanceCount}");
        diagnostics.Add($"Total Volume: {totalVolume:F4} cubic feet");
        
        if (minPoint == null || maxPoint == null)
        {
            diagnostics.Add("No valid 3D solid geometry found");
            return null;
        }
        
        BoundingBoxXYZ result = new BoundingBoxXYZ();
        result.Min = minPoint;
        result.Max = maxPoint;
        return result;
    }
    
    // --------------------------------------------------------------------------
    BoundingBoxXYZ GetInstanceSolidBoundingBox(GeometryElement geomElem, ref List<string> geometryInfo, bool collectDebugInfo = false)
    {
        XYZ minPoint = null;
        XYZ maxPoint = null;
        int solidCount = 0;
        
        foreach (GeometryObject geomObj in geomElem)
        {
            if (geomObj is Solid solid && solid.Volume > 0)
            {
                solidCount++;
                BoundingBoxXYZ solidBBox = solid.GetBoundingBox();
                if (solidBBox != null)
                {
                    UpdateBoundingBox(ref minPoint, ref maxPoint, solidBBox);
                    if (collectDebugInfo)
                        geometryInfo.Add($"  Instance Solid {solidCount}: Vol={solid.Volume:F4}");
                }
            }
        }
        
        if (minPoint == null || maxPoint == null)
            return null;
        
        BoundingBoxXYZ result = new BoundingBoxXYZ();
        result.Min = minPoint;
        result.Max = maxPoint;
        return result;
    }
    
    // --------------------------------------------------------------------------
    void UpdateBoundingBox(ref XYZ minPoint, ref XYZ maxPoint, BoundingBoxXYZ newBBox)
    {
        if (minPoint == null)
        {
            minPoint = newBBox.Min;
            maxPoint = newBBox.Max;
        }
        else
        {
            minPoint = new XYZ(
                Math.Min(minPoint.X, newBBox.Min.X),
                Math.Min(minPoint.Y, newBBox.Min.Y),
                Math.Min(minPoint.Z, newBBox.Min.Z));
            maxPoint = new XYZ(
                Math.Max(maxPoint.X, newBBox.Max.X),
                Math.Max(maxPoint.Y, newBBox.Max.Y),
                Math.Max(maxPoint.Z, newBBox.Max.Z));
        }
    }
    
    // --------------------------------------------------------------------------
    string GetElementName(Element elem)
    {
        // Try to get element name
        string name = elem.Name;
        
        // If Name is empty, try to get type name
        if (string.IsNullOrWhiteSpace(name))
        {
            Parameter typeParam = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
            if (typeParam != null)
            {
                ElementId typeId = typeParam.AsElementId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                {
                    Element elemType = elem.Document.GetElement(typeId);
                    if (elemType != null)
                        name = elemType.Name;
                }
            }
        }

        // If still no name, use category name
        if (string.IsNullOrWhiteSpace(name))
        {
            if (elem.Category != null)
                name = elem.Category.Name;
            else
                name = "Element";
        }

        // Sanitize the name for DirectShape naming
        return Sanitize(name);
    }

    // --------------------------------------------------------------------------
    string GetElementTypeName(Document doc, Element elem)
    {
        string typeName = elem.GetType().Name;
        
        // Try to get the Revit element type name
        ElementId typeId = elem.GetTypeId();
        if (typeId != null && typeId != ElementId.InvalidElementId)
        {
            Element elemType = doc.GetElement(typeId);
            if (elemType != null && !string.IsNullOrWhiteSpace(elemType.Name))
            {
                typeName = $"{elem.Category?.Name ?? typeName}: {elemType.Name}";
            }
        }
        else if (elem.Category != null)
        {
            typeName = elem.Category.Name;
        }

        return typeName;
    }

    // --------------------------------------------------------------------------
    BoundingBoxXYZ GetMeshBoundingBox(Mesh mesh)
    {
        if (mesh == null || mesh.Vertices.Count == 0)
            return null;
            
        XYZ firstVertex = mesh.Vertices[0];
        double minX = firstVertex.X, minY = firstVertex.Y, minZ = firstVertex.Z;
        double maxX = firstVertex.X, maxY = firstVertex.Y, maxZ = firstVertex.Z;
        
        foreach (XYZ vertex in mesh.Vertices)
        {
            minX = Math.Min(minX, vertex.X);
            minY = Math.Min(minY, vertex.Y);
            minZ = Math.Min(minZ, vertex.Z);
            maxX = Math.Max(maxX, vertex.X);
            maxY = Math.Max(maxY, vertex.Y);
            maxZ = Math.Max(maxZ, vertex.Z);
        }
        
        BoundingBoxXYZ bbox = new BoundingBoxXYZ();
        bbox.Min = new XYZ(minX, minY, minZ);
        bbox.Max = new XYZ(maxX, maxY, maxZ);
        return bbox;
    }
    
    // --------------------------------------------------------------------------
    static string Sanitize(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return "Unnamed";
            
        // Remove illegal characters for Revit names
        System.Text.RegularExpressions.Regex illegal = 
            new System.Text.RegularExpressions.Regex(@"[<>:{}|;?*\\/\[\]]");
        string clean = illegal.Replace(raw, "_").Trim();
        
        // Limit length
        return clean.Length > 250 ? clean.Substring(0, 250) : clean;
    }
}
