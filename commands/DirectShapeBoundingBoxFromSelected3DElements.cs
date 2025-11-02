using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

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
        
        using (Transaction t = new Transaction(doc, "Create DirectShape Bounding Boxes for 3D Elements"))
        {
            t.Start();
            foreach (Element elem in elements3D)
            {
                ElementId dsId = TryCreate3DBoundingBoxDirectShape(doc, elem, activeView);
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
    ElementId TryCreate3DBoundingBoxDirectShape(Document doc, Element sourceElement, View view)
    {
        // Get standard bounding box for Z-adjustment comparison
        BoundingBoxXYZ standardBBox = sourceElement.get_BoundingBox(null);
        
        // Get the actual 3D solid geometry bounding box
        BoundingBoxXYZ bbox = GetTrue3DSolidBoundingBox(sourceElement, doc);
        
        if (bbox == null)
        {
            // Fallback to standard bounding box if no solid geometry found
            bbox = sourceElement.get_BoundingBox(null);
            if (bbox == null || !bbox.Enabled)
                bbox = sourceElement.get_BoundingBox(view);
        }
        else if (standardBBox != null && Math.Abs(bbox.Min.Z - standardBBox.Min.Z) > 0.01)
        {
            // Adjust Z position for families where geometry is centered at origin
            double zAdjustment = standardBBox.Min.Z - bbox.Min.Z;
            XYZ adjustedMin = new XYZ(bbox.Min.X, bbox.Min.Y, standardBBox.Min.Z);
            XYZ adjustedMax = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z + zAdjustment);
            
            bbox = new BoundingBoxXYZ();
            bbox.Min = adjustedMin;
            bbox.Max = adjustedMax;
        }
        
        if (bbox == null || !bbox.Enabled)
            return ElementId.InvalidElementId;

        // Get min and max points
        XYZ min = bbox.Min;
        XYZ max = bbox.Max;

        // Check if bounding box has valid dimensions
        double width = max.X - min.X;
        double depth = max.Y - min.Y;
        double height = max.Z - min.Z;
        
        // Allow very thin elements but require some minimum dimension
        double minDimension = 0.001; // About 0.3mm
        if (width < minDimension || depth < minDimension || height < minDimension)
            return ElementId.InvalidElementId;

        // Create a solid box from the bounding box
        Solid boxSolid = null;
        try
        {
            // Create the bottom face profile (rectangle at min Z)
            XYZ p0 = new XYZ(min.X, min.Y, min.Z);
            XYZ p1 = new XYZ(max.X, min.Y, min.Z);
            XYZ p2 = new XYZ(max.X, max.Y, min.Z);
            XYZ p3 = new XYZ(min.X, max.Y, min.Z);

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
        }
        catch
        {
            return ElementId.InvalidElementId;
        }

        if (boxSolid == null)
            return ElementId.InvalidElementId;

        // Create DirectShape
        DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        if (ds == null)
            return ElementId.InvalidElementId;

        try 
        { 
            ds.SetShape(new GeometryObject[] { boxSolid });
        }
        catch 
        { 
            doc.Delete(ds.Id); 
            return ElementId.InvalidElementId; 
        }

        // Set name based on source element
        string elementName = GetElementName(sourceElement);
        ds.Name = $"BBox3D_{elementName}";

        // Get rotation angle of the source element
        double rotationAngleDegrees = GetElementRotationAngle(sourceElement);

        // Set Comments with Type, Id, and Rotation information as key-value pairs
        string elementTypeName = GetElementTypeName(doc, sourceElement);
        string comments = $"Type: {elementTypeName}, Id: {sourceElement.Id.IntegerValue}, Rotation: {rotationAngleDegrees:F2}°";
        
        Parameter commentsParam = ds.LookupParameter("Comments") ??
                                 ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (commentsParam != null && commentsParam.StorageType == StorageType.String)
        {
            commentsParam.Set(comments);
        }

        return ds.Id;
    }

    // --------------------------------------------------------------------------
    double GetElementRotationAngle(Element elem)
    {
        double rotationAngleDegrees = 0.0;

        // For FamilyInstance elements, we can get the rotation directly
        if (elem is FamilyInstance fi)
        {
            // Get the transform of the family instance
            Transform transform = fi.GetTransform();
            
            // Extract rotation angle from the transform
            // The BasisX vector tells us the rotation around Z axis
            XYZ basisX = transform.BasisX;
            
            // Calculate angle from X-axis (1,0,0) to the element's X-axis
            // atan2 gives us the angle in radians
            double angleRadians = Math.Atan2(basisX.Y, basisX.X);
            rotationAngleDegrees = angleRadians * 180.0 / Math.PI;
            
            // Normalize to 0-360 range
            if (rotationAngleDegrees < 0)
                rotationAngleDegrees += 360.0;
        }
        else
        {
            // For non-FamilyInstance elements, try to get rotation from Location
            Location loc = elem.Location;
            if (loc is LocationPoint locPoint)
            {
                // Some elements store rotation in LocationPoint
                double angleRadians = locPoint.Rotation;
                rotationAngleDegrees = angleRadians * 180.0 / Math.PI;
                
                // Normalize to 0-360 range
                if (rotationAngleDegrees < 0)
                    rotationAngleDegrees += 360.0;
            }
            else if (loc is LocationCurve locCurve)
            {
                // For linear elements like beams, columns, we can get the angle from the curve direction
                Curve curve = locCurve.Curve;
                if (curve is Line line)
                {
                    XYZ direction = line.Direction;
                    // Project to XY plane and calculate angle
                    XYZ xyDirection = new XYZ(direction.X, direction.Y, 0).Normalize();
                    double angleRadians = Math.Atan2(xyDirection.Y, xyDirection.X);
                    rotationAngleDegrees = angleRadians * 180.0 / Math.PI;
                    
                    // Normalize to 0-360 range
                    if (rotationAngleDegrees < 0)
                        rotationAngleDegrees += 360.0;
                }
            }
            else
            {
                // Try to get rotation from geometry transform if available
                Options geomOptions = new Options();
                geomOptions.ComputeReferences = false;
                geomOptions.DetailLevel = ViewDetailLevel.Coarse;
                
                GeometryElement geomElem = elem.get_Geometry(geomOptions);
                if (geomElem != null)
                {
                    foreach (GeometryObject geomObj in geomElem)
                    {
                        if (geomObj is GeometryInstance instance)
                        {
                            Transform transform = instance.Transform;
                            XYZ basisX = transform.BasisX;
                            double angleRadians = Math.Atan2(basisX.Y, basisX.X);
                            rotationAngleDegrees = angleRadians * 180.0 / Math.PI;
                            
                            // Normalize to 0-360 range
                            if (rotationAngleDegrees < 0)
                                rotationAngleDegrees += 360.0;
                            
                            break; // Use first instance transform found
                        }
                    }
                }
            }
        }

        return rotationAngleDegrees;
    }

    // --------------------------------------------------------------------------
    BoundingBoxXYZ GetTrue3DSolidBoundingBox(Element elem, Document doc)
    {
        // Options for getting geometry - include only visible objects
        Options geomOptions = new Options();
        geomOptions.ComputeReferences = false;
        geomOptions.DetailLevel = ViewDetailLevel.Fine;
        geomOptions.IncludeNonVisibleObjects = false; // Exclude invisible geometry
        
        GeometryElement geomElem = elem.get_Geometry(geomOptions);
        if (geomElem == null)
            return null;
        
        // Calculate bounding box from actual visible 3D solids only
        XYZ minPoint = null;
        XYZ maxPoint = null;
        
        foreach (GeometryObject geomObj in geomElem)
        {
            if (geomObj is Solid solid)
            {
                if (solid.Volume > 0)
                {
                    BoundingBoxXYZ solidBBox = solid.GetBoundingBox();
                    if (solidBBox != null)
                    {
                        UpdateBoundingBox(ref minPoint, ref maxPoint, solidBBox);
                    }
                }
            }
            else if (geomObj is GeometryInstance instance)
            {
                // Get transform details
                Transform transform = instance.Transform;
                
                // Process instance geometry WITHOUT transform first
                GeometryElement instGeom = instance.GetInstanceGeometry();
                if (instGeom != null)
                {
                    BoundingBoxXYZ instBBox = GetInstanceSolidBoundingBox(instGeom);
                    if (instBBox != null)
                    {
                        // Check if we need to apply correction for sideways-authored geometry
                        instBBox = ApplySidewaysGeometryCorrection(elem, instBBox, transform);
                        
                        // Transform all 8 corners to world coordinates
                        XYZ[] corners = new XYZ[]
                        {
                            transform.OfPoint(new XYZ(instBBox.Min.X, instBBox.Min.Y, instBBox.Min.Z)),
                            transform.OfPoint(new XYZ(instBBox.Max.X, instBBox.Min.Y, instBBox.Min.Z)),
                            transform.OfPoint(new XYZ(instBBox.Min.X, instBBox.Max.Y, instBBox.Min.Z)),
                            transform.OfPoint(new XYZ(instBBox.Max.X, instBBox.Max.Y, instBBox.Min.Z)),
                            transform.OfPoint(new XYZ(instBBox.Min.X, instBBox.Min.Y, instBBox.Max.Z)),
                            transform.OfPoint(new XYZ(instBBox.Max.X, instBBox.Min.Y, instBBox.Max.Z)),
                            transform.OfPoint(new XYZ(instBBox.Min.X, instBBox.Max.Y, instBBox.Max.Z)),
                            transform.OfPoint(new XYZ(instBBox.Max.X, instBBox.Max.Y, instBBox.Max.Z))
                        };
                        
                        // Find the actual min and max after transformation
                        XYZ worldMin = corners[0];
                        XYZ worldMax = corners[0];
                        foreach (XYZ corner in corners)
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
                        
                        UpdateBoundingBox(ref minPoint, ref maxPoint, worldBBox);
                    }
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
    BoundingBoxXYZ ApplySidewaysGeometryCorrection(Element elem, BoundingBoxXYZ instBBox, Transform transform)
    {
        // Check if this is a door/window with geometry authored sideways
        bool isDoorOrWindow = elem.Category != null && 
            (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors ||
             elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows);
        
        if (!isDoorOrWindow)
            return instBBox;
        
        double localWidth = instBBox.Max.X - instBBox.Min.X;
        double localDepth = instBBox.Max.Y - instBBox.Min.Y;
        bool geometryAuthoredSideways = localDepth > localWidth * 1.5;
        
        if (!geometryAuthoredSideways)
            return instBBox;
        
        // Check if element is mirrored, flipped, or rotated
        bool needsCorrection = false;
        if (elem is FamilyInstance fi)
        {
            // Check if mirrored or hand-flipped
            if (fi.Mirrored || fi.HandFlipped)
            {
                needsCorrection = true;
            }
            else
            {
                // Check if rotated 90° or 270°
                double basisX_Y = Math.Abs(transform.BasisX.Y);
                bool isRotated90Or270 = Math.Abs(basisX_Y - 1.0) < 0.01;
                if (isRotated90Or270)
                    needsCorrection = true;
            }
        }
        
        // Apply correction by swapping X and Y
        if (needsCorrection)
        {
            XYZ correctedMin = new XYZ(instBBox.Min.Y, instBBox.Min.X, instBBox.Min.Z);
            XYZ correctedMax = new XYZ(instBBox.Max.Y, instBBox.Max.X, instBBox.Max.Z);
            
            BoundingBoxXYZ correctedBBox = new BoundingBoxXYZ();
            correctedBBox.Min = correctedMin;
            correctedBBox.Max = correctedMax;
            return correctedBBox;
        }
        
        return instBBox;
    }
    
    // --------------------------------------------------------------------------
    BoundingBoxXYZ GetInstanceSolidBoundingBox(GeometryElement geomElem)
    {
        XYZ minPoint = null;
        XYZ maxPoint = null;
        
        foreach (GeometryObject geomObj in geomElem)
        {
            if (geomObj is Solid solid && solid.Volume > 0)
            {
                BoundingBoxXYZ solidBBox = solid.GetBoundingBox();
                if (solidBBox != null)
                {
                    UpdateBoundingBox(ref minPoint, ref maxPoint, solidBBox);
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
