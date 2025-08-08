using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
    // Helper classes for transformation
    private class ReferenceElements
    {
        public XYZ GroupOrigin { get; set; }
        public List<ElementInfo> Elements { get; set; } = new List<ElementInfo>();
    }
    
    private class ElementInfo
    {
        public Element Element { get; set; }
        public string UniqueKey { get; set; }
        public XYZ Point1 { get; set; }
        public XYZ Point2 { get; set; }
    }
    
    private class TransformResult
    {
        public XYZ Translation { get; set; }
        public double Rotation { get; set; }
        public bool IsMirrored { get; set; }
        public XYZ MirrorAxis { get; set; } // Which axis/axes are mirrored
        public double Scale { get; set; }
        public int MatchingElements { get; set; }
    }
    
    // Get reference elements from a group for transformation calculation
    private ReferenceElements GetReferenceElements(Group group, Document doc)
    {
        ReferenceElements refElements = new ReferenceElements();
        refElements.GroupOrigin = (group.Location as LocationPoint).Point;
        
        // Get group members
        ICollection<ElementId> memberIds = group.GetMemberIds();
        
        // Collect all elements by type
        List<Wall> walls = new List<Wall>();
        List<Element> curveElements = new List<Element>();
        List<Element> pointElements = new List<Element>();
        
        foreach (ElementId id in memberIds)
        {
            Element elem = doc.GetElement(id);
            if (elem == null) continue;
            
            if (elem is Wall)
            {
                walls.Add(elem as Wall);
            }
            else if (elem.Location is LocationCurve)
            {
                curveElements.Add(elem);
            }
            else if (elem.Location is LocationPoint)
            {
                pointElements.Add(elem);
            }
        }
        
        // Build unique elements list
        Dictionary<string, ElementInfo> uniqueElements = new Dictionary<string, ElementInfo>();
        
        // Process walls first (preferred because they have two points)
        foreach (Wall wall in walls)
        {
            string key = GetEnhancedUniqueKey(wall);
            
            if (!uniqueElements.ContainsKey(key))
            {
                ElementInfo info = new ElementInfo();
                info.Element = wall;
                info.UniqueKey = key;
                
                LocationCurve locCurve = wall.Location as LocationCurve;
                if (locCurve != null)
                {
                    info.Point1 = locCurve.Curve.GetEndPoint(0);
                    info.Point2 = locCurve.Curve.GetEndPoint(1);
                }
                
                uniqueElements[key] = info;
            }
        }
        
        // If we need more elements, add curve elements
        if (uniqueElements.Count < 2)
        {
            foreach (Element elem in curveElements)
            {
                string key = GetEnhancedUniqueKey(elem);
                
                if (!uniqueElements.ContainsKey(key))
                {
                    ElementInfo info = new ElementInfo();
                    info.Element = elem;
                    info.UniqueKey = key;
                    
                    LocationCurve locCurve = elem.Location as LocationCurve;
                    if (locCurve != null)
                    {
                        info.Point1 = locCurve.Curve.GetEndPoint(0);
                        info.Point2 = locCurve.Curve.GetEndPoint(1);
                    }
                    
                    uniqueElements[key] = info;
                }
            }
        }
        
        // If still need more, add point elements
        if (uniqueElements.Count < 2)
        {
            foreach (Element elem in pointElements)
            {
                string key = GetEnhancedUniqueKey(elem);
                
                if (!uniqueElements.ContainsKey(key))
                {
                    ElementInfo info = new ElementInfo();
                    info.Element = elem;
                    info.UniqueKey = key;
                    
                    LocationPoint locPoint = elem.Location as LocationPoint;
                    if (locPoint != null)
                    {
                        info.Point1 = locPoint.Point;
                        info.Point2 = locPoint.Point; // Same point for point elements
                    }
                    
                    uniqueElements[key] = info;
                }
            }
        }
        
        refElements.Elements = uniqueElements.Values.ToList();
        return refElements;
    }
    
    // Create unique key for element identification
    private string GetEnhancedUniqueKey(Element elem)
    {
        // Create unique key based on element type and parameters
        string key = elem.GetType().Name + "|";
        
        // Add category
        if (elem.Category != null)
            key += elem.Category.Id.IntegerValue + "|";
        
        // Add type id
        ElementId typeId = elem.GetTypeId();
        if (typeId != ElementId.InvalidElementId)
            key += typeId.IntegerValue + "|";
        
        // Add geometric properties for walls
        if (elem is Wall)
        {
            Wall wall = elem as Wall;
            
            // Add wall length
            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve != null)
            {
                double length = locCurve.Curve.Length;
                key += $"L:{length:F6}|";
            }
            
            // Add wall height
            Parameter heightParam = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
            if (heightParam != null && heightParam.HasValue)
            {
                key += $"H:{heightParam.AsDouble():F6}|";
            }
        }
        // Add properties for other elements
        else if (elem.Category != null)
        {
            // For curve-based elements, add length
            if (elem.Location is LocationCurve)
            {
                LocationCurve locCurve = elem.Location as LocationCurve;
                double length = locCurve.Curve.Length;
                key += $"L:{length:F6}|";
            }
        }
        
        // Add instance comments if present
        Parameter nameParam = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (nameParam != null && nameParam.HasValue && !string.IsNullOrEmpty(nameParam.AsString()))
            key += nameParam.AsString() + "|";
        
        return key;
    }
    
    // Calculate transformation between reference group and target group
    private TransformResult CalculateTransformation(ReferenceElements refElements, Group otherGroup, Document doc)
    {
        try
        {
            // Get corresponding elements in other group
            List<ElementInfo> otherElements = GetCorrespondingElements(refElements, otherGroup, doc);
            
            // We need at least 2 matching elements to calculate transformation
            if (otherElements == null || otherElements.Count < 2) 
            {
                diagnosticLog.AppendLine($"    Group {otherGroup.Id}: Only {otherElements?.Count ?? 0} corresponding elements found");
                
                if (otherElements == null || otherElements.Count == 0)
                    transformFailureReasons["No Corresponding Elements"]++;
                else
                    transformFailureReasons["Insufficient Matching Elements"]++;
                
                return null;
            }
            
            TransformResult result = new TransformResult();
            result.MatchingElements = otherElements.Count;
            
            // Calculate translation (difference in group origins)
            XYZ otherOrigin = (otherGroup.Location as LocationPoint).Point;
            result.Translation = otherOrigin - refElements.GroupOrigin;
            
            // Find two matching elements for transformation calculation
            ElementInfo ref1 = null;
            ElementInfo ref2 = null;
            ElementInfo other1 = null;
            ElementInfo other2 = null;
            
            // Find first matching pair
            foreach (var refElem in refElements.Elements)
            {
                var match = otherElements.FirstOrDefault(e => e.UniqueKey == refElem.UniqueKey);
                if (match != null)
                {
                    ref1 = refElem;
                    other1 = match;
                    break;
                }
            }
            
            // Find second matching pair (different from first)
            foreach (var refElem in refElements.Elements)
            {
                if (refElem == ref1) continue;
                var match = otherElements.FirstOrDefault(e => e.UniqueKey == refElem.UniqueKey);
                if (match != null)
                {
                    ref2 = refElem;
                    other2 = match;
                    break;
                }
            }
            
            if (ref1 == null || ref2 == null || other1 == null || other2 == null)
            {
                // If we can't find two different matching elements, try to work with just one
                if (ref1 != null && other1 != null && ref1.Point1 != null && ref1.Point2 != null)
                {
                    // Can still calculate basic rotation from one element
                    XYZ refVector = (ref1.Point2 - ref1.Point1).Normalize();
                    XYZ otherVector = (other1.Point2 - other1.Point1).Normalize();
                    
                    double singleDotProduct = refVector.DotProduct(otherVector);
                    double singleAngle = Math.Acos(Math.Max(-1.0, Math.Min(1.0, singleDotProduct)));
                    
                    XYZ singleCrossProduct = refVector.CrossProduct(otherVector);
                    if (singleCrossProduct.Z < 0) singleAngle = -singleAngle;
                    
                    result.Rotation = singleAngle * 180 / Math.PI;
                    result.IsMirrored = false; // Can't determine mirroring with one element
                    result.MirrorAxis = XYZ.Zero;
                    result.Scale = 1.0;
                    
                    return result;
                }
                
                diagnosticLog.AppendLine($"    Group {otherGroup.Id}: Could not find sufficient matching element pairs");
                transformFailureReasons["Insufficient Matching Elements"]++;
                return null;
            }
            
            // Calculate vectors in reference group
            XYZ refVector1 = (ref1.Point2 - ref1.Point1).Normalize();
            XYZ refVector2 = (ref2.Point2 - ref2.Point1).Normalize();
            
            // Calculate vectors in other group  
            XYZ otherVector1 = (other1.Point2 - other1.Point1).Normalize();
            XYZ otherVector2 = (other2.Point2 - other2.Point1).Normalize();
            
            // For mirrored groups, the matching elements might have reversed directions
            // Check if vectors are opposite
            double dotProduct = refVector1.DotProduct(otherVector1);
            bool vectorsReversed = dotProduct < -0.9; // Nearly opposite
            
            if (vectorsReversed)
            {
                // For mirrored walls, the direction might be flipped
                // This is common when groups are mirrored
                otherVector1 = -otherVector1;
                otherVector2 = -otherVector2;
                dotProduct = refVector1.DotProduct(otherVector1);
            }
            
            double angle = Math.Acos(Math.Max(-1.0, Math.Min(1.0, dotProduct)));
            
            // Determine rotation direction
            XYZ crossProduct = refVector1.CrossProduct(otherVector1);
            if (crossProduct.Z < 0) angle = -angle;
            
            result.Rotation = angle * 180 / Math.PI;
            
            // Enhanced mirror detection for different mirror configurations
            // Check the relative positions
            XYZ refMid1 = (ref1.Point1 + ref1.Point2) * 0.5;
            XYZ refMid2 = (ref2.Point1 + ref2.Point2) * 0.5;
            XYZ otherMid1 = (other1.Point1 + other1.Point2) * 0.5;
            XYZ otherMid2 = (other2.Point1 + other2.Point2) * 0.5;
            
            // Calculate relative positions from group origins
            XYZ refPos1 = refMid1 - refElements.GroupOrigin;
            XYZ refPos2 = refMid2 - refElements.GroupOrigin;
            XYZ otherPos1 = otherMid1 - otherOrigin;
            XYZ otherPos2 = otherMid2 - otherOrigin;
            
            // Check for different mirror configurations
            double tolerance = 0.5; // Increased tolerance for real-world models
            
            // Check if X is mirrored
            bool xMirrored = Math.Abs(refPos1.X + otherPos1.X) < tolerance && 
                             Math.Abs(refPos2.X + otherPos2.X) < tolerance;
            
            // Check if Y is mirrored
            bool yMirrored = Math.Abs(refPos1.Y + otherPos1.Y) < tolerance && 
                             Math.Abs(refPos2.Y + otherPos2.Y) < tolerance;
            
            // Determine mirror configuration
            if (xMirrored && yMirrored)
            {
                // Both axes mirrored (180-degree rotation)
                result.IsMirrored = true;
                result.MirrorAxis = new XYZ(1, 1, 0); // Both X and Y
                diagnosticLog.AppendLine($"    Detected XY mirror (180° rotation)");
            }
            else if (xMirrored)
            {
                // Only X mirrored
                result.IsMirrored = true;
                result.MirrorAxis = XYZ.BasisX;
                diagnosticLog.AppendLine($"    Detected X mirror");
            }
            else if (yMirrored)
            {
                // Only Y mirrored
                result.IsMirrored = true;
                result.MirrorAxis = XYZ.BasisY;
                diagnosticLog.AppendLine($"    Detected Y mirror");
            }
            else if (vectorsReversed)
            {
                // Vectors reversed but positions don't clearly indicate axis
                result.IsMirrored = true;
                result.MirrorAxis = XYZ.BasisX; // Default to X
                diagnosticLog.AppendLine($"    Detected mirror from reversed vectors");
            }
            else
            {
                result.IsMirrored = false;
                result.MirrorAxis = XYZ.Zero;
            }
            
            // Scale (typically 1.0 for groups)
            result.Scale = 1.0;
            
            return result;
        }
        catch (Exception ex)
        {
            diagnosticLog.AppendLine($"    Group {otherGroup.Id}: Exception during transform calculation: {ex.Message}");
            transformFailureReasons["Exception During Transform"]++;
            return null;
        }
    }
    
    // Get corresponding elements in target group
    private List<ElementInfo> GetCorrespondingElements(ReferenceElements refElements, Group otherGroup, Document doc)
    {
        List<ElementInfo> otherElements = new List<ElementInfo>();
        
        // Get all members of the other group
        ICollection<ElementId> memberIds = otherGroup.GetMemberIds();
        
        // Create a dictionary of elements in the other group by their unique key
        Dictionary<string, Element> otherGroupElements = new Dictionary<string, Element>();
        foreach (ElementId id in memberIds)
        {
            Element elem = doc.GetElement(id);
            if (elem == null) continue;
            
            string key = GetEnhancedUniqueKey(elem);
            if (!otherGroupElements.ContainsKey(key))
            {
                otherGroupElements[key] = elem;
            }
        }
        
        // Find matching elements
        foreach (ElementInfo refInfo in refElements.Elements)
        {
            if (otherGroupElements.ContainsKey(refInfo.UniqueKey))
            {
                Element elem = otherGroupElements[refInfo.UniqueKey];
                
                ElementInfo info = new ElementInfo();
                info.Element = elem;
                info.UniqueKey = refInfo.UniqueKey;
                
                if (elem.Location is LocationCurve)
                {
                    LocationCurve locCurve = elem.Location as LocationCurve;
                    info.Point1 = locCurve.Curve.GetEndPoint(0);
                    info.Point2 = locCurve.Curve.GetEndPoint(1);
                }
                else if (elem.Location is LocationPoint)
                {
                    LocationPoint locPoint = elem.Location as LocationPoint;
                    info.Point1 = locPoint.Point;
                    info.Point2 = locPoint.Point;
                }
                
                otherElements.Add(info);
            }
        }
        
        return otherElements;
    }
    
    // FIXED CreateTransform method that properly handles rotation and mirroring
    // THIS IS THE KEY METHOD THAT NEEDS FIXING FOR "Core - B3" AND SIMILAR GROUPS
    private Transform CreateTransform(TransformResult transformResult, XYZ refOrigin, XYZ targetOrigin)
    {
        // Calculate the translation needed
        XYZ translation = targetOrigin - refOrigin;
        
        diagnosticLog.AppendLine($"      Translation: X={translation.X:F3}, Y={translation.Y:F3}, Z={translation.Z:F3}");
        diagnosticLog.AppendLine($"      Rotation: {transformResult.Rotation:F2} degrees, Mirrored: {transformResult.IsMirrored}");
        
        // For mirrored groups, we need a different approach based on which axes are mirrored
        if (transformResult.IsMirrored)
        {
            Transform combinedTransform = null;
            
            // Check which axes are mirrored
            if (transformResult.MirrorAxis != null)
            {
                double xComponent = Math.Abs(transformResult.MirrorAxis.X);
                double yComponent = Math.Abs(transformResult.MirrorAxis.Y);
                
                if (xComponent > 0.5 && yComponent > 0.5)
                {
                    // Both X and Y mirrored (180-degree rotation)
                    diagnosticLog.AppendLine($"      Applying 180-degree rotation (XY mirror)");
                    
                    // Create 180-degree rotation around Z-axis at reference origin
                    Transform rotTransform = Transform.CreateRotationAtPoint(
                        XYZ.BasisZ, 
                        Math.PI, // 180 degrees
                        refOrigin
                    );
                    
                    // Then translate
                    Transform translateTransform = Transform.CreateTranslation(translation);
                    combinedTransform = translateTransform.Multiply(rotTransform);
                }
                else if (yComponent > 0.5)
                {
                    // Y-axis mirror (mirror across XZ plane)
                    diagnosticLog.AppendLine($"      Applying Y-axis mirror");
                    
                    // Create mirror transform around XZ plane at reference origin
                    Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisY, refOrigin);
                    Transform mirrorTransform = Transform.CreateReflection(mirrorPlane);
                    
                    // Then translate
                    Transform translateTransform = Transform.CreateTranslation(translation);
                    combinedTransform = translateTransform.Multiply(mirrorTransform);
                }
                else
                {
                    // X-axis mirror (mirror across YZ plane) - default
                    diagnosticLog.AppendLine($"      Applying X-axis mirror");
                    
                    // Create mirror transform around YZ plane at reference origin
                    Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, refOrigin);
                    Transform mirrorTransform = Transform.CreateReflection(mirrorPlane);
                    
                    // Then translate
                    Transform translateTransform = Transform.CreateTranslation(translation);
                    combinedTransform = translateTransform.Multiply(mirrorTransform);
                }
            }
            else
            {
                // Fallback to X-axis mirror if MirrorAxis not set
                diagnosticLog.AppendLine($"      Applying default X-axis mirror");
                Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, refOrigin);
                Transform mirrorTransform = Transform.CreateReflection(mirrorPlane);
                Transform translateTransform = Transform.CreateTranslation(translation);
                combinedTransform = translateTransform.Multiply(mirrorTransform);
            }
            
            diagnosticLog.AppendLine($"      Final Transform Origin: X={combinedTransform.Origin.X:F3}, Y={combinedTransform.Origin.Y:F3}, Z={combinedTransform.Origin.Z:F3}");
            
            // Test the transform
            XYZ testPoint = refOrigin + new XYZ(1, 0, 0);
            XYZ transformed = combinedTransform.OfPoint(testPoint);
            diagnosticLog.AppendLine($"      Transform test: RefOrigin+(1,0,0) -> ({transformed.X:F3},{transformed.Y:F3},{transformed.Z:F3})");
            
            return combinedTransform;
        }
        else
        {
            // For non-mirrored groups, handle rotation and translation normally
            Transform transform = Transform.Identity;
            
            // Apply rotation if needed
            if (Math.Abs(transformResult.Rotation) > 0.001) // Tolerance for rotation
            {
                double angleRad = transformResult.Rotation * Math.PI / 180.0;
                
                // Create rotation around the reference origin
                Transform rotationTransform = Transform.CreateRotationAtPoint(
                    XYZ.BasisZ, 
                    angleRad, 
                    refOrigin
                );
                
                transform = transform.Multiply(rotationTransform);
                diagnosticLog.AppendLine($"      Applied rotation: {transformResult.Rotation:F2} degrees around reference origin");
            }
            
            // Apply translation
            Transform translationTransform = Transform.CreateTranslation(translation);
            transform = transform.Multiply(translationTransform);
            
            diagnosticLog.AppendLine($"      Final Transform Origin: X={transform.Origin.X:F3}, Y={transform.Origin.Y:F3}, Z={transform.Origin.Z:F3}");
            
            return transform;
        }
    }
}
