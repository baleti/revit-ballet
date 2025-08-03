using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

[Transaction(TransactionMode.Manual)]
public class CopySelectedElementsAlongContainingGroups : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        try
        {
            // Get selected elements
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            
            if (selectedIds.Count == 0)
            {
                message = "Please select elements and groups first";
                return Result.Failed;
            }
            
            // Separate groups and elements
            List<Group> selectedGroups = new List<Group>();
            List<Element> selectedElements = new List<Element>();
            
            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is Group)
                {
                    selectedGroups.Add(elem as Group);
                }
                else if (elem != null && !(elem is ElementType))
                {
                    selectedElements.Add(elem);
                }
            }
            
            if (selectedGroups.Count == 0)
            {
                message = "No groups found in selection. Please select at least one group.";
                return Result.Failed;
            }
            
            if (selectedElements.Count == 0)
            {
                message = "No elements found in selection. Please select elements to copy.";
                return Result.Failed;
            }
            
            // Match elements to their containing groups
            Dictionary<Group, List<Element>> groupElementMap = new Dictionary<Group, List<Element>>();
            
            foreach (Group group in selectedGroups)
            {
                BoundingBoxXYZ groupBB = group.get_BoundingBox(null);
                if (groupBB == null) continue;
                
                // Expand bounding box slightly to account for tolerance
                XYZ min = groupBB.Min - new XYZ(0.1, 0.1, 0.1);
                XYZ max = groupBB.Max + new XYZ(0.1, 0.1, 0.1);
                
                List<Element> containedElements = new List<Element>();
                
                foreach (Element elem in selectedElements)
                {
                    LocationPoint locPoint = elem.Location as LocationPoint;
                    LocationCurve locCurve = elem.Location as LocationCurve;
                    
                    bool isContained = false;
                    
                    if (locPoint != null)
                    {
                        XYZ point = locPoint.Point;
                        isContained = IsPointInBoundingBox(point, min, max);
                    }
                    else if (locCurve != null)
                    {
                        // Check if both endpoints are within bounding box
                        XYZ start = locCurve.Curve.GetEndPoint(0);
                        XYZ end = locCurve.Curve.GetEndPoint(1);
                        isContained = IsPointInBoundingBox(start, min, max) && 
                                     IsPointInBoundingBox(end, min, max);
                    }
                    else
                    {
                        // For other elements, check their bounding box
                        BoundingBoxXYZ elemBB = elem.get_BoundingBox(null);
                        if (elemBB != null)
                        {
                            isContained = IsPointInBoundingBox(elemBB.Min, min, max) && 
                                         IsPointInBoundingBox(elemBB.Max, min, max);
                        }
                    }
                    
                    if (isContained)
                    {
                        containedElements.Add(elem);
                    }
                }
                
                if (containedElements.Count > 0)
                {
                    groupElementMap[group] = containedElements;
                }
            }
            
            if (groupElementMap.Count == 0)
            {
                message = "No selected elements are contained within the selected groups";
                return Result.Failed;
            }
            
            // Process each group and copy elements
            StringBuilder results = new StringBuilder();
            results.AppendLine("Copy Elements Following Groups - Results");
            results.AppendLine(new string('=', 50));
            
            int totalCopied = 0;
            
            using (Transaction trans = new Transaction(doc, "Copy Elements Following Groups"))
            {
                trans.Start();
                
                foreach (var kvp in groupElementMap)
                {
                    Group referenceGroup = kvp.Key;
                    List<Element> elementsToCopy = kvp.Value;
                    
                    // Get the group type
                    GroupType groupType = doc.GetElement(referenceGroup.GetTypeId()) as GroupType;
                    
                    results.AppendLine($"\nGroup Type: {groupType.Name}");
                    results.AppendLine($"Reference Group ID: {referenceGroup.Id}");
                    results.AppendLine($"Elements to copy: {elementsToCopy.Count}");
                    
                    // Get all instances of this group type
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    IList<Element> allGroupInstances = collector
                        .OfClass(typeof(Group))
                        .WhereElementIsNotElementType()
                        .Where(g => g.GetTypeId() == groupType.Id)
                        .ToList();
                    
                    results.AppendLine($"Total group instances: {allGroupInstances.Count}");
                    
                    // Get reference elements for transformation calculation
                    ReferenceElements refElements = GetReferenceElements(referenceGroup, doc);
                    
                    if (refElements == null || refElements.Elements.Count < 2)
                    {
                        results.AppendLine("WARNING: Could not determine transformation for this group type");
                        continue;
                    }
                    
                    // Copy elements to other instances
                    int copiedForThisGroup = 0;
                    
                    foreach (Element elem in allGroupInstances)
                    {
                        Group otherGroup = elem as Group;
                        if (otherGroup.Id == referenceGroup.Id) continue;
                        
                        TransformResult transformResult = CalculateTransformation(refElements, otherGroup, doc);
                        
                        if (transformResult != null)
                        {
                            // Copy elements with transformation
                            foreach (Element elementToCopy in elementsToCopy)
                            {
                                try
                                {
                                    // Create the transformation
                                    Transform transform = CreateTransform(transformResult, 
                                        (referenceGroup.Location as LocationPoint).Point,
                                        (otherGroup.Location as LocationPoint).Point);
                                    
                                    // Copy the element
                                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                        doc, 
                                        new List<ElementId> { elementToCopy.Id },
                                        doc,
                                        transform,
                                        null);
                                    
                                    // Update Comments parameter for copied elements
                                    foreach (ElementId copiedId in copiedIds)
                                    {
                                        Element copiedElem = doc.GetElement(copiedId);
                                        Parameter commentsParam = copiedElem.get_Parameter(
                                            BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                        
                                        if (commentsParam != null && !commentsParam.IsReadOnly)
                                        {
                                            commentsParam.Set($"Copied to {groupType.Name}");
                                        }
                                        
                                        copiedForThisGroup++;
                                        totalCopied++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    results.AppendLine($"  Error copying element {elementToCopy.Id}: {ex.Message}");
                                }
                            }
                        }
                    }
                    
                    results.AppendLine($"Elements copied to other instances: {copiedForThisGroup}");
                }
                
                trans.Commit();
            }
            
            results.AppendLine("\n" + new string('=', 50));
            results.AppendLine($"\nTotal elements copied: {totalCopied}");
            
            // Display results
            TaskDialog dlg = new TaskDialog("Copy Elements Following Groups");
            dlg.MainContent = results.ToString();
            dlg.Show();
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
    
    private bool IsPointInBoundingBox(XYZ point, XYZ min, XYZ max)
    {
        return point.X >= min.X && point.X <= max.X &&
               point.Y >= min.Y && point.Y <= max.Y &&
               point.Z >= min.Z && point.Z <= max.Z;
    }
    
    private Transform CreateTransform(TransformResult transformResult, XYZ refOrigin, XYZ targetOrigin)
    {
        Transform transform = Transform.Identity;
        
        // Apply translation
        transform = transform.Multiply(Transform.CreateTranslation(transformResult.Translation));
        
        // Apply rotation around target origin
        if (Math.Abs(transformResult.Rotation) > 0.01)
        {
            double radians = transformResult.Rotation * Math.PI / 180;
            XYZ axis = XYZ.BasisZ;
            transform = transform.Multiply(Transform.CreateRotationAtPoint(axis, radians, targetOrigin));
        }
        
        // Apply mirroring if needed
        if (transformResult.IsMirrored)
        {
            // Create mirror plane through target origin
            Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, targetOrigin);
            transform = transform.Multiply(Transform.CreateReflection(mirrorPlane));
        }
        
        return transform;
    }
    
    private ReferenceElements GetReferenceElements(Group group, Document doc)
    {
        ReferenceElements refElements = new ReferenceElements();
        refElements.GroupOrigin = (group.Location as LocationPoint).Point;
        
        // Get group members
        ICollection<ElementId> memberIds = group.GetMemberIds();
        
        // Check if this group has excluded members
        string groupName = group.Name;
        bool hasExcludedMembers = groupName != null && groupName.Contains("(members excluded)");
        
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
        refElements.HasExcludedMembers = hasExcludedMembers;
        return refElements;
    }
    
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
            
            // Add wall height (but NOT area or volume as they change with intersections)
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
    
    private TransformResult CalculateTransformation(ReferenceElements refElements, Group otherGroup, Document doc)
    {
        try
        {
            // Get corresponding elements in other group
            List<ElementInfo> otherElements = GetCorrespondingElements(refElements, otherGroup, doc);
            
            // We need at least 2 matching elements to calculate transformation
            if (otherElements == null || otherElements.Count < 2) return null;
            
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
                    result.Scale = 1.0;
                    return result;
                }
                return null;
            }
            
            // Calculate vectors in reference group
            XYZ refVector1 = (ref1.Point2 - ref1.Point1).Normalize();
            XYZ refVector2 = (ref2.Point2 - ref2.Point1).Normalize();
            
            // Calculate vectors in other group  
            XYZ otherVector1 = (other1.Point2 - other1.Point1).Normalize();
            XYZ otherVector2 = (other2.Point2 - other2.Point1).Normalize();
            
            // Calculate rotation angle from first element
            double dotProduct = refVector1.DotProduct(otherVector1);
            double angle = Math.Acos(Math.Max(-1.0, Math.Min(1.0, dotProduct)));
            
            // Determine rotation direction
            XYZ crossProduct = refVector1.CrossProduct(otherVector1);
            if (crossProduct.Z < 0) angle = -angle;
            
            result.Rotation = angle * 180 / Math.PI;
            
            // Mirror detection: Check the spatial arrangement of elements
            // Get midpoints of the elements
            XYZ refMid1 = (ref1.Point1 + ref1.Point2) * 0.5;
            XYZ refMid2 = (ref2.Point1 + ref2.Point2) * 0.5;
            XYZ otherMid1 = (other1.Point1 + other1.Point2) * 0.5;
            XYZ otherMid2 = (other2.Point1 + other2.Point2) * 0.5;
            
            // Calculate relative positions from group origins
            XYZ refPos1 = refMid1 - refElements.GroupOrigin;
            XYZ refPos2 = refMid2 - refElements.GroupOrigin;
            XYZ otherPos1 = otherMid1 - otherOrigin;
            XYZ otherPos2 = otherMid2 - otherOrigin;
            
            // Create a vector between the two element centers
            XYZ refBetween = refPos2 - refPos1;
            XYZ otherBetween = otherPos2 - otherPos1;
            
            // Remove rotation effect to isolate mirror detection
            // Rotate the other vector back by the detected angle
            double radAngle = -angle * Math.PI / 180;
            double cos = Math.Cos(radAngle);
            double sin = Math.Sin(radAngle);
            
            // Apply 2D rotation (assuming Z is up)
            XYZ rotatedOtherBetween = new XYZ(
                otherBetween.X * cos - otherBetween.Y * sin,
                otherBetween.X * sin + otherBetween.Y * cos,
                otherBetween.Z
            );
            
            // Now check if the arrangement is mirrored
            // The vectors should be opposite if mirrored
            double dotBetween = refBetween.Normalize().DotProduct(rotatedOtherBetween.Normalize());
            
            // If dot product is negative, the arrangement is flipped
            result.IsMirrored = dotBetween < -0.5;
            
            // Additional check using determinant method
            if (!result.IsMirrored && Math.Abs(angle) < 1.0) // Only for non-rotated cases
            {
                // Check if the handedness changed using cross product
                XYZ refNormal = refVector1.CrossProduct(refBetween);
                XYZ otherNormal = otherVector1.CrossProduct(otherBetween);
                
                // If normals point in opposite directions, it's mirrored
                double normalDot = refNormal.Normalize().DotProduct(otherNormal.Normalize());
                result.IsMirrored = normalDot < -0.5;
            }
            
            // Scale (typically 1.0 for groups)
            result.Scale = 1.0;
            
            return result;
        }
        catch
        {
            return null;
        }
    }
    
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
    
    private string FormatPoint(XYZ point)
    {
        if (point == null) return "null";
        return $"({point.X:F3}, {point.Y:F3}, {point.Z:F3})";
    }
    
    private class ReferenceElements
    {
        public XYZ GroupOrigin { get; set; }
        public List<ElementInfo> Elements { get; set; } = new List<ElementInfo>();
        public bool HasExcludedMembers { get; set; }
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
        public double Scale { get; set; }
        public int MatchingElements { get; set; }
    }
}
