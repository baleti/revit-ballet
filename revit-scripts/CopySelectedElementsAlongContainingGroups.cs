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
    private StringBuilder diagnosticLog = new StringBuilder();
    
    // Configuration options
    private bool allowDuplicates = false; // Set to true to bypass duplicate checking
    private bool verboseDiagnostics = true; // Set to true for detailed diagnostics
    
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
            
            // Match elements to their containing groups using proper containment check
            Dictionary<Group, List<Element>> groupElementMap = new Dictionary<Group, List<Element>>();
            
            foreach (Group group in selectedGroups)
            {
                List<Element> containedElements = GetElementsContainedInGroup(group, selectedElements, doc);
                
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
            
            if (allowDuplicates)
            {
                results.AppendLine("NOTE: Duplicate checking is DISABLED");
            }
            
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
                    results.AppendLine($"Elements to copy: {elementsToCopy.Count}");
                    
                    // Update Comments parameter for original source elements
                    foreach (Element elem in elementsToCopy)
                    {
                        Parameter commentsParam = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                        if (commentsParam != null && !commentsParam.IsReadOnly)
                        {
                            commentsParam.Set($"{groupType.Name}");
                        }
                    }
                    
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
                    
                    // Clear diagnostics for new group type
                    diagnosticLog.Clear();
                    diagnosticLog.AppendLine("\n=== TRANSFORMATION DIAGNOSTICS ===");
                    diagnosticLog.AppendLine($"Reference Group Origin: {FormatPoint(refElements.GroupOrigin)}");
                    
                    // Copy elements to other instances
                    int copiedForThisGroup = 0;
                    
                    foreach (Element elem in allGroupInstances)
                    {
                        Group otherGroup = elem as Group;
                        if (otherGroup.Id == referenceGroup.Id) continue;
                        
                        XYZ otherOrigin = (otherGroup.Location as LocationPoint).Point;
                        
                        // Add diagnostic header for this target group
                        diagnosticLog.AppendLine($"\n--- Target Group: {otherGroup.Id} ---");
                        diagnosticLog.AppendLine($"Target Group Origin: {FormatPoint(otherOrigin)}");
                        
                        TransformResult transformResult = CalculateTransformation(refElements, otherGroup, doc);
                        
                        if (transformResult != null)
                        {
                            // Log transformation results
                            diagnosticLog.AppendLine($"Transformation Found:");
                            diagnosticLog.AppendLine($"  Translation: {FormatPoint(transformResult.Translation)}");
                            diagnosticLog.AppendLine($"  Rotation: {transformResult.Rotation:F2}°");
                            diagnosticLog.AppendLine($"  Mirrored: {transformResult.IsMirrored}");
                            diagnosticLog.AppendLine($"  Matching Elements: {transformResult.MatchingElements}");
                            
                            // Check if elements already exist at target location to avoid duplicates
                            bool elementsAlreadyExist = false;
                            
                            if (!allowDuplicates)
                            {
                                elementsAlreadyExist = CheckIfElementsExistAtTarget(
                                    elementsToCopy, transformResult, 
                                    refElements.GroupOrigin, otherOrigin, doc);
                            }
                            
                            if (elementsAlreadyExist)
                            {
                                diagnosticLog.AppendLine("  Elements already exist at target location - skipping");
                                
                                if (verboseDiagnostics)
                                {
                                    // Add diagnostic info about what was found
                                    diagnosticLog.AppendLine("  Diagnostic: Checking each element individually...");
                                    foreach (Element elemm in elementsToCopy)
                                    {
                                        bool exists = CheckIfSingleElementExistsAtTarget(
                                            elemm, transformResult, refElements.GroupOrigin, otherOrigin, doc);
                                        diagnosticLog.AppendLine($"    Element {elemm.Id}: {(exists ? "EXISTS" : "NOT FOUND")}");
                                    }
                                }
                                
                                continue;
                            }
                            
                            // Copy elements with transformation
                            foreach (Element elementToCopy in elementsToCopy)
                            {
                                try
                                {
                                    // Create the transformation
                                    Transform transform = CreateTransform(transformResult, 
                                        refElements.GroupOrigin,
                                        otherOrigin);
                                    
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
                                            commentsParam.Set($"{groupType.Name}");
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
                        else
                        {
                            diagnosticLog.AppendLine("  No transformation could be calculated");
                        }
                    }
                    
                    results.AppendLine($"Elements copied to other instances: {copiedForThisGroup}");
                }
                
                trans.Commit();
            }
            
            results.AppendLine("\n" + new string('=', 50));
            results.AppendLine($"\nTotal elements copied: {totalCopied}");
            
            // Add diagnostics to results
            results.AppendLine("\n" + diagnosticLog.ToString());
            
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
    
    // Improved containment check - using convex hull approach
    private List<Element> GetElementsContainedInGroup(Group group, List<Element> candidateElements, Document doc)
    {
        List<Element> containedElements = new List<Element>();
        
        // Get all member locations to create a containment boundary
        ICollection<ElementId> memberIds = group.GetMemberIds();
        List<XYZ> boundaryPoints = new List<XYZ>();
        
        // Collect all boundary points from group members
        foreach (ElementId id in memberIds)
        {
            Element member = doc.GetElement(id);
            if (member == null) continue;
            
            // Get points based on element location
            LocationPoint locPoint = member.Location as LocationPoint;
            LocationCurve locCurve = member.Location as LocationCurve;
            
            if (locPoint != null)
            {
                boundaryPoints.Add(locPoint.Point);
            }
            else if (locCurve != null)
            {
                Curve curve = locCurve.Curve;
                boundaryPoints.Add(curve.GetEndPoint(0));
                boundaryPoints.Add(curve.GetEndPoint(1));
                // Add some intermediate points for better accuracy
                for (int i = 1; i < 4; i++)
                {
                    boundaryPoints.Add(curve.Evaluate(i * 0.25, true));
                }
            }
            else
            {
                // For other elements, use bounding box corners
                BoundingBoxXYZ bb = member.get_BoundingBox(null);
                if (bb != null)
                {
                    boundaryPoints.Add(bb.Min);
                    boundaryPoints.Add(bb.Max);
                    boundaryPoints.Add(new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z));
                    boundaryPoints.Add(new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z));
                }
            }
        }
        
        if (boundaryPoints.Count < 3)
        {
            // Not enough points to form a boundary, fall back to bounding box
            return GetElementsInBoundingBox(group, candidateElements);
        }
        
        // For now, use an enhanced bounding box approach with tighter tolerance
        // A full convex hull implementation would be more complex
        
        // Find the extremes of the boundary points
        double minX = boundaryPoints.Min(p => p.X);
        double maxX = boundaryPoints.Max(p => p.X);
        double minY = boundaryPoints.Min(p => p.Y);
        double maxY = boundaryPoints.Max(p => p.Y);
        double minZ = boundaryPoints.Min(p => p.Z);
        double maxZ = boundaryPoints.Max(p => p.Z);
        
        // Create a tight bounding region with minimal tolerance
        double tolerance = 0.001; // About 0.3mm
        XYZ min = new XYZ(minX - tolerance, minY - tolerance, minZ - tolerance);
        XYZ max = new XYZ(maxX + tolerance, maxY + tolerance, maxZ + tolerance);
        
        // Check each candidate element
        foreach (Element elem in candidateElements)
        {
            if (IsElementInBoundingRegion(elem, min, max, boundaryPoints))
            {
                containedElements.Add(elem);
            }
        }
        
        return containedElements;
    }
    
    // Enhanced containment check considering actual boundary points
    private bool IsElementInBoundingRegion(Element elem, XYZ min, XYZ max, List<XYZ> boundaryPoints)
    {
        // First do a quick bounding box check
        LocationPoint locPoint = elem.Location as LocationPoint;
        LocationCurve locCurve = elem.Location as LocationCurve;
        
        if (locPoint != null)
        {
            XYZ point = locPoint.Point;
            if (!IsPointInBoundingBox(point, min, max))
                return false;
                
            // Additional check: ensure point is reasonably close to some boundary points
            // This helps exclude elements that are in the bounding box but not part of the group
            double minDistance = boundaryPoints.Min(p => p.DistanceTo(point));
            return minDistance < 10.0; // Within 10 feet/units of some group element
        }
        else if (locCurve != null)
        {
            XYZ start = locCurve.Curve.GetEndPoint(0);
            XYZ end = locCurve.Curve.GetEndPoint(1);
            
            // Both endpoints must be in bounds
            if (!IsPointInBoundingBox(start, min, max) || !IsPointInBoundingBox(end, min, max))
                return false;
                
            // Check proximity to boundary
            double minStartDist = boundaryPoints.Min(p => p.DistanceTo(start));
            double minEndDist = boundaryPoints.Min(p => p.DistanceTo(end));
            return Math.Max(minStartDist, minEndDist) < 10.0;
        }
        else
        {
            // For other elements, check their bounding box
            BoundingBoxXYZ elemBB = elem.get_BoundingBox(null);
            if (elemBB != null)
            {
                if (!IsPointInBoundingBox(elemBB.Min, min, max) || !IsPointInBoundingBox(elemBB.Max, min, max))
                    return false;
                    
                XYZ center = (elemBB.Min + elemBB.Max) * 0.5;
                double minDistance = boundaryPoints.Min(p => p.DistanceTo(center));
                return minDistance < 10.0;
            }
        }
        
        return false;
    }
    
    // Fallback bounding box method with tighter tolerance
    private List<Element> GetElementsInBoundingBox(Group group, List<Element> candidateElements)
    {
        List<Element> containedElements = new List<Element>();
        
        BoundingBoxXYZ groupBB = group.get_BoundingBox(null);
        if (groupBB == null) return containedElements;
        
        // Use very small tolerance
        double tolerance = 0.01; // About 3mm
        XYZ min = groupBB.Min - new XYZ(tolerance, tolerance, tolerance);
        XYZ max = groupBB.Max + new XYZ(tolerance, tolerance, tolerance);
        
        foreach (Element elem in candidateElements)
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
        
        return containedElements;
    }
    
    private bool IsPointInBoundingBox(XYZ point, XYZ min, XYZ max)
    {
        return point.X >= min.X && point.X <= max.X &&
               point.Y >= min.Y && point.Y <= max.Y &&
               point.Z >= min.Z && point.Z <= max.Z;
    }
    
    // Check if elements already exist at target location
    private bool CheckIfElementsExistAtTarget(List<Element> elementsToCopy, 
        TransformResult transformResult, XYZ refOrigin, XYZ targetOrigin, Document doc)
    {
        // Create the transformation
        Transform transform = CreateTransform(transformResult, refOrigin, targetOrigin);
        
        // Count how many potential duplicates we find
        int duplicateCount = 0;
        
        foreach (Element elem in elementsToCopy)
        {
            // Get element location(s)
            List<XYZ> testPoints = new List<XYZ>();
            
            LocationPoint locPoint = elem.Location as LocationPoint;
            LocationCurve locCurve = elem.Location as LocationCurve;
            
            if (locPoint != null)
            {
                testPoints.Add(locPoint.Point);
            }
            else if (locCurve != null)
            {
                // For curves, check both endpoints
                testPoints.Add(locCurve.Curve.GetEndPoint(0));
                testPoints.Add(locCurve.Curve.GetEndPoint(1));
            }
            
            if (testPoints.Count == 0) continue;
            
            bool foundDuplicate = false;
            
            foreach (XYZ testPoint in testPoints)
            {
                // Transform the test point
                XYZ targetPoint = transform.OfPoint(testPoint);
                
                // Check if an element of same type exists at target location
                // Using a very small search radius for precision
                double searchRadius = 0.01; // About 3mm
                BoundingBoxXYZ searchBox = new BoundingBoxXYZ();
                searchBox.Min = targetPoint - new XYZ(searchRadius, searchRadius, searchRadius);
                searchBox.Max = targetPoint + new XYZ(searchRadius, searchRadius, searchRadius);
                
                // Create a bounding box filter
                Outline outline = new Outline(searchBox.Min, searchBox.Max);
                BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);
                
                // Filter for same category
                ElementCategoryFilter categoryFilter = new ElementCategoryFilter(elem.Category.Id);
                
                // Combine filters
                LogicalAndFilter andFilter = new LogicalAndFilter(bbFilter, categoryFilter);
                
                // Search for existing elements
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Element> existingElements = collector
                    .WhereElementIsNotElementType()
                    .WherePasses(andFilter)
                    .ToList();
                
                // Check if any existing element matches our element type and is not the source element
                foreach (Element existing in existingElements)
                {
                    if (existing.Id == elem.Id) continue; // Skip the source element itself
                    
                    if (existing.GetTypeId() == elem.GetTypeId())
                    {
                        // Additional check for curves - verify both endpoints match
                        if (locCurve != null && existing.Location is LocationCurve)
                        {
                            LocationCurve existingCurve = existing.Location as LocationCurve;
                            XYZ existingStart = existingCurve.Curve.GetEndPoint(0);
                            XYZ existingEnd = existingCurve.Curve.GetEndPoint(1);
                            
                            // Transform both endpoints of source curve
                            XYZ transformedStart = transform.OfPoint(locCurve.Curve.GetEndPoint(0));
                            XYZ transformedEnd = transform.OfPoint(locCurve.Curve.GetEndPoint(1));
                            
                            // Check if endpoints match (in either order)
                            bool endpointsMatch = 
                                (existingStart.IsAlmostEqualTo(transformedStart, 0.01) && 
                                 existingEnd.IsAlmostEqualTo(transformedEnd, 0.01)) ||
                                (existingStart.IsAlmostEqualTo(transformedEnd, 0.01) && 
                                 existingEnd.IsAlmostEqualTo(transformedStart, 0.01));
                            
                            if (endpointsMatch)
                            {
                                foundDuplicate = true;
                                break;
                            }
                        }
                        else
                        {
                            foundDuplicate = true;
                            break;
                        }
                    }
                }
                
                if (foundDuplicate) break;
            }
            
            if (foundDuplicate)
            {
                duplicateCount++;
            }
        }
        
        // Only skip if ALL elements already exist (not just some)
        // This handles cases where some elements might be shared between groups
        return duplicateCount == elementsToCopy.Count;
    }
    
    private Transform CreateTransform(TransformResult transformResult, XYZ refOrigin, XYZ targetOrigin)
    {
        diagnosticLog.AppendLine($"\n  Creating Transform:");
        diagnosticLog.AppendLine($"    Ref Origin: {FormatPoint(refOrigin)}");
        diagnosticLog.AppendLine($"    Target Origin: {FormatPoint(targetOrigin)}");
        diagnosticLog.AppendLine($"    Rotation: {transformResult.Rotation:F2}°, Mirrored: {transformResult.IsMirrored}");
        
        // For mirrored transformations with no rotation
        if (transformResult.IsMirrored && Math.Abs(transformResult.Rotation) < 0.01)
        {
            diagnosticLog.AppendLine($"    Using mirror-only transformation");
            
            // Determine which axis is mirrored based on the translation
            XYZ midpoint = (refOrigin + targetOrigin) * 0.5;
            
            // Check which coordinates changed
            double xDiff = Math.Abs(refOrigin.X - targetOrigin.X);
            double yDiff = Math.Abs(refOrigin.Y - targetOrigin.Y);
            double zDiff = Math.Abs(refOrigin.Z - targetOrigin.Z);
            
            Transform mirror;
            
            if (xDiff > yDiff && xDiff > zDiff)
            {
                // X coordinates differ most - mirror about Y-Z plane
                diagnosticLog.AppendLine($"    Mirror about Y-Z plane (X-axis mirror)");
                Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, midpoint);
                mirror = Transform.CreateReflection(mirrorPlane);
            }
            else if (yDiff > xDiff && yDiff > zDiff)
            {
                // Y coordinates differ most - mirror about X-Z plane
                diagnosticLog.AppendLine($"    Mirror about X-Z plane (Y-axis mirror)");
                Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisY, midpoint);
                mirror = Transform.CreateReflection(mirrorPlane);
            }
            else
            {
                // Z coordinates differ most - mirror about X-Y plane
                diagnosticLog.AppendLine($"    Mirror about X-Y plane (Z-axis mirror)");
                Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, midpoint);
                mirror = Transform.CreateReflection(mirrorPlane);
            }
            
            diagnosticLog.AppendLine($"    Mirror plane at: {FormatPoint(midpoint)}");
            
            return mirror;
        }
        
        // For complex transformations involving both rotation and mirroring
        if (Math.Abs(Math.Abs(transformResult.Rotation) - 180.0) < 1.0 && transformResult.IsMirrored)
        {
            diagnosticLog.AppendLine($"    Using special 180° + mirror transformation");
            
            // This is equivalent to a mirror about Y axis
            Transform t = Transform.Identity;
            
            // Set the basis vectors
            t.BasisX = -XYZ.BasisX;  // (-1, 0, 0)
            t.BasisY = XYZ.BasisY;   // (0, 1, 0) 
            t.BasisZ = XYZ.BasisZ;   // (0, 0, 1)
            
            // Set the translation
            t.Origin = new XYZ(
                2 * ((refOrigin.X + targetOrigin.X) / 2) - refOrigin.X,
                targetOrigin.Y - refOrigin.Y,
                targetOrigin.Z - refOrigin.Z
            );
            
            diagnosticLog.AppendLine($"    Transform Origin: {FormatPoint(t.Origin)}");
            diagnosticLog.AppendLine($"    Transform BasisX: {FormatPoint(t.BasisX)}");
            diagnosticLog.AppendLine($"    Transform BasisY: {FormatPoint(t.BasisY)}");
            
            return t;
        }
        
        // Standard case: build transformation step by step
        diagnosticLog.AppendLine($"    Using standard transformation");
        
        Transform transform = Transform.Identity;
        
        // Move to origin
        transform = transform.Multiply(Transform.CreateTranslation(-refOrigin));
        
        // Apply rotation
        if (Math.Abs(transformResult.Rotation) > 0.01)
        {
            double radians = transformResult.Rotation * Math.PI / 180;
            transform = transform.Multiply(Transform.CreateRotation(XYZ.BasisZ, radians));
        }
        
        // Apply mirror if needed
        if (transformResult.IsMirrored)
        {
            Plane mirrorPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, XYZ.Zero);
            transform = transform.Multiply(Transform.CreateReflection(mirrorPlane));
        }
        
        // Move to target
        transform = transform.Multiply(Transform.CreateTranslation(targetOrigin));
        
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
            diagnosticLog.AppendLine($"\n  Calculating Transformation:");
            
            // Get corresponding elements in other group
            List<ElementInfo> otherElements = GetCorrespondingElements(refElements, otherGroup, doc);
            
            // We need at least 2 matching elements to calculate transformation
            if (otherElements == null || otherElements.Count < 2) 
            {
                diagnosticLog.AppendLine($"    Not enough matching elements: {otherElements?.Count ?? 0}");
                return null;
            }
            
            TransformResult result = new TransformResult();
            result.MatchingElements = otherElements.Count;
            
            // Calculate translation (difference in group origins)
            XYZ otherOrigin = (otherGroup.Location as LocationPoint).Point;
            result.Translation = otherOrigin - refElements.GroupOrigin;
            
            diagnosticLog.AppendLine($"    Found {otherElements.Count} matching elements");
            diagnosticLog.AppendLine($"    Translation: {FormatPoint(result.Translation)}");
            
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
                diagnosticLog.AppendLine($"    Could not find two matching element pairs");
                
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
                    
                    diagnosticLog.AppendLine($"    Using single element for rotation: {result.Rotation:F2}°");
                    
                    return result;
                }
                return null;
            }
            
            // Log element details
            diagnosticLog.AppendLine($"\n    Element 1 (Ref):");
            diagnosticLog.AppendLine($"      Key: {ref1.UniqueKey}");
            diagnosticLog.AppendLine($"      Point1: {FormatPoint(ref1.Point1)}");
            diagnosticLog.AppendLine($"      Point2: {FormatPoint(ref1.Point2)}");
            
            diagnosticLog.AppendLine($"    Element 1 (Other):");
            diagnosticLog.AppendLine($"      Point1: {FormatPoint(other1.Point1)}");
            diagnosticLog.AppendLine($"      Point2: {FormatPoint(other1.Point2)}");
            
            diagnosticLog.AppendLine($"\n    Element 2 (Ref):");
            diagnosticLog.AppendLine($"      Key: {ref2.UniqueKey}");
            diagnosticLog.AppendLine($"      Point1: {FormatPoint(ref2.Point1)}");
            diagnosticLog.AppendLine($"      Point2: {FormatPoint(ref2.Point2)}");
            
            diagnosticLog.AppendLine($"    Element 2 (Other):");
            diagnosticLog.AppendLine($"      Point1: {FormatPoint(other2.Point1)}");
            diagnosticLog.AppendLine($"      Point2: {FormatPoint(other2.Point2)}");
            
            // Calculate vectors in reference group
            XYZ refVector1 = (ref1.Point2 - ref1.Point1).Normalize();
            XYZ refVector2 = (ref2.Point2 - ref2.Point1).Normalize();
            
            // Calculate vectors in other group  
            XYZ otherVector1 = (other1.Point2 - other1.Point1).Normalize();
            XYZ otherVector2 = (other2.Point2 - other2.Point1).Normalize();
            
            diagnosticLog.AppendLine($"\n    Ref Vector 1: {FormatPoint(refVector1)}");
            diagnosticLog.AppendLine($"    Other Vector 1: {FormatPoint(otherVector1)}");
            
            // For mirrored groups, the matching elements might have reversed directions
            // Check if vectors are opposite
            double dotProduct = refVector1.DotProduct(otherVector1);
            bool vectorsReversed = dotProduct < -0.9; // Nearly opposite
            
            if (vectorsReversed)
            {
                diagnosticLog.AppendLine($"    Vectors are reversed (dot product: {dotProduct:F4})");
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
            
            diagnosticLog.AppendLine($"    Dot Product: {dotProduct:F4}");
            diagnosticLog.AppendLine($"    Cross Product Z: {crossProduct.Z:F4}");
            diagnosticLog.AppendLine($"    Rotation: {result.Rotation:F2}°");
            
            // Simplified mirror detection for axis-aligned mirrors
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
            
            diagnosticLog.AppendLine($"\n    Mirror Detection:");
            diagnosticLog.AppendLine($"    Ref Pos 1 (relative): {FormatPoint(refPos1)}");
            diagnosticLog.AppendLine($"    Other Pos 1 (relative): {FormatPoint(otherPos1)}");
            
            // For simple mirror across Y-Z plane (X-axis mirror)
            // X coordinates should be negated, Y and Z should be same
            bool xMirrored = Math.Abs(refPos1.X + otherPos1.X) < 0.1 && 
                             Math.Abs(refPos1.Y - otherPos1.Y) < 0.1 &&
                             Math.Abs(refPos1.Z - otherPos1.Z) < 0.1;
            
            // For simple mirror across X-Z plane (Y-axis mirror)  
            bool yMirrored = Math.Abs(refPos1.X - otherPos1.X) < 0.1 && 
                             Math.Abs(refPos1.Y + otherPos1.Y) < 0.1 &&
                             Math.Abs(refPos1.Z - otherPos1.Z) < 0.1;
            
            result.IsMirrored = xMirrored || yMirrored || vectorsReversed;
            
            diagnosticLog.AppendLine($"    X-Mirror Check: {xMirrored}");
            diagnosticLog.AppendLine($"    Y-Mirror Check: {yMirrored}");
            diagnosticLog.AppendLine($"    Vectors Reversed: {vectorsReversed}");
            diagnosticLog.AppendLine($"    Final Mirror Detection: {result.IsMirrored}");
            
            // Scale (typically 1.0 for groups)
            result.Scale = 1.0;
            
            diagnosticLog.AppendLine($"\n    Final Transformation:");
            diagnosticLog.AppendLine($"    Rotation: {result.Rotation:F2}°");
            diagnosticLog.AppendLine($"    Mirrored: {result.IsMirrored}");
            diagnosticLog.AppendLine($"    Scale: {result.Scale}");
            
            return result;
        }
        catch (Exception ex)
        {
            diagnosticLog.AppendLine($"    Exception in CalculateTransformation: {ex.Message}");
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
    
    // Check if a single element exists at target location
    private bool CheckIfSingleElementExistsAtTarget(Element elem, 
        TransformResult transformResult, XYZ refOrigin, XYZ targetOrigin, Document doc)
    {
        // Create the transformation
        Transform transform = CreateTransform(transformResult, refOrigin, targetOrigin);
        
        // Get element location(s)
        List<XYZ> testPoints = new List<XYZ>();
        
        LocationPoint locPoint = elem.Location as LocationPoint;
        LocationCurve locCurve = elem.Location as LocationCurve;
        
        if (locPoint != null)
        {
            testPoints.Add(locPoint.Point);
        }
        else if (locCurve != null)
        {
            // For curves, check both endpoints
            testPoints.Add(locCurve.Curve.GetEndPoint(0));
            testPoints.Add(locCurve.Curve.GetEndPoint(1));
        }
        
        if (testPoints.Count == 0) return false;
        
        foreach (XYZ testPoint in testPoints)
        {
            // Transform the test point
            XYZ targetPoint = transform.OfPoint(testPoint);
            
            // Check if an element of same type exists at target location
            double searchRadius = 0.01; // About 3mm
            BoundingBoxXYZ searchBox = new BoundingBoxXYZ();
            searchBox.Min = targetPoint - new XYZ(searchRadius, searchRadius, searchRadius);
            searchBox.Max = targetPoint + new XYZ(searchRadius, searchRadius, searchRadius);
            
            // Create filters
            Outline outline = new Outline(searchBox.Min, searchBox.Max);
            BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);
            ElementCategoryFilter categoryFilter = new ElementCategoryFilter(elem.Category.Id);
            LogicalAndFilter andFilter = new LogicalAndFilter(bbFilter, categoryFilter);
            
            // Search for existing elements
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            IList<Element> existingElements = collector
                .WhereElementIsNotElementType()
                .WherePasses(andFilter)
                .ToList();
            
            // Check if any existing element matches
            foreach (Element existing in existingElements)
            {
                if (existing.Id == elem.Id) continue;
                
                if (existing.GetTypeId() == elem.GetTypeId())
                {
                    // For curves, verify endpoints match
                    if (locCurve != null && existing.Location is LocationCurve)
                    {
                        LocationCurve existingCurve = existing.Location as LocationCurve;
                        XYZ existingStart = existingCurve.Curve.GetEndPoint(0);
                        XYZ existingEnd = existingCurve.Curve.GetEndPoint(1);
                        
                        XYZ transformedStart = transform.OfPoint(locCurve.Curve.GetEndPoint(0));
                        XYZ transformedEnd = transform.OfPoint(locCurve.Curve.GetEndPoint(1));
                        
                        bool endpointsMatch = 
                            (existingStart.IsAlmostEqualTo(transformedStart, 0.01) && 
                             existingEnd.IsAlmostEqualTo(transformedEnd, 0.01)) ||
                            (existingStart.IsAlmostEqualTo(transformedEnd, 0.01) && 
                             existingEnd.IsAlmostEqualTo(transformedStart, 0.01));
                        
                        if (endpointsMatch) return true;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
        }
        
        return false;
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
