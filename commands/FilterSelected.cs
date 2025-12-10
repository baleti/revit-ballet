using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public static class ElementDataHelper
{
    public static List<Dictionary<string, object>> GetElementData(UIDocument uiDoc, bool selectedOnly = false, bool includeParameters = false)
    {
        Document doc = uiDoc.Document;
        var elementData = new List<Dictionary<string, object>>();

        // Get all scope boxes in the document upfront
        var scopeBoxes = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
            .WhereElementIsNotElementType()
            .Cast<Element>()
            .Where(e => e.Name != null)
            .ToList();

        // Also collect scope boxes from linked models if SelectInLinksMode is enabled
        var linkedScopeBoxes = new List<Tuple<Element, RevitLinkInstance>>();
        if (SelectInLinksMode.IsEnabled())
        {
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => link.GetLinkDocument() != null)
                .ToList();

            foreach (var linkInstance in linkInstances)
            {
                try
                {
                    Document linkedDoc = linkInstance.GetLinkDocument();
                    if (linkedDoc != null)
                    {
                        var linkedSBs = new FilteredElementCollector(linkedDoc)
                            .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                            .WhereElementIsNotElementType()
                            .Cast<Element>()
                            .Where(e => e.Name != null)
                            .ToList();

                        foreach (var sb in linkedSBs)
                        {
                            linkedScopeBoxes.Add(new Tuple<Element, RevitLinkInstance>(sb, linkInstance));
                        }
                    }
                }
                catch { /* Skip problematic links */ }
            }
        }

        if (selectedOnly)
        {
            // Handle both regular elements and linked elements (via References)
            var selectedIds = uiDoc.GetSelectionIds();
            var selectedRefs = uiDoc.GetReferences();

            if (!selectedIds.Any() && !selectedRefs.Any())
                throw new InvalidOperationException("No elements are selected.");

            // Keep track of processed elements to avoid duplicates
            var processedElements = new HashSet<ElementId>();

            // Process regular elements from current document
            foreach (var id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    processedElements.Add(id);
                    var data = GetElementDataDictionary(element, doc, null, null, null, includeParameters, scopeBoxes, linkedScopeBoxes);
                    elementData.Add(data);
                }
            }

            // Process linked elements via References
            foreach (var reference in selectedRefs)
            {
                try
                {
                    // Get the linked instance
                    var linkedInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkedInstance != null)
                    {
                        Document linkedDoc = linkedInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            // Get the actual element in the linked document
                            var linkedElementId = reference.LinkedElementId;
                            if (linkedElementId != ElementId.InvalidElementId)
                            {
                                Element linkedElement = linkedDoc.GetElement(linkedElementId);
                                if (linkedElement != null)
                                {
                                    // Store the linked instance and element ID for later reference creation
                                    var data = GetElementDataDictionary(linkedElement, linkedDoc, linkedInstance.Name, linkedInstance, linkedElement.Id, includeParameters, scopeBoxes, linkedScopeBoxes);
                                    elementData.Add(data);
                                }
                            }
                        }
                    }
                    else
                    {
                        // This might be a regular element selected via reference
                        // Only process if we haven't already processed it via selectedIds
                        if (!processedElements.Contains(reference.ElementId))
                        {
                            Element element = doc.GetElement(reference);
                            if (element != null)
                            {
                                processedElements.Add(reference.ElementId);
                                var data = GetElementDataDictionary(element, doc, null, null, null, includeParameters, scopeBoxes, linkedScopeBoxes);
                                elementData.Add(data);
                            }
                        }
                    }
                }
                catch { /* Skip problematic references */ }
            }
        }
        else
        {
            // Get elements from active view (current document only)
            var elementIds = new FilteredElementCollector(doc, doc.ActiveView.Id).ToElementIds();
            foreach (var id in elementIds)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    var data = GetElementDataDictionary(element, doc, null, null, null, includeParameters, scopeBoxes, linkedScopeBoxes);
                    elementData.Add(data);
                }
            }
        }

        return elementData;
    }

    private static Dictionary<string, object> GetElementDataDictionary(Element element, Document elementDoc, string linkName, RevitLinkInstance linkInstance, ElementId linkedElementId, bool includeParameters, List<Element> scopeBoxes, List<Tuple<Element, RevitLinkInstance>> linkedScopeBoxes)
    {
        string groupName = string.Empty;
        if (element.GroupId != null && element.GroupId != ElementId.InvalidElementId && element.GroupId.AsLong() != -1)
        {
            if (elementDoc.GetElement(element.GroupId) is Group g)
                groupName = g.Name;
        }

        string ownerViewName = string.Empty;
        if (element.OwnerViewId != null && element.OwnerViewId != ElementId.InvalidElementId)
        {
            if (elementDoc.GetElement(element.OwnerViewId) is View v)
                ownerViewName = v.Name;
        }

        var data = new Dictionary<string, object>
        {
            ["Name"] = element.Name,
            ["Category"] = element.Category?.Name ?? string.Empty,
            ["Group"] = groupName,
            ["OwnerView"] = ownerViewName,
            ["Id"] = element.Id.AsLong(),
            ["IsLinked"] = !string.IsNullOrEmpty(linkName),
            ["LinkName"] = linkName ?? string.Empty,
            ["ElementIdObject"] = element.Id, // Store full ElementId for selection
            ["LinkInstanceObject"] = linkInstance, // Store link instance for reference creation
            ["LinkedElementIdObject"] = linkedElementId // Store linked element ID for reference creation
        };

        // Add scope box information
        var containingScopeBoxes = new List<string>();
        try
        {
            // Get element's bounding box
            BoundingBoxXYZ elementBB = element.get_BoundingBox(null);
            if (elementBB == null)
            {
                // Try to get bounding box from geometry
                var options = new Options();
                var geom = element.get_Geometry(options);
                if (geom != null)
                {
                    elementBB = geom.GetBoundingBox();
                }
            }

            if (elementBB != null)
            {
                // For linked elements, transform the bounding box
                if (linkInstance != null)
                {
                    Transform transform = linkInstance.GetTotalTransform();
                    elementBB = TransformBoundingBox(elementBB, transform);
                }

                // Check each scope box in current document
                foreach (var scopeBox in scopeBoxes)
                {
                    BoundingBoxXYZ scopeBB = scopeBox.get_BoundingBox(null);
                    if (scopeBB != null && DoesBoundingBoxIntersect(elementBB, scopeBB))
                    {
                        containingScopeBoxes.Add(scopeBox.Name);
                    }
                }

                // Check each scope box in linked models (if enabled)
                foreach (var linkedSB in linkedScopeBoxes)
                {
                    Element linkedScopeBox = linkedSB.Item1;
                    RevitLinkInstance linkedSBInstance = linkedSB.Item2;

                    BoundingBoxXYZ linkedScopeBB = linkedScopeBox.get_BoundingBox(null);
                    if (linkedScopeBB != null)
                    {
                        // Transform the linked scope box to host coordinates
                        Transform linkedTransform = linkedSBInstance.GetTotalTransform();
                        BoundingBoxXYZ transformedScopeBB = TransformBoundingBox(linkedScopeBB, linkedTransform);

                        if (DoesBoundingBoxIntersect(elementBB, transformedScopeBB))
                        {
                            // Add with link name prefix for clarity
                            string linkedSBName = $"{linkedSBInstance.Name}:{linkedScopeBox.Name}";
                            containingScopeBoxes.Add(linkedSBName);
                        }
                    }
                }
            }
        }
        catch { /* Skip if we can't get bounding box */ }

        data["ScopeBoxes"] = string.Join(", ", containingScopeBoxes.OrderBy(s => s));

        // Add centroid coordinates (converted to display units)
        try
        {
            var (xCentroid, yCentroid, zCentroid) = GetElementCentroid(element, linkInstance);

            // Convert from internal units (feet) to display units
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
            Units projectUnits = elementDoc.GetUnits();
            FormatOptions lengthOpts = projectUnits.GetFormatOptions(SpecTypeId.Length);
            ForgeTypeId unitTypeId = lengthOpts.GetUnitTypeId();

            data["X Centroid"] = xCentroid.HasValue ? UnitUtils.ConvertFromInternalUnits(xCentroid.Value, unitTypeId) : (double?)null;
            data["Y Centroid"] = yCentroid.HasValue ? UnitUtils.ConvertFromInternalUnits(yCentroid.Value, unitTypeId) : (double?)null;
            data["Z Centroid"] = zCentroid.HasValue ? UnitUtils.ConvertFromInternalUnits(zCentroid.Value, unitTypeId) : (double?)null;
#else
            // Revit 2017-2020: Use DisplayUnitType
            Units projectUnits = elementDoc.GetUnits();
            FormatOptions lengthOpts = projectUnits.GetFormatOptions(UnitType.UT_Length);
            DisplayUnitType unitType = lengthOpts.DisplayUnits;

            data["X Centroid"] = xCentroid.HasValue ? UnitUtils.ConvertFromInternalUnits(xCentroid.Value, unitType) : (double?)null;
            data["Y Centroid"] = yCentroid.HasValue ? UnitUtils.ConvertFromInternalUnits(yCentroid.Value, unitType) : (double?)null;
            data["Z Centroid"] = zCentroid.HasValue ? UnitUtils.ConvertFromInternalUnits(zCentroid.Value, unitType) : (double?)null;
#endif
        }
        catch
        {
            // If we can't get centroid, set to null
            data["X Centroid"] = null;
            data["Y Centroid"] = null;
            data["Z Centroid"] = null;
        }

        // Include parameters if requested
        if (includeParameters)
        {
            foreach (Parameter p in element.Parameters)
            {
                try
                {
                    string pName = p.Definition.Name;

                    // Skip parameters we already have from element properties or that show incorrect/redundant information
                    if (pName.Equals("Category", StringComparison.OrdinalIgnoreCase) ||
                        pName.Equals("Name", StringComparison.OrdinalIgnoreCase) ||
                        pName.Equals("Type Name", StringComparison.OrdinalIgnoreCase) ||
                        pName.Equals("Family Name", StringComparison.OrdinalIgnoreCase) ||
                        pName.Equals("Level", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string pValue = p.AsValueString() ?? p.AsString() ?? "None";

                    // Avoid conflicts with existing keys
                    if (!data.ContainsKey(pName))
                    {
                        data[pName] = pValue;
                    }
                    else
                    {
                        // Add suffix to avoid key collision
                        data[$"{pName}_param"] = pValue;
                    }
                }
                catch { /* Skip problematic parameters */ }
            }
        }

        return data;
    }

    /// <summary>
    /// Transform a bounding box by a transform
    /// </summary>
    private static BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bb, Transform transform)
    {
        XYZ min = bb.Min;
        XYZ max = bb.Max;
        
        // Get all 8 corners of the bounding box
        var corners = new[]
        {
            new XYZ(min.X, min.Y, min.Z),
            new XYZ(max.X, min.Y, min.Z),
            new XYZ(min.X, max.Y, min.Z),
            new XYZ(max.X, max.Y, min.Z),
            new XYZ(min.X, min.Y, max.Z),
            new XYZ(max.X, min.Y, max.Z),
            new XYZ(min.X, max.Y, max.Z),
            new XYZ(max.X, max.Y, max.Z)
        };

        // Transform all corners
        var transformedCorners = corners.Select(c => transform.OfPoint(c)).ToArray();

        // Find new min and max
        double minX = transformedCorners.Min(p => p.X);
        double minY = transformedCorners.Min(p => p.Y);
        double minZ = transformedCorners.Min(p => p.Z);
        double maxX = transformedCorners.Max(p => p.X);
        double maxY = transformedCorners.Max(p => p.Y);
        double maxZ = transformedCorners.Max(p => p.Z);

        var result = new BoundingBoxXYZ();
        result.Min = new XYZ(minX, minY, minZ);
        result.Max = new XYZ(maxX, maxY, maxZ);
        
        return result;
    }

    /// <summary>
    /// Check if two bounding boxes intersect
    /// </summary>
    private static bool DoesBoundingBoxIntersect(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        // Check if bb1 is completely outside bb2 in any dimension
        if (bb1.Max.X < bb2.Min.X || bb1.Min.X > bb2.Max.X) return false;
        if (bb1.Max.Y < bb2.Min.Y || bb1.Min.Y > bb2.Max.Y) return false;
        if (bb1.Max.Z < bb2.Min.Z || bb1.Min.Z > bb2.Max.Z) return false;

        return true;
    }

    /// <summary>
    /// Get centroid coordinates for an element
    /// Returns (x, y, z) or (null, null, null) if not applicable
    /// </summary>
    private static (double?, double?, double?) GetElementCentroid(Element element, RevitLinkInstance linkInstance)
    {
        // Special case: Viewports on sheets
        if (element is Viewport viewport)
        {
            XYZ center = viewport.GetBoxCenter();
            if (linkInstance != null)
            {
                Transform transform = linkInstance.GetTotalTransform();
                center = transform.OfPoint(center);
            }
            return (center.X, center.Y, null); // Viewports are 2D on sheets
        }

        // Get location point (for point-based elements like families)
        LocationPoint locationPoint = element.Location as LocationPoint;
        if (locationPoint != null)
        {
            XYZ point = locationPoint.Point;
            if (linkInstance != null)
            {
                Transform transform = linkInstance.GetTotalTransform();
                point = transform.OfPoint(point);
            }

            // Check if element is constrained vertically (like walls constrained by levels)
            // Walls and other level-constrained elements should not allow Z editing
            bool isLevelConstrained = IsLevelConstrained(element);

            return (point.X, point.Y, isLevelConstrained ? (double?)null : point.Z);
        }

        // Get location curve (for curve-based elements like walls, beams)
        LocationCurve locationCurve = element.Location as LocationCurve;
        if (locationCurve != null)
        {
            Curve curve = locationCurve.Curve;
            XYZ midpoint = (curve.GetEndPoint(0) + curve.GetEndPoint(1)) / 2.0;
            if (linkInstance != null)
            {
                Transform transform = linkInstance.GetTotalTransform();
                midpoint = transform.OfPoint(midpoint);
            }

            // Curve-based elements like walls are typically level-constrained
            bool isLevelConstrained = IsLevelConstrained(element);

            return (midpoint.X, midpoint.Y, isLevelConstrained ? (double?)null : midpoint.Z);
        }

        // For elements without location, use bounding box centroid
        BoundingBoxXYZ bb = element.get_BoundingBox(null);
        if (bb == null)
        {
            var options = new Options();
            var geom = element.get_Geometry(options);
            if (geom != null)
            {
                bb = geom.GetBoundingBox();
            }
        }

        if (bb != null)
        {
            XYZ centroid = (bb.Min + bb.Max) / 2.0;
            if (linkInstance != null)
            {
                Transform transform = linkInstance.GetTotalTransform();
                centroid = transform.OfPoint(centroid);
            }

            // Bounding box elements - allow all 3D movement
            return (centroid.X, centroid.Y, centroid.Z);
        }

        // No location or bounding box available
        return (null, null, null);
    }

    /// <summary>
    /// Check if element is constrained by levels (like walls)
    /// </summary>
    private static bool IsLevelConstrained(Element element)
    {
        // Walls are always level-constrained
        if (element is Wall)
            return true;

        // Check for base/top constraint parameters
        Parameter baseConstraint = element.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
        Parameter topConstraint = element.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);

        if (baseConstraint != null && baseConstraint.HasValue)
            return true;
        if (topConstraint != null && topConstraint.HasValue)
            return true;

        // Check for level parameter
        Parameter levelParam = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
        if (levelParam != null && levelParam.HasValue)
        {
            // Face-based or hosted families may be level-constrained
            if (element is FamilyInstance famInst)
            {
                if (famInst.Host != null)
                    return true; // Hosted elements on faces/walls
            }
        }

        return false;
    }
}

/// <summary>
/// Base class for commands that display Revit elements in a custom data‑grid for filtering and re‑selection.
/// Now supports elements from linked models and includes scope box information.
/// </summary>
public abstract class FilterElementsBase : IExternalCommand
{
    public abstract bool SpanAllScreens { get; }
    public abstract bool UseSelectedElements { get; }
    public abstract bool IncludeParameters { get; }

    public Result Execute(ExternalCommandData cData, ref string message, ElementSet elements)
    {
        try
        {
            var uiDoc = cData.Application.ActiveUIDocument;
            var elementData = ElementDataHelper.GetElementData(uiDoc, UseSelectedElements, IncludeParameters);

            if (!elementData.Any())
            {
                TaskDialog.Show("Info", "No elements found.");
                return Result.Cancelled;
            }

            // Get property names, excluding internal object fields
            var propertyNames = elementData.First().Keys
                .Where(k => !k.EndsWith("Object"))
                .ToList();

            // Check which optional columns have any non-empty values
            bool hasScopeBoxes = elementData.Any(d => d.ContainsKey("ScopeBoxes") && !string.IsNullOrEmpty(d["ScopeBoxes"]?.ToString()));
            bool hasLinkName = elementData.Any(d => d.ContainsKey("LinkName") && !string.IsNullOrEmpty(d["LinkName"]?.ToString()));
            bool hasGroup = elementData.Any(d => d.ContainsKey("Group") && !string.IsNullOrEmpty(d["Group"]?.ToString()));
            bool hasOwnerView = elementData.Any(d => d.ContainsKey("OwnerView") && !string.IsNullOrEmpty(d["OwnerView"]?.ToString()));

            // Build ordered list, only including columns that have values
            var orderedProps = new List<string> { "Name" };
            if (hasScopeBoxes) orderedProps.Add("ScopeBoxes");
            orderedProps.Add("Category");
            if (hasLinkName) orderedProps.Add("LinkName");
            if (hasGroup) orderedProps.Add("Group");
            if (hasOwnerView) orderedProps.Add("OwnerView");

            // Add centroid columns
            orderedProps.Add("X Centroid");
            orderedProps.Add("Y Centroid");
            orderedProps.Add("Z Centroid");

            orderedProps.Add("Id");

            var remainingProps = propertyNames.Except(orderedProps).OrderBy(p => p);
            propertyNames = orderedProps.Where(p => propertyNames.Contains(p))
                .Concat(remainingProps)
                .ToList();

            // Set the current UIDocument for edit operations
            CustomGUIs.SetCurrentUIDocument(uiDoc);

            var chosenRows = CustomGUIs.DataGrid(elementData, propertyNames, SpanAllScreens);

            // Apply any pending edits to Revit elements
            if (CustomGUIs.HasPendingEdits())
            {
                CustomGUIs.ApplyCellEditsToEntities();
            }

            if (chosenRows.Count == 0)
                return Result.Cancelled;

            // Separate regular elements and linked elements
            var regularIds = new List<ElementId>();
            var linkedReferences = new List<Reference>();

            foreach (var row in chosenRows)
            {
                // Check if this is a linked element
                if (row.TryGetValue("LinkInstanceObject", out var linkObj) && linkObj is RevitLinkInstance linkInstance &&
                    row.TryGetValue("LinkedElementIdObject", out var linkedIdObj) && linkedIdObj is ElementId linkedElementId)
                {
                    // This is a linked element - create reference the same way SelectCategories does
                    try
                    {
                        Document linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            Element linkedElement = linkedDoc.GetElement(linkedElementId);
                            if (linkedElement != null)
                            {
                                // Create reference for the element
                                Reference elemRef = new Reference(linkedElement);
                                Reference linkedRef = elemRef.CreateLinkReference(linkInstance);

                                if (linkedRef != null)
                                {
                                    linkedReferences.Add(linkedRef);
                                }
                            }
                        }
                    }
                    catch { }
                }
                else if (row.TryGetValue("ElementIdObject", out var idObj) && idObj is ElementId elemId)
                {
                    // This is a regular element (not linked)
                    regularIds.Add(elemId);
                }
                // Handle backward compatibility - if someone stored just the integer ID
                else if (row.TryGetValue("Id", out var intId) && intId is int id)
                {
                    regularIds.Add(id.ToElementId());
                }
            }

            // Set selection based on what we have
            if (linkedReferences.Any() && !regularIds.Any())
            {
                // Only linked elements - use SetReferences
                uiDoc.SetReferences(linkedReferences);
            }
            else if (!linkedReferences.Any() && regularIds.Any())
            {
                // Only regular elements - use SetElementIds
                uiDoc.SetSelectionIds(regularIds);
            }
            else if (linkedReferences.Any() && regularIds.Any())
            {
                // Mixed selection - Revit API doesn't support mixing SetElementIds with SetReferences
                // We need to choose one approach. Let's prioritize what we have more of.
                if (linkedReferences.Count >= regularIds.Count)
                {
                    // More linked elements - convert regular elements to references if possible
                    var allReferences = new List<Reference>(linkedReferences);
                    
                    // Try to add regular elements as references
                    foreach (var elemId in regularIds)
                    {
                        try
                        {
                            Element elem = cData.Application.ActiveUIDocument.Document.GetElement(elemId);
                            if (elem != null)
                            {
                                Reference elemRef = new Reference(elem);
                                if (elemRef != null)
                                {
                                    allReferences.Add(elemRef);
                                }
                            }
                        }
                        catch { }
                    }
                    
                    uiDoc.SetReferences(allReferences);
                }
                else
                {
                    // More regular elements - just select those
                    uiDoc.SetSelectionIds(regularIds);
                    
                    TaskDialog.Show("Mixed Selection",
                        $"Selected {regularIds.Count} regular elements.\n" +
                        $"Note: {linkedReferences.Count} linked elements could not be included in the selection " +
                        "because Revit doesn't support mixing regular and linked selections.");
                }
            }

            return Result.Succeeded;
        }
        catch (InvalidOperationException ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
        catch (Exception ex)
        {
            message = $"Unexpected error: {ex.Message}";
            return Result.Failed;
        }
    }
}

#region Concrete commands

[Transaction(TransactionMode.Manual)]
public class FilterSelectedInProject : FilterElementsBase
{
    public override bool SpanAllScreens      => false;
    public override bool UseSelectedElements => true;
    public override bool IncludeParameters   => true;
}

[Transaction(TransactionMode.Manual)]
public class FilterSelectedInViews : IExternalCommand
{
    public Result Execute(ExternalCommandData cData, ref string message, ElementSet elements)
    {
        try
        {
            var uiDoc = cData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;

            // Get selected elements
            var elementData = ElementDataHelper.GetElementData(uiDoc, selectedOnly: true, includeParameters: true);

            if (!elementData.Any())
            {
                TaskDialog.Show("Info", "No elements selected.");
                return Result.Cancelled;
            }

            // Get current view and any selected views
            var currentView = doc.ActiveView;
            var viewsToCheck = new HashSet<ElementId> { currentView.Id };

            // Check if any selected elements are views
            var selectedIds = uiDoc.GetSelectionIds();
            foreach (var id in selectedIds)
            {
                var elem = doc.GetElement(id);
                if (elem is View view && !view.IsTemplate)
                {
                    viewsToCheck.Add(view.Id);
                }
            }

            // Filter elements to only those visible in the views we're checking
            var filteredData = new List<Dictionary<string, object>>();

            foreach (var data in elementData)
            {
                bool isVisibleInAnyView = false;

                // Get the element ID
                ElementId elementId = null;
                if (data.TryGetValue("ElementIdObject", out var elemIdObj) && elemIdObj is ElementId eid)
                {
                    elementId = eid;
                }
                else if (data.TryGetValue("Id", out var intId) && intId is int id)
                {
                    elementId = id.ToElementId();
                }

                if (elementId == null)
                    continue;

                // Check if linked element
                bool isLinked = data.TryGetValue("IsLinked", out var linkedObj) && linkedObj is bool linked && linked;

                if (isLinked)
                {
                    // For linked elements, we can't easily check view visibility, so include them all
                    isVisibleInAnyView = true;
                }
                else
                {
                    // Check each view
                    foreach (var viewId in viewsToCheck)
                    {
                        try
                        {
                            var view = doc.GetElement(viewId) as View;
                            if (view == null)
                                continue;

                            // Use FilteredElementCollector to check if element is in view
                            var elementsInView = new FilteredElementCollector(doc, viewId)
                                .ToElementIds();

                            if (elementsInView.Contains(elementId))
                            {
                                isVisibleInAnyView = true;
                                break;
                            }
                        }
                        catch { /* Skip problematic views */ }
                    }
                }

                if (isVisibleInAnyView)
                {
                    filteredData.Add(data);
                }
            }

            if (!filteredData.Any())
            {
                string viewNames = string.Join(", ", viewsToCheck.Select(id => doc.GetElement(id)?.Name ?? "Unknown"));
                TaskDialog.Show("Info", $"None of the selected elements are visible in the checked view(s): {viewNames}");
                return Result.Cancelled;
            }

            // Get property names, excluding internal object fields
            var propertyNames = filteredData.First().Keys
                .Where(k => !k.EndsWith("Object"))
                .ToList();

            // Check which optional columns have any non-empty values
            bool hasScopeBoxes = filteredData.Any(d => d.ContainsKey("ScopeBoxes") && !string.IsNullOrEmpty(d["ScopeBoxes"]?.ToString()));
            bool hasLinkName = filteredData.Any(d => d.ContainsKey("LinkName") && !string.IsNullOrEmpty(d["LinkName"]?.ToString()));
            bool hasGroup = filteredData.Any(d => d.ContainsKey("Group") && !string.IsNullOrEmpty(d["Group"]?.ToString()));
            bool hasOwnerView = filteredData.Any(d => d.ContainsKey("OwnerView") && !string.IsNullOrEmpty(d["OwnerView"]?.ToString()));

            // Build ordered list, only including columns that have values
            var orderedProps = new List<string> { "Name" };
            if (hasScopeBoxes) orderedProps.Add("ScopeBoxes");
            orderedProps.Add("Category");
            if (hasLinkName) orderedProps.Add("LinkName");
            if (hasGroup) orderedProps.Add("Group");
            if (hasOwnerView) orderedProps.Add("OwnerView");

            // Add centroid columns
            orderedProps.Add("X Centroid");
            orderedProps.Add("Y Centroid");
            orderedProps.Add("Z Centroid");

            orderedProps.Add("Id");

            var remainingProps = propertyNames.Except(orderedProps).OrderBy(p => p);
            propertyNames = orderedProps.Where(p => propertyNames.Contains(p))
                .Concat(remainingProps)
                .ToList();

            // Set the current UIDocument for edit operations
            CustomGUIs.SetCurrentUIDocument(uiDoc);

            var chosenRows = CustomGUIs.DataGrid(filteredData, propertyNames, false);

            // Apply any pending edits to Revit elements
            if (CustomGUIs.HasPendingEdits())
            {
                CustomGUIs.ApplyCellEditsToEntities();
            }

            if (chosenRows.Count == 0)
                return Result.Cancelled;

            // Separate regular elements and linked elements
            var regularIds = new List<ElementId>();
            var linkedReferences = new List<Reference>();

            foreach (var row in chosenRows)
            {
                // Check if this is a linked element
                if (row.TryGetValue("LinkInstanceObject", out var linkObj) && linkObj is RevitLinkInstance linkInstance &&
                    row.TryGetValue("LinkedElementIdObject", out var linkedIdObj) && linkedIdObj is ElementId linkedElementId)
                {
                    // This is a linked element - create reference the same way SelectCategories does
                    try
                    {
                        Document linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            Element linkedElement = linkedDoc.GetElement(linkedElementId);
                            if (linkedElement != null)
                            {
                                // Create reference for the element
                                Reference elemRef = new Reference(linkedElement);
                                Reference linkedRef = elemRef.CreateLinkReference(linkInstance);

                                if (linkedRef != null)
                                {
                                    linkedReferences.Add(linkedRef);
                                }
                            }
                        }
                    }
                    catch { }
                }
                else if (row.TryGetValue("ElementIdObject", out var idObj) && idObj is ElementId elemId)
                {
                    // This is a regular element (not linked)
                    regularIds.Add(elemId);
                }
                // Handle backward compatibility - if someone stored just the integer ID
                else if (row.TryGetValue("Id", out var intId) && intId is int id)
                {
                    regularIds.Add(id.ToElementId());
                }
            }

            // Set selection based on what we have
            if (linkedReferences.Any() && !regularIds.Any())
            {
                // Only linked elements - use SetReferences
                uiDoc.SetReferences(linkedReferences);
            }
            else if (!linkedReferences.Any() && regularIds.Any())
            {
                // Only regular elements - use SetElementIds
                uiDoc.SetSelectionIds(regularIds);
            }
            else if (linkedReferences.Any() && regularIds.Any())
            {
                // Mixed selection - Revit API doesn't support mixing SetElementIds with SetReferences
                // We need to choose one approach. Let's prioritize what we have more of.
                if (linkedReferences.Count >= regularIds.Count)
                {
                    // More linked elements - convert regular elements to references if possible
                    var allReferences = new List<Reference>(linkedReferences);

                    // Try to add regular elements as references
                    foreach (var elemId in regularIds)
                    {
                        try
                        {
                            Element elem = cData.Application.ActiveUIDocument.Document.GetElement(elemId);
                            if (elem != null)
                            {
                                Reference elemRef = new Reference(elem);
                                if (elemRef != null)
                                {
                                    allReferences.Add(elemRef);
                                }
                            }
                        }
                        catch { }
                    }

                    uiDoc.SetReferences(allReferences);
                }
                else
                {
                    // More regular elements - just select those
                    uiDoc.SetSelectionIds(regularIds);

                    TaskDialog.Show("Mixed Selection",
                        $"Selected {regularIds.Count} regular elements.\n" +
                        $"Note: {linkedReferences.Count} linked elements could not be included in the selection " +
                        "because Revit doesn't support mixing regular and linked selections.");
                }
            }

            return Result.Succeeded;
        }
        catch (InvalidOperationException ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
        catch (Exception ex)
        {
            message = $"Unexpected error: {ex.Message}";
            return Result.Failed;
        }
    }
}

#endregion
