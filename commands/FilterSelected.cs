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
        return GetElementData(uiDoc, selectedOnly, includeParameters, null, null);
    }

    public static List<Dictionary<string, object>> GetElementData(UIDocument uiDoc, bool selectedOnly, bool includeParameters, Func<bool> checkCancellation, CancellableProgressDialog progressDialog = null)
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

            // Set total for progress tracking
            progressDialog?.SetTotal(selectedIds.Count + selectedRefs.Count);

            // Keep track of processed elements to avoid duplicates
            var processedElements = new HashSet<ElementId>();

            // Process regular elements from current document
            foreach (var id in selectedIds)
            {
                // Check and show progress dialog if needed
                progressDialog?.CheckAndShow();

                // Check for cancellation
                if (checkCancellation != null && checkCancellation())
                    throw new OperationCanceledException("Operation cancelled by user.");

                Element element = doc.GetElement(id);
                if (element != null)
                {
                    processedElements.Add(id);
                    var data = GetElementDataDictionary(element, doc, null, null, null, includeParameters, scopeBoxes, linkedScopeBoxes);
                    elementData.Add(data);
                }

                // Increment progress
                progressDialog?.IncrementProgress();
            }

            // Process linked elements via References
            foreach (var reference in selectedRefs)
            {
                // Check and show progress dialog if needed
                progressDialog?.CheckAndShow();

                // Check for cancellation
                if (checkCancellation != null && checkCancellation())
                    throw new OperationCanceledException("Operation cancelled by user.");

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

                // Increment progress
                progressDialog?.IncrementProgress();
            }
        }
        else
        {
            // Get elements from active view (current document only)
            var elementIds = new FilteredElementCollector(doc, doc.ActiveView.Id).ToElementIds();
            var elementIdsList = elementIds.ToList();

            // Set total for progress tracking
            progressDialog?.SetTotal(elementIdsList.Count);

            foreach (var id in elementIdsList)
            {
                // Check and show progress dialog if needed
                progressDialog?.CheckAndShow();

                // Check for cancellation
                if (checkCancellation != null && checkCancellation())
                    throw new OperationCanceledException("Operation cancelled by user.");

                Element element = doc.GetElement(id);
                if (element != null)
                {
                    var data = GetElementDataDictionary(element, doc, null, null, null, includeParameters, scopeBoxes, linkedScopeBoxes);
                    elementData.Add(data);
                }

                // Increment progress
                progressDialog?.IncrementProgress();
            }
        }

        return elementData;
    }

    /// <summary>
    /// Creates a data dictionary for an element with standard columns including Family and Type Name.
    ///
    /// FRAMEWORK NOTE: This method provides automatic family/type editing support for all commands.
    /// Any command using ElementDataHelper.GetElementData() will automatically get:
    /// - "Family" column - Editable family name (FamilySymbol.FamilyName or system type family name)
    /// - "Type Name" column - Editable type name (ElementType.Name)
    /// - "ElementIdObject" - Stored ElementId for reliable element lookup after edits
    /// - "_IsImportSymbol" flag - Marks DWG import symbols for optional filtering
    ///
    /// To enable editing in your command:
    /// 1. Call CustomGUIs.SetCurrentUIDocument(uidoc) before showing DataGrid
    /// 2. Check CustomGUIs.HasPendingEdits() after DataGrid closes
    /// 3. Call CustomGUIs.ApplyCellEditsToEntities() to apply edits
    ///
    /// The edit system automatically handles:
    /// - Renaming families (via Family.Name property)
    /// - Renaming types (via ElementType.Name property)
    /// - Validation and transactions
    /// </summary>
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

        // Add Type Name and Family columns for elements with types
        ElementId typeId = element.GetTypeId();
        if (typeId != null && typeId != ElementId.InvalidElementId)
        {
            Element typeElement = elementDoc.GetElement(typeId);
            if (typeElement != null)
            {
                data["Type Name"] = typeElement.Name;

                // For family symbols, get the family name
                if (typeElement is FamilySymbol familySymbol)
                {
                    data["Family"] = familySymbol.FamilyName;

                    // Filter out DWG import symbols
                    if (data["Family"].ToString().Contains("Import Symbol") ||
                        (element.Category != null && element.Category.Name.Contains("Import Symbol")))
                    {
                        // Mark as import symbol so it can be filtered if needed
                        data["_IsImportSymbol"] = true;
                    }
                }
                else
                {
                    // For system types, try to get family name from parameter
                    Parameter familyParam = typeElement.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                    data["Family"] = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
                        ? familyParam.AsString()
                        : "System Type";
                }
            }
        }

        // Add scope box information
        var containingScopeBoxes = new List<string>();
        try
        {
            // Get element's bounding box
            // IMPORTANT: Do NOT call get_Geometry() as it triggers regeneration for viewports (~1000ms each)
            // Only use get_BoundingBox(null) which reads from the database
            BoundingBoxXYZ elementBB = element.get_BoundingBox(null);

            // For viewports, use GetBoxOutline() if bounding box is not available
            if (elementBB == null && element is Viewport vp)
            {
                try
                {
                    Outline outline = vp.GetBoxOutline();
                    elementBB = new BoundingBoxXYZ();
                    elementBB.Min = outline.MinimumPoint;
                    elementBB.Max = outline.MaximumPoint;
                }
                catch { /* Skip if we can't get outline */ }
            }
            // Note: We intentionally do NOT fall back to get_Geometry() here to avoid
            // triggering regeneration of inactive views/sheets

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

        // Add view-specific scope box property (assigned scope box, not containing)
        if (element is View view)
        {
            try
            {
                Parameter scopeBoxParam = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP);
                if (scopeBoxParam != null && scopeBoxParam.HasValue)
                {
                    ElementId scopeBoxId = scopeBoxParam.AsElementId();
                    if (scopeBoxId != null && scopeBoxId != ElementId.InvalidElementId)
                    {
                        Element assignedScopeBox = elementDoc.GetElement(scopeBoxId);
                        data["Scope Box"] = assignedScopeBox?.Name ?? "";
                    }
                    else
                    {
                        data["Scope Box"] = "";
                    }
                }
                else
                {
                    data["Scope Box"] = "";
                }
            }
            catch
            {
                data["Scope Box"] = "";
            }
        }

        // Add view boolean properties for views and viewports
        try
        {
            Autodesk.Revit.DB.View viewForProperties = null;
            if (element is Autodesk.Revit.DB.View viewElem)
            {
                viewForProperties = viewElem;
            }
            else if (element is Viewport viewport)
            {
                viewForProperties = elementDoc.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;
            }

            if (viewForProperties != null)
            {
                data["Crop View"] = viewForProperties.CropBoxActive;
                data["Crop Region Visible"] = viewForProperties.CropBoxVisible;

                // AnnotationCropActive may not be available in all Revit versions
                try
                {
                    data["Annotation Crop"] = viewForProperties.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)?.AsInteger() == 1;
                }
                catch
                {
                    // Property not available in this Revit version
                }
            }
        }
        catch
        {
            // If we can't get view properties, skip these columns
        }

        // Add crop region columns for views and viewports (only if rectangular)
        try
        {
            Autodesk.Revit.DB.View viewForCrop = null;
            if (element is Autodesk.Revit.DB.View viewElem)
            {
                viewForCrop = viewElem;
            }
            else if (element is Viewport viewport)
            {
                viewForCrop = elementDoc.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;
            }

            if (viewForCrop != null && viewForCrop.CropBoxActive && viewForCrop.CropBox != null)
            {
                // Check if crop region is rectangular
                bool isRectangular = true;
                if (viewForCrop is ViewPlan plan)
                {
                    try
                    {
                        var managerProperty = plan.GetType().GetProperty("CropRegionShapeManager");
                        if (managerProperty != null)
                        {
                            object manager = managerProperty.GetValue(plan, null);
                            if (manager != null)
                            {
                                var method = manager.GetType().GetMethod("GetCropRegionShape");
                                if (method != null)
                                {
                                    CurveLoop cropLoop = method.Invoke(manager, null) as CurveLoop;
                                    if (cropLoop != null)
                                    {
                                        isRectangular = cropLoop.ToList().Count == 4;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        isRectangular = false;
                    }
                }

                if (isRectangular)
                {
                    BoundingBoxXYZ bbox = viewForCrop.CropBox;
                    Transform transform = bbox.Transform;

                    // Get survey point elevation
                    double surveyElevation = 0.0;
                    try
                    {
                        FilteredElementCollector surveyCollector = new FilteredElementCollector(elementDoc);
                        var surveyPoint = surveyCollector.OfCategory(BuiltInCategory.OST_SharedBasePoint)
                            .FirstOrDefault() as BasePoint;
                        if (surveyPoint != null)
                        {
                            surveyElevation = surveyPoint.get_BoundingBox(null).Min.Z;
                        }
                    }
                    catch { }

                    // Transform crop box corners to project coordinates
                    XYZ minCorner = transform.OfPoint(bbox.Min);
                    XYZ maxCorner = transform.OfPoint(bbox.Max);

                    // Determine coordinate interpretation based on view type
                    bool isElevationOrSection = viewForCrop.ViewType == ViewType.Elevation || viewForCrop.ViewType == ViewType.Section;

                    double top, bottom, left, right;
                    if (isElevationOrSection)
                    {
                        // For elevation/section: Top/Bottom are elevations (Z) relative to survey point
                        top = maxCorner.Z - surveyElevation;
                        bottom = minCorner.Z - surveyElevation;
                        left = minCorner.X;
                        right = maxCorner.X;
                    }
                    else
                    {
                        // For plan: Top/Bottom are Y, Left/Right are X
                        top = maxCorner.Y;
                        bottom = minCorner.Y;
                        left = minCorner.X;
                        right = maxCorner.X;
                    }

                    // Convert from internal units (feet) to display units
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                    Units projectUnits = elementDoc.GetUnits();
                    FormatOptions lengthOpts = projectUnits.GetFormatOptions(SpecTypeId.Length);
                    ForgeTypeId unitTypeId = lengthOpts.GetUnitTypeId();

                    data["Crop Region Top"] = UnitUtils.ConvertFromInternalUnits(top, unitTypeId);
                    data["Crop Region Bottom"] = UnitUtils.ConvertFromInternalUnits(bottom, unitTypeId);
                    data["Crop Region Left"] = UnitUtils.ConvertFromInternalUnits(left, unitTypeId);
                    data["Crop Region Right"] = UnitUtils.ConvertFromInternalUnits(right, unitTypeId);
#else
                    // Revit 2017-2020: Use DisplayUnitType
                    Units projectUnits = elementDoc.GetUnits();
                    FormatOptions lengthOpts = projectUnits.GetFormatOptions(UnitType.UT_Length);
                    DisplayUnitType unitType = lengthOpts.DisplayUnits;

                    data["Crop Region Top"] = UnitUtils.ConvertFromInternalUnits(top, unitType);
                    data["Crop Region Bottom"] = UnitUtils.ConvertFromInternalUnits(bottom, unitType);
                    data["Crop Region Left"] = UnitUtils.ConvertFromInternalUnits(left, unitType);
                    data["Crop Region Right"] = UnitUtils.ConvertFromInternalUnits(right, unitType);
#endif
                }
            }
        }
        catch
        {
            // If we can't get crop region, skip these columns
        }

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
        // PERFORMANCE: Use GetBoxOutline() instead of GetBoxCenter() to avoid triggering regeneration
        // GetBoxOutline() gets viewport box bounds (not label) and should be faster than GetBoxCenter()
        // GetBoxCenter() triggers regeneration (~350ms), GetBoxOutline() should be database-only
        // We manually calculate center from outline: (Min + Max) / 2
        if (element is Viewport viewport)
        {
            try
            {
                Outline outline = viewport.GetBoxOutline();
                XYZ min = outline.MinimumPoint;
                XYZ max = outline.MaximumPoint;
                XYZ center = (min + max) / 2.0;

                if (linkInstance != null)
                {
                    Transform transform = linkInstance.GetTotalTransform();
                    center = transform.OfPoint(center);
                }
                return (center.X, center.Y, null); // Viewports are 2D on sheets
            }
            catch
            {
                // If we can't get box outline, return null
                return (null, null, null);
            }
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
        // IMPORTANT: Do NOT call get_Geometry() as it can trigger regeneration
        // Only use get_BoundingBox(null) which reads from the database
        BoundingBoxXYZ bb = element.get_BoundingBox(null);

        // Special case: Text elements on sheets have Location type but get_BoundingBox(null) returns null
        // Need to use get_BoundingBox(view) with the owner view
        if (bb == null && element.OwnerViewId != null && element.OwnerViewId != ElementId.InvalidElementId)
        {
            try
            {
                Document elementDoc = element.Document;
                View ownerView = elementDoc.GetElement(element.OwnerViewId) as View;
                if (ownerView != null)
                {
                    bb = element.get_BoundingBox(ownerView);
                }
            }
            catch { /* Skip if we can't get owner view bounding box */ }
        }
        // Note: We intentionally do NOT fall back to get_Geometry() here to avoid
        // triggering regeneration of inactive views/sheets

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
            List<Dictionary<string, object>> elementData;

            // Use cancellable progress dialog for potentially long operations
            using (var progress = new CancellableProgressDialog("Collecting element data"))
            {
                progress.Start();
                try
                {
                    elementData = ElementDataHelper.GetElementData(uiDoc, UseSelectedElements, IncludeParameters, () => progress.IsCancelled, progress);
                }
                catch (OperationCanceledException)
                {
                    message = "Operation cancelled by user.";
                    return Result.Cancelled;
                }
            }

            if (!elementData.Any())
            {
                TaskDialog.Show("Info", "No elements found.");
                return Result.Cancelled;
            }

            // Get ALL property names from ALL elements (union, not intersection)
            var allPropertyNames = new HashSet<string>();
            foreach (var data in elementData)
            {
                foreach (var key in data.Keys)
                {
                    if (!key.EndsWith("Object"))  // Exclude internal object fields
                    {
                        allPropertyNames.Add(key);
                    }
                }
            }

            // Build ordered list with standard columns in preferred order
            var orderedProps = new List<string> { "Name" };
            if (allPropertyNames.Contains("Type Name")) orderedProps.Add("Type Name");  // Editable
            if (allPropertyNames.Contains("Family")) orderedProps.Add("Family");  // Editable
            if (allPropertyNames.Contains("Scope Box")) orderedProps.Add("Scope Box");  // View scope box (editable)
            if (allPropertyNames.Contains("ScopeBoxes")) orderedProps.Add("ScopeBoxes");  // Element containment (read-only)
            orderedProps.Add("Category");
            if (allPropertyNames.Contains("LinkName")) orderedProps.Add("LinkName");
            if (allPropertyNames.Contains("Group")) orderedProps.Add("Group");
            if (allPropertyNames.Contains("OwnerView")) orderedProps.Add("OwnerView");

            // Add view boolean properties
            if (allPropertyNames.Contains("Crop View")) orderedProps.Add("Crop View");
            if (allPropertyNames.Contains("Crop Region Visible")) orderedProps.Add("Crop Region Visible");
            if (allPropertyNames.Contains("Annotation Crop")) orderedProps.Add("Annotation Crop");

            // Add crop region columns (editable)
            if (allPropertyNames.Contains("Crop Region Top")) orderedProps.Add("Crop Region Top");
            if (allPropertyNames.Contains("Crop Region Bottom")) orderedProps.Add("Crop Region Bottom");
            if (allPropertyNames.Contains("Crop Region Left")) orderedProps.Add("Crop Region Left");
            if (allPropertyNames.Contains("Crop Region Right")) orderedProps.Add("Crop Region Right");

            // Add centroid columns
            if (allPropertyNames.Contains("X Centroid")) orderedProps.Add("X Centroid");
            if (allPropertyNames.Contains("Y Centroid")) orderedProps.Add("Y Centroid");
            if (allPropertyNames.Contains("Z Centroid")) orderedProps.Add("Z Centroid");

            orderedProps.Add("Id");

            var remainingProps = allPropertyNames.Except(orderedProps).OrderBy(p => p);
            var propertyNames = orderedProps.Where(p => allPropertyNames.Contains(p))
                .Concat(remainingProps)
                .ToList();

            // Set the current UIDocument for edit operations
            CustomGUIs.SetCurrentUIDocument(uiDoc);

            var chosenRows = CustomGUIs.DataGrid(elementData, propertyNames, SpanAllScreens);

            // Apply any pending edits to Revit elements
            // CRITICAL: Only apply edits if user pressed Enter (not Escape)
            if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
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
public class FilterSelectedInDocument : FilterElementsBase
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

            // Get selected elements with cancellable progress dialog
            List<Dictionary<string, object>> elementData;
            using (var progress = new CancellableProgressDialog("Collecting element data"))
            {
                progress.Start();
                try
                {
                    elementData = ElementDataHelper.GetElementData(uiDoc, selectedOnly: true, includeParameters: true, () => progress.IsCancelled, progress);
                }
                catch (OperationCanceledException)
                {
                    message = "Operation cancelled by user.";
                    return Result.Cancelled;
                }
            }

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

            // Get ALL property names from ALL elements (union, not intersection)
            var allPropertyNames = new HashSet<string>();
            foreach (var data in filteredData)
            {
                foreach (var key in data.Keys)
                {
                    if (!key.EndsWith("Object"))  // Exclude internal object fields
                    {
                        allPropertyNames.Add(key);
                    }
                }
            }

            // Build ordered list with standard columns in preferred order
            var orderedProps = new List<string> { "Name" };
            if (allPropertyNames.Contains("Type Name")) orderedProps.Add("Type Name");  // Editable
            if (allPropertyNames.Contains("Family")) orderedProps.Add("Family");  // Editable
            if (allPropertyNames.Contains("Scope Box")) orderedProps.Add("Scope Box");  // View scope box (editable)
            if (allPropertyNames.Contains("ScopeBoxes")) orderedProps.Add("ScopeBoxes");  // Element containment (read-only)
            orderedProps.Add("Category");
            if (allPropertyNames.Contains("LinkName")) orderedProps.Add("LinkName");
            if (allPropertyNames.Contains("Group")) orderedProps.Add("Group");
            if (allPropertyNames.Contains("OwnerView")) orderedProps.Add("OwnerView");

            // Add view boolean properties
            if (allPropertyNames.Contains("Crop View")) orderedProps.Add("Crop View");
            if (allPropertyNames.Contains("Crop Region Visible")) orderedProps.Add("Crop Region Visible");
            if (allPropertyNames.Contains("Annotation Crop")) orderedProps.Add("Annotation Crop");

            // Add crop region columns (editable)
            if (allPropertyNames.Contains("Crop Region Top")) orderedProps.Add("Crop Region Top");
            if (allPropertyNames.Contains("Crop Region Bottom")) orderedProps.Add("Crop Region Bottom");
            if (allPropertyNames.Contains("Crop Region Left")) orderedProps.Add("Crop Region Left");
            if (allPropertyNames.Contains("Crop Region Right")) orderedProps.Add("Crop Region Right");

            // Add centroid columns
            if (allPropertyNames.Contains("X Centroid")) orderedProps.Add("X Centroid");
            if (allPropertyNames.Contains("Y Centroid")) orderedProps.Add("Y Centroid");
            if (allPropertyNames.Contains("Z Centroid")) orderedProps.Add("Z Centroid");

            orderedProps.Add("Id");

            var remainingProps = allPropertyNames.Except(orderedProps).OrderBy(p => p);
            var propertyNames = orderedProps.Where(p => allPropertyNames.Contains(p))
                .Concat(remainingProps)
                .ToList();

            // Set the current UIDocument for edit operations
            CustomGUIs.SetCurrentUIDocument(uiDoc);

            var chosenRows = CustomGUIs.DataGrid(filteredData, propertyNames, false);

            // Apply any pending edits to Revit elements
            // CRITICAL: Only apply edits if user pressed Enter (not Escape)
            if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
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
