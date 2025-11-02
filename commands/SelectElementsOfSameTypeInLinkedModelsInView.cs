using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectElementsOfSameTypeInLinkedModelsInView : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;
        UIDocument uiDoc = uiapp.ActiveUIDocument;
        View activeView = uiapp.ActiveUIDocument.ActiveView;
        
        // Get currently selected elements using SelectionModeManager methods
        var selectedIds = uiDoc.GetSelectionIds();
        var selectedRefs = uiDoc.GetReferences();
        
        if (!selectedIds.Any() && !selectedRefs.Any())
        {
            TaskDialog.Show("Info", "No elements are selected.");
            return Result.Cancelled;
        }
        
        // Collect type information from selected elements
        var selectedTypes = new HashSet<string>(); // Use type signatures for comparison
        var typeDescriptions = new List<string>(); // For user feedback
        var selectedLinkIds = new HashSet<ElementId>(); // Track link instance IDs that contain selected elements
        
        // Process regular elements from host document
        foreach (var id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element != null && !(element is RevitLinkInstance))
            {
                string typeSignature = GetElementTypeSignature(element);
                if (!string.IsNullOrEmpty(typeSignature))
                {
                    selectedTypes.Add(typeSignature);
                    typeDescriptions.Add(GetElementTypeDescription(element));
                }
            }
        }
        
        // Process linked elements from references
        foreach (var reference in selectedRefs)
        {
            try
            {
                if (reference.LinkedElementId != ElementId.InvalidElementId)
                {
                    // This is a linked element reference
                    var linkedInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkedInstance != null)
                    {
                        // Track this link instance
                        selectedLinkIds.Add(reference.ElementId);
                        
                        Document linkedDoc = linkedInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            Element linkedElement = linkedDoc.GetElement(reference.LinkedElementId);
                            if (linkedElement != null)
                            {
                                string typeSignature = GetElementTypeSignature(linkedElement);
                                if (!string.IsNullOrEmpty(typeSignature))
                                {
                                    selectedTypes.Add(typeSignature);
                                    typeDescriptions.Add(GetElementTypeDescription(linkedElement));
                                }
                            }
                        }
                    }
                }
                else if (reference.ElementId != ElementId.InvalidElementId)
                {
                    // Regular element reference
                    Element element = doc.GetElement(reference.ElementId);
                    if (element != null && !(element is RevitLinkInstance))
                    {
                        string typeSignature = GetElementTypeSignature(element);
                        if (!string.IsNullOrEmpty(typeSignature))
                        {
                            selectedTypes.Add(typeSignature);
                            typeDescriptions.Add(GetElementTypeDescription(element));
                        }
                    }
                }
            }
            catch { /* Skip problematic references */ }
        }
        
        if (!selectedTypes.Any())
        {
            TaskDialog.Show("Info", "No valid element types found in selection.");
            return Result.Cancelled;
        }
        
        // Get all loaded linked models
        var allLinkInstances = new FilteredElementCollector(doc)
            .OfClass(typeof(RevitLinkInstance))
            .Cast<RevitLinkInstance>()
            .Where(link => link.GetLinkDocument() != null)
            .ToList();
        
        if (!allLinkInstances.Any())
        {
            TaskDialog.Show("Info", "No loaded linked models found in the project.");
            return Result.Cancelled;
        }
        
        // Create dictionary entries for the DataGrid
        var linkEntries = new List<Dictionary<string, object>>();
        var linkToIndexMap = new Dictionary<ElementId, int>(); // Map link ElementId to index
        
        // Sort link instances by name first
        allLinkInstances = allLinkInstances.OrderBy(link => link.Name).ToList();
        
        for (int i = 0; i < allLinkInstances.Count; i++)
        {
            var link = allLinkInstances[i];
            var entry = new Dictionary<string, object>
            {
                { "Name", link.Name },
                { "LinkType", link.GetTypeId() != ElementId.InvalidElementId ? 
                    doc.GetElement(link.GetTypeId())?.Name ?? "Unknown" : "Unknown" }
            };
            linkEntries.Add(entry);
            linkToIndexMap[link.Id] = i;
        }
        
        // Determine which models should be pre-selected based on selectedLinkIds
        var preSelectedIndices = new List<int>();
        foreach (var linkId in selectedLinkIds)
        {
            if (linkToIndexMap.ContainsKey(linkId))
            {
                preSelectedIndices.Add(linkToIndexMap[linkId]);
            }
        }
        
        // Define properties to display
        var linkPropertyNames = new List<string> { "Name", "LinkType" };
        
        // Show the DataGrid to let the user select linked models to search
        var selectedEntries = CustomGUIs.DataGrid(
            linkEntries,
            linkPropertyNames,
            false, // spanAllScreens
            preSelectedIndices.Count > 0 ? preSelectedIndices : null
        );
        
        if (selectedEntries == null || selectedEntries.Count == 0)
            return Result.Cancelled;
        
        // Extract the selected link instances based on selected entries
        var selectedLinkInstances = new List<RevitLinkInstance>();
        foreach (var entry in selectedEntries)
        {
            var linkName = entry["Name"].ToString();
            var matchingLink = allLinkInstances.FirstOrDefault(link => link.Name == linkName);
            if (matchingLink != null)
            {
                selectedLinkInstances.Add(matchingLink);
            }
        }
        
        List<Reference> matchingReferences = new List<Reference>();
        int linksSearched = 0;
        int totalElementsFound = 0;
        int elementsInViewRange = 0;
        
        // Search only the user-selected linked models for matching element types
        foreach (var linkInstance in selectedLinkInstances)
        {
            Document linkedDoc = linkInstance.GetLinkDocument();
            if (linkedDoc == null) continue;
            
            linksSearched++;
            
            FilteredElementCollector collector = new FilteredElementCollector(linkedDoc);
            collector.WhereElementIsNotElementType();
            
            foreach (Element elem in collector)
            {
                try
                {
                    string elemTypeSignature = GetElementTypeSignature(elem);
                    if (selectedTypes.Contains(elemTypeSignature))
                    {
                        totalElementsFound++;
                        
                        // Check if element's bounding box is within current view range
                        bool isInViewRange = IsElementBoundingBoxInCurrentView(elem, linkInstance, activeView);
                        if (!isInViewRange)
                            continue;
                        
                        elementsInViewRange++;
                        
                        Reference elemRef = new Reference(elem);
                        Reference linkedRef = elemRef.CreateLinkReference(linkInstance);
                        
                        if (linkedRef != null)
                        {
                            matchingReferences.Add(linkedRef);
                        }
                    }
                }
                catch
                {
                    // Some elements might not support reference creation
                    continue;
                }
            }
        }
        
        if (matchingReferences.Count > 0)
        {
            // Select all matching elements using SelectionModeManager
            uiDoc.SetReferences(matchingReferences);
        }
        else
        {
            string failMessage = $"No matching elements found in current view range.\n\n" +
                                $"Total matching elements found: {totalElementsFound}\n" +
                                $"Elements in view range: {elementsInViewRange}\n\n" +
                                $"Selected types:\n";
            foreach (var typeDesc in typeDescriptions.Distinct())
            {
                failMessage += $"- {typeDesc}\n";
            }
            
            TaskDialog.Show("Info", failMessage);
        }
        
        return Result.Succeeded;
    }
    
    private bool IsElementBoundingBoxInCurrentView(Element linkedElement, RevitLinkInstance linkInstance, View currentView)
    {
        try
        {
            // Get element's bounding box in linked model coordinates
            BoundingBoxXYZ elementBB = linkedElement.get_BoundingBox(null);
            if (elementBB == null) return false;
            
            // Transform to host coordinates
            Transform linkTransform = linkInstance.GetTotalTransform();
            XYZ transformedMin = linkTransform.OfPoint(elementBB.Min);
            XYZ transformedMax = linkTransform.OfPoint(elementBB.Max);
            
            // Check based on view type
            if (currentView.ViewType == ViewType.FloorPlan || 
                currentView.ViewType == ViewType.CeilingPlan ||
                currentView.ViewType == ViewType.AreaPlan ||
                currentView.ViewType == ViewType.EngineeringPlan)
            {
                return IsInPlanViewRange(transformedMin, transformedMax, currentView as ViewPlan);
            }
            else if (currentView.ViewType == ViewType.ThreeD)
            {
                return IsIn3DViewRange(transformedMin, transformedMax, currentView as View3D);
            }
            else if (currentView.ViewType == ViewType.Section || currentView.ViewType == ViewType.Elevation)
            {
                return IsInSectionElevationViewRange(transformedMin, transformedMax, currentView);
            }
            
            return false; // Default to not visible for other view types
        }
        catch
        {
            return false;
        }
    }
    
    private bool IsInPlanViewRange(XYZ transformedMin, XYZ transformedMax, ViewPlan planView)
    {
        try
        {
            // Get the actual view range (handles inheritance from view templates)
            PlanViewRange viewRange = planView.GetViewRange();
            if (viewRange == null) return false;
            
            Level viewLevel = planView.GenLevel;
            if (viewLevel == null) return false;
            
            double levelElevation = viewLevel.ProjectElevation;
            
            // Get the actual levels and offsets for each plane
            // Top Clip Plane
            ElementId topLevelId = viewRange.GetLevelId(PlanViewPlane.TopClipPlane);
            Level topLevel = planView.Document.GetElement(topLevelId) as Level;
            double topLevelElevation = topLevel?.ProjectElevation ?? levelElevation;
            double topOffset = viewRange.GetOffset(PlanViewPlane.TopClipPlane);
            double topElevation = topLevelElevation + topOffset;
            
            // Cut Plane (used for visibility determination)
            ElementId cutLevelId = viewRange.GetLevelId(PlanViewPlane.CutPlane);
            Level cutLevel = planView.Document.GetElement(cutLevelId) as Level;
            double cutLevelElevation = cutLevel?.ProjectElevation ?? levelElevation;
            double cutOffset = viewRange.GetOffset(PlanViewPlane.CutPlane);
            double cutElevation = cutLevelElevation + cutOffset;
            
            // Bottom Clip Plane
            ElementId bottomLevelId = viewRange.GetLevelId(PlanViewPlane.BottomClipPlane);
            Level bottomLevel = planView.Document.GetElement(bottomLevelId) as Level;
            double bottomLevelElevation = bottomLevel?.ProjectElevation ?? levelElevation;
            double bottomOffset = viewRange.GetOffset(PlanViewPlane.BottomClipPlane);
            double bottomElevation = bottomLevelElevation + bottomOffset;
            
            // View Depth (elements below cut plane but above view depth are shown)
            ElementId viewDepthLevelId = viewRange.GetLevelId(PlanViewPlane.ViewDepthPlane);
            Level viewDepthLevel = planView.Document.GetElement(viewDepthLevelId) as Level;
            double viewDepthLevelElevation = viewDepthLevel?.ProjectElevation ?? levelElevation;
            double viewDepthOffset = viewRange.GetOffset(PlanViewPlane.ViewDepthPlane);
            double viewDepthElevation = viewDepthLevelElevation + viewDepthOffset;
            
            // Check if element intersects with view range
            // An element is visible if:
            // 1. It's between bottom and top clip planes AND
            // 2. Either above cut plane OR above view depth
            double elementMinZ = transformedMin.Z;
            double elementMaxZ = transformedMax.Z;
            
            bool inClipRange = (elementMaxZ >= bottomElevation && elementMinZ <= topElevation);
            bool aboveCutPlane = (elementMaxZ >= cutElevation);
            bool aboveViewDepth = (elementMaxZ >= viewDepthElevation);
            
            bool inViewRange = inClipRange && (aboveCutPlane || aboveViewDepth);
            
            if (!inViewRange) return false;
            
            // Check crop region horizontally
            if (planView.CropBoxActive)
            {
                BoundingBoxXYZ cropBox = planView.CropBox;
                if (cropBox == null) return false;
                
                // Transform crop box to world coordinates
                Transform viewTransform = cropBox.Transform;
                XYZ cropMin = viewTransform.OfPoint(cropBox.Min);
                XYZ cropMax = viewTransform.OfPoint(cropBox.Max);
                
                // Check intersection in X and Y only
                bool inCropRegion = !(transformedMax.X < cropMin.X || transformedMin.X > cropMax.X ||
                                     transformedMax.Y < cropMin.Y || transformedMin.Y > cropMax.Y);
                return inCropRegion;
            }
            
            return true; // In view range and no crop box to check
        }
        catch
        {
            return false;
        }
    }
    
    private bool IsIn3DViewRange(XYZ transformedMin, XYZ transformedMax, View3D view3D)
    {
        try
        {
            // Check section box if active
            if (view3D.IsSectionBoxActive)
            {
                BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
                if (sectionBox == null) return false;
                
                // Transform section box to world coordinates
                Transform sectionTransform = sectionBox.Transform;
                XYZ sectionMin = sectionTransform.OfPoint(sectionBox.Min);
                XYZ sectionMax = sectionTransform.OfPoint(sectionBox.Max);
                
                // Create axis-aligned bounds
                double minX = System.Math.Min(sectionMin.X, sectionMax.X);
                double maxX = System.Math.Max(sectionMin.X, sectionMax.X);
                double minY = System.Math.Min(sectionMin.Y, sectionMax.Y);
                double maxY = System.Math.Max(sectionMin.Y, sectionMax.Y);
                double minZ = System.Math.Min(sectionMin.Z, sectionMax.Z);
                double maxZ = System.Math.Max(sectionMin.Z, sectionMax.Z);
                
                // Check intersection
                bool inSectionBox = !(transformedMax.X < minX || transformedMin.X > maxX ||
                                     transformedMax.Y < minY || transformedMin.Y > maxY ||
                                     transformedMax.Z < minZ || transformedMin.Z > maxZ);
                return inSectionBox;
            }
            
            // Check crop box if active (for regular 3D views)
            if (view3D.CropBoxActive)
            {
                BoundingBoxXYZ cropBox = view3D.CropBox;
                if (cropBox == null) return false;
                
                Transform cropTransform = cropBox.Transform;
                XYZ cropMin = cropTransform.OfPoint(cropBox.Min);
                XYZ cropMax = cropTransform.OfPoint(cropBox.Max);
                
                // Create axis-aligned bounds
                double minX = System.Math.Min(cropMin.X, cropMax.X);
                double maxX = System.Math.Max(cropMin.X, cropMax.X);
                double minY = System.Math.Min(cropMin.Y, cropMax.Y);
                double maxY = System.Math.Max(cropMin.Y, cropMax.Y);
                double minZ = System.Math.Min(cropMin.Z, cropMax.Z);
                double maxZ = System.Math.Max(cropMin.Z, cropMax.Z);
                
                bool inCropBox = !(transformedMax.X < minX || transformedMin.X > maxX ||
                                  transformedMax.Y < minY || transformedMin.Y > maxY ||
                                  transformedMax.Z < minZ || transformedMin.Z > maxZ);
                return inCropBox;
            }
            
            return true; // No section box or crop box, so assume visible
        }
        catch
        {
            return false;
        }
    }
    
    private bool IsInSectionElevationViewRange(XYZ transformedMin, XYZ transformedMax, View view)
    {
        try
        {
            // Check crop box if active
            if (view.CropBoxActive)
            {
                BoundingBoxXYZ cropBox = view.CropBox;
                if (cropBox == null) return false;
                
                // Get view direction and transform
                ViewSection viewSection = view as ViewSection;
                Transform viewTransform = cropBox.Transform;
                
                // Transform the element bounds to view coordinates
                Transform inverse = viewTransform.Inverse;
                XYZ minInView = inverse.OfPoint(transformedMin);
                XYZ maxInView = inverse.OfPoint(transformedMax);
                
                // Create axis-aligned bounds in view space
                double minX = System.Math.Min(minInView.X, maxInView.X);
                double maxX = System.Math.Max(minInView.X, maxInView.X);
                double minY = System.Math.Min(minInView.Y, maxInView.Y);
                double maxY = System.Math.Max(minInView.Y, maxInView.Y);
                double minZ = System.Math.Min(minInView.Z, maxInView.Z);
                double maxZ = System.Math.Max(minInView.Z, maxInView.Z);
                
                // Check if within crop box bounds
                // For sections/elevations, typically check X and Y (horizontal and vertical in view)
                // Z is the depth into/out of the view
                bool inCropBox = !(maxX < cropBox.Min.X || minX > cropBox.Max.X ||
                                  maxY < cropBox.Min.Y || minY > cropBox.Max.Y);
                
                // Also check view depth if it's a section
                if (viewSection != null)
                {
                    // Check if element is within the far clip offset
                    Parameter farClipParam = view.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR);
                    if (farClipParam != null)
                    {
                        double farClipOffset = farClipParam.AsDouble();
                        inCropBox = inCropBox && (minZ <= farClipOffset);
                    }
                }
                
                return inCropBox;
            }
            
            return true; // No crop box, so assume visible
        }
        catch
        {
            return false;
        }
    }
    
    /// <summary>
    /// Creates a unique signature for an element type that can be used for cross-document comparison
    /// </summary>
    private string GetElementTypeSignature(Element element)
    {
        try
        {
            // Get the element's type
            ElementId typeId = element.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                ElementType elementType = element.Document.GetElement(typeId) as ElementType;
                if (elementType != null)
                {
                    // For loadable families, use family name + type name
                    if (!string.IsNullOrEmpty(elementType.FamilyName))
                    {
                        return $"Family:{elementType.FamilyName}:{elementType.Name}";
                    }
                    else
                    {
                        // For system families, use category + type name
                        string categoryName = element.Category?.Name ?? "Unknown";
                        return $"System:{categoryName}:{elementType.Name}";
                    }
                }
            }
            
            // Fallback for elements without types (like some annotation elements)
            if (element.Category != null)
            {
                // Use category + some identifying parameters
                string categoryName = element.Category.Name;
                
                // Try to get some distinguishing parameters
                string additionalInfo = "";
                
                // For text elements, include text style
                if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_TextNotes)
                {
                    var styleParam = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                    if (styleParam != null)
                    {
                        var textStyle = element.Document.GetElement(styleParam.AsElementId()) as TextNoteType;
                        additionalInfo = textStyle?.Name ?? "";
                    }
                }
                // For dimensions, include dimension style
                else if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Dimensions)
                {
                    var styleParam = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                    if (styleParam != null)
                    {
                        var dimStyle = element.Document.GetElement(styleParam.AsElementId()) as DimensionType;
                        additionalInfo = dimStyle?.Name ?? "";
                    }
                }
                
                return $"NoType:{categoryName}:{additionalInfo}:{element.GetType().Name}";
            }
            
            // Last resort fallback
            return $"Unknown:{element.GetType().Name}";
        }
        catch
        {
            return string.Empty;
        }
    }
    
    /// <summary>
    /// Gets a human-readable description of the element type for user feedback
    /// </summary>
    private string GetElementTypeDescription(Element element)
    {
        try
        {
            ElementId typeId = element.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                ElementType elementType = element.Document.GetElement(typeId) as ElementType;
                if (elementType != null)
                {
                    if (!string.IsNullOrEmpty(elementType.FamilyName))
                    {
                        return $"{elementType.FamilyName}: {elementType.Name}";
                    }
                    else
                    {
                        return $"{element.Category?.Name ?? "Unknown"}: {elementType.Name}";
                    }
                }
            }
            
            // Fallback
            return element.Category?.Name ?? element.GetType().Name;
        }
        catch
        {
            return "Unknown Type";
        }
    }
}
