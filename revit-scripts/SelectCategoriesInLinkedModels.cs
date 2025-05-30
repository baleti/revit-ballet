using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

// Extended wrapper class for category information with link context
public class LinkedCategoryWrapper
{
    public ElementId CategoryId { get; set; }
    public string CategoryName { get; set; }
    public string LinkName { get; set; }
    public RevitLinkInstance LinkInstance { get; set; }
    
    // Display name for the DataGrid
    public string Name => $"{CategoryName} [{LinkName}]";
    
    public LinkedCategoryWrapper(ElementId categoryId, string categoryName, string linkName, RevitLinkInstance linkInstance)
    {
        CategoryId = categoryId;
        CategoryName = categoryName;
        LinkName = linkName;
        LinkInstance = linkInstance;
    }
}

// Wrapper class for linked model selection
public class LinkedModelWrapper
{
    public RevitLinkInstance LinkInstance { get; set; }
    public string Name { get; set; }
    public string LinkType { get; set; }
    
    public LinkedModelWrapper(RevitLinkInstance linkInstance)
    {
        LinkInstance = linkInstance;
        Name = linkInstance.Name;
        
        // Get the link type name
        var linkType = linkInstance.Document.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
        LinkType = linkType?.Name ?? "Unknown";
    }
}

public abstract class SelectCategoriesInSelectedLinkedModelsBase : IExternalCommand
{
    public abstract bool selectInView { get; }
    
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
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;
        View activeView = uiapp.ActiveUIDocument.ActiveView;
        UIDocument uiDoc = uiapp.ActiveUIDocument;
        
        // Get currently selected linked models
        var selectedLinkInstances = new List<RevitLinkInstance>();
        var currentSelection = uiDoc.Selection.GetElementIds();
        
        foreach (ElementId id in currentSelection)
        {
            Element elem = doc.GetElement(id);
            if (elem is RevitLinkInstance linkInstance && linkInstance.GetLinkDocument() != null)
            {
                selectedLinkInstances.Add(linkInstance);
            }
        }
        
        // If no linked models are selected, show a selection dialog
        if (!selectedLinkInstances.Any())
        {
            // Collect all loaded RevitLinkInstances
            var allLinkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => link.GetLinkDocument() != null) // Only loaded links
                .ToList();
            
            if (!allLinkInstances.Any())
            {
                TaskDialog.Show("Info", "No loaded linked models found in the project.");
                return Result.Cancelled;
            }
            
            // Create wrappers for the DataGrid
            var linkWrappers = allLinkInstances
                .Select(link => new LinkedModelWrapper(link))
                .OrderBy(w => w.Name)
                .ToList();
            
            // Define properties to display
            var linkPropertyNames = new List<string> { "Name", "LinkType" };
            
            // Show the DataGrid to let the user select linked models
            var selectedLinkWrappers = CustomGUIs.DataGrid<LinkedModelWrapper>(linkWrappers, linkPropertyNames);
            if (selectedLinkWrappers.Count == 0)
                return Result.Cancelled;
            
            // Extract the selected link instances
            selectedLinkInstances = selectedLinkWrappers.Select(w => w.LinkInstance).ToList();
        }
        
        // Build a list of categories across selected linked models
        // Show ALL categories regardless of mode, filtering happens during element selection
        List<LinkedCategoryWrapper> allLinkedCategories = new List<LinkedCategoryWrapper>();
        
        foreach (var linkInstance in selectedLinkInstances)
        {
            Document linkedDoc = linkInstance.GetLinkDocument();
            if (linkedDoc == null) continue;
            
            string linkName = linkInstance.Name;
            
            // Get all categories from the linked model
            FilteredElementCollector collector = new FilteredElementCollector(linkedDoc);
            collector.WhereElementIsNotElementType();
            
            HashSet<ElementId> categoryIds = new HashSet<ElementId>();
            foreach (Element elem in collector)
            {
                if (elem.Category != null)
                {
                    categoryIds.Add(elem.Category.Id);
                }
            }
            
            // Build category wrappers for this link
            foreach (ElementId id in categoryIds)
            {
                Category cat = Category.GetCategory(linkedDoc, id);
                if (cat != null)
                {
                    allLinkedCategories.Add(new LinkedCategoryWrapper(cat.Id, cat.Name, linkName, linkInstance));
                }
            }
            
            // Handle Views category separately
            var viewElements = new FilteredElementCollector(linkedDoc)
                .OfCategory(BuiltInCategory.OST_Viewers)
                .WhereElementIsNotElementType()
                .ToList();
            
            if (viewElements.Any())
            {
                var viewFamilies = viewElements
                    .GroupBy(e => e.get_Parameter(BuiltInParameter.VIEW_FAMILY)?.AsValueString() ?? "Unknown")
                    .Select(g => g.Key)
                    .ToList();
                
                foreach (string familyName in viewFamilies)
                {
                    allLinkedCategories.Add(new LinkedCategoryWrapper(
                        new ElementId((int)BuiltInCategory.OST_Viewers), 
                        "Views: " + familyName, 
                        linkName, 
                        linkInstance));
                }
            }
        }
        
        if (!allLinkedCategories.Any())
        {
            TaskDialog.Show("Info", "No categories found in selected linked models.");
            return Result.Cancelled;
        }
        
        // Sort categories by name for better user experience
        allLinkedCategories = allLinkedCategories.OrderBy(c => c.Name).ToList();
        
        // Define properties to display
        var propertyNames = new List<string> { "Name" };
        
        // Show the DataGrid to let the user select categories
        List<LinkedCategoryWrapper> selectedCategories = CustomGUIs.DataGrid<LinkedCategoryWrapper>(allLinkedCategories, propertyNames);
        if (selectedCategories.Count == 0)
            return Result.Cancelled;
        
        // Gather references for elements in selected categories
        List<Reference> elementReferences = new List<Reference>();
        int totalElementsInCategory = 0;
        int elementsInViewRange = 0;
        int elementsSuccessfullySelected = 0;
        
        foreach (LinkedCategoryWrapper selectedCategory in selectedCategories)
        {
            Document linkedDoc = selectedCategory.LinkInstance.GetLinkDocument();
            if (linkedDoc == null) continue;
            
            FilteredElementCollector categoryCollector = new FilteredElementCollector(linkedDoc);
            categoryCollector.WhereElementIsNotElementType();
            
            List<Element> elementsToCheck;
            
            // Handle split Views categories
            if (selectedCategory.CategoryId.IntegerValue == (int)BuiltInCategory.OST_Viewers &&
                selectedCategory.CategoryName.StartsWith("Views: "))
            {
                string familyName = selectedCategory.CategoryName.Substring("Views: ".Length);
                elementsToCheck = categoryCollector
                    .OfCategory(BuiltInCategory.OST_Viewers)
                    .Where(e => (e.get_Parameter(BuiltInParameter.VIEW_FAMILY)?.AsValueString() ?? "Unknown") == familyName)
                    .ToList();
            }
            else
            {
                elementsToCheck = categoryCollector
                    .OfCategory((BuiltInCategory)selectedCategory.CategoryId.IntegerValue)
                    .ToList();
            }
            
            totalElementsInCategory += elementsToCheck.Count;
            
            // Check each element's bounding box against current view range
            foreach (Element elem in elementsToCheck)
            {
                try
                {
                    // For in-view mode, check if element's bounding box is within current view range
                    if (selectInView)
                    {
                        bool isInViewRange = IsElementBoundingBoxInCurrentView(elem, selectedCategory.LinkInstance, activeView);
                        if (!isInViewRange)
                            continue;
                        
                        elementsInViewRange++;
                    }
                    
                    // Create reference for the element
                    Reference elemRef = new Reference(elem);
                    Reference linkedRef = elemRef.CreateLinkReference(selectedCategory.LinkInstance);
                    
                    if (linkedRef != null)
                    {
                        elementReferences.Add(linkedRef);
                        elementsSuccessfullySelected++;
                    }
                }
                catch (System.Exception ex)
                {
                    // Log the exception for debugging
                    System.Diagnostics.Debug.WriteLine($"Error processing element {elem.Id}: {ex.Message}");
                    continue;
                }
            }
        }
        
        if (elementReferences.Count > 0)
        {
            // Highlight all selected elements using references
            uiDoc.Selection.SetReferences(elementReferences);
            
            string resultMessage;
            if (selectInView)
            {
                resultMessage = $"Selected {elementsSuccessfullySelected} elements whose bounding boxes intersect with current view range.\n" +
                               $"Total elements in categories: {totalElementsInCategory}\n" +
                               $"Elements in view range: {elementsInViewRange}";
            }
            else
            {
                resultMessage = $"Selected {elementsSuccessfullySelected} elements from {selectedLinkInstances.Count} linked model(s).";
            }
            
            TaskDialog.Show("Success", resultMessage);
        }
        else
        {
            string failMessage;
            if (selectInView)
            {
                failMessage = $"No elements from selected categories have bounding boxes that intersect with current view range.\n" +
                             $"Total elements checked: {totalElementsInCategory}\n" +
                             $"Elements in view range: {elementsInViewRange}";
            }
            else
            {
                failMessage = "Found elements but couldn't create references for selection.";
            }
            
            TaskDialog.Show("Info", failMessage);
        }
        
        return Result.Succeeded;
    }
}

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectCategoriesInViewOfSelectedLinkedModels : SelectCategoriesInSelectedLinkedModelsBase
{
    public override bool selectInView => true;
}

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectCategoriesInProjectOfSelectedLinkedModels : SelectCategoriesInSelectedLinkedModelsBase
{
    public override bool selectInView => false;
}
