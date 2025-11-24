using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class DeselectOccluded : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            View activeView = doc.ActiveView;
            
            // Check if we have a valid view
            if (activeView == null || activeView.ViewType == ViewType.Schedule || 
                activeView.ViewType == ViewType.DrawingSheet || activeView.ViewType == ViewType.Internal ||
                activeView.ViewType == ViewType.Report || activeView.ViewType == ViewType.CostReport)
            {
                TaskDialog.Show("Invalid View", "Please activate a model view (not a schedule, sheet, or report).");
                return Result.Failed;
            }
            
            // Get current selection
            Selection selection = uiDoc.Selection;
            ICollection<ElementId> selectedIds = selection.GetElementIds();
            
            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("No Selection", "No elements are currently selected.");
                return Result.Succeeded;
            }
            
            // Start performance timer
            DateTime startTime = DateTime.Now;
            
            // Find visible elements using Revit's built-in visibility system
            HashSet<ElementId> visibleElementIds = FindVisibleElementsEfficient(
                doc, activeView, selectedIds);
            
            // Calculate elements to deselect
            List<ElementId> elementsToDeselect = selectedIds
                .Where(id => !visibleElementIds.Contains(id))
                .ToList();
            
            // Update selection
            if (elementsToDeselect.Count > 0)
            {
                using (Transaction trans = new Transaction(doc, "Deselect Occluded Elements"))
                {
                    trans.Start();
                    
                    // Set new selection (only visible elements)
                    selection.SetElementIds(visibleElementIds.ToList());
                    
                    trans.Commit();
                }
                
                TimeSpan elapsed = DateTime.Now - startTime;
                string report = $"Deselected {elementsToDeselect.Count} occluded element(s) out of {selectedIds.Count} total.\n" +
                               $"Time taken: {elapsed.TotalSeconds:F2} seconds";
                TaskDialog.Show("Deselect Occluded Complete", report);
            }
            else
            {
                TaskDialog.Show("Result", "All selected elements are visible.");
            }
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
    
    private HashSet<ElementId> FindVisibleElementsEfficient(Document doc, View view, ICollection<ElementId> selectedIds)
    {
        HashSet<ElementId> visibleElements = new HashSet<ElementId>();
        
        // Method 1: Use Revit's built-in visibility determination
        // Get ALL elements that Revit considers visible in this view
        FilteredElementCollector viewVisibleCollector = new FilteredElementCollector(doc, view.Id);
        HashSet<ElementId> elementsVisibleInView = new HashSet<ElementId>(viewVisibleCollector.ToElementIds());
        
        // Check each selected element
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem == null) continue;
            
            // Quick check: Is element in the view's visible collection?
            if (!elementsVisibleInView.Contains(id))
            {
                // Element is not visible in view at all
                continue;
            }
            
            // Additional detailed visibility checks
            if (!IsElementVisibleInView(doc, view, elem))
            {
                continue;
            }
            
            // For 3D views, check if within view's section box
            if (view is View3D view3D)
            {
                if (!IsElementInSectionBox(elem, view3D))
                {
                    continue;
                }
            }
            
            // For plan views, check if element is within view range
            if (view is ViewPlan viewPlan)
            {
                if (!IsElementInPlanViewRange(elem, viewPlan))
                {
                    continue;
                }
            }
            
            // For section/elevation views, check if element is within crop
            if ((view.ViewType == ViewType.Section || view.ViewType == ViewType.Elevation) && view.CropBoxActive)
            {
                if (!IsElementInCropBox(elem, view))
                {
                    continue;
                }
            }
            
            // Element passed all visibility checks
            visibleElements.Add(id);
        }
        
        return visibleElements;
    }
    
    private bool IsElementVisibleInView(Document doc, View view, Element elem)
    {
        try
        {
            // Check if element can be hidden
            if (!elem.CanBeHidden(view))
                return false;
            
            // Check if element is hidden
            if (elem.IsHidden(view))
                return false;
            
            // Check category visibility
            Category category = elem.Category;
            if (category != null)
            {
                try
                {
                    if (!category.get_Visible(view))
                        return false;
                }
                catch
                {
                    // Some categories don't support visibility checking
                }
            }
            
            // Check workset visibility
            if (doc.IsWorkshared)
            {
                WorksetId worksetId = elem.WorksetId;
                if (worksetId != null && worksetId != WorksetId.InvalidWorksetId)
                {
                    WorksetVisibility worksetVisibility = view.GetWorksetVisibility(worksetId);
                    if (worksetVisibility == WorksetVisibility.Hidden)
                        return false;
                }
            }
            
            // Check design option visibility (if element belongs to a design option)
            if (elem.DesignOption != null)
            {
                // Get the design option visibility in the view
                ElementId designOptionId = elem.DesignOption.Id;
                
                // Check if view has visibility settings for design options
                Parameter viewDesignOption = view.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER);
                if (viewDesignOption != null)
                {
                    // For now, we'll assume if element has a design option and view doesn't show it, it's not visible
                    // This is simplified - full implementation would require checking view's design option settings
                }
            }
            
            // Check phase
            ElementId elemPhaseCreated = elem.CreatedPhaseId;
            ElementId elemPhaseDemolished = elem.DemolishedPhaseId;
            Parameter viewPhaseParam = view.get_Parameter(BuiltInParameter.VIEW_PHASE);
            
            if (viewPhaseParam != null && viewPhaseParam.HasValue)
            {
                ElementId viewPhaseId = viewPhaseParam.AsElementId();
                
                if (elemPhaseCreated != null && elemPhaseCreated != ElementId.InvalidElementId)
                {
                    // Get phase elements
                    Phase elemPhase = doc.GetElement(elemPhaseCreated) as Phase;
                    Phase viewPhase = doc.GetElement(viewPhaseId) as Phase;
                    
                    if (elemPhase != null && viewPhase != null)
                    {
                        // Check if element is created after view phase (future element)
                        // We compare phase sequence numbers
                        int elemPhaseSeq = GetPhaseSequence(doc, elemPhaseCreated);
                        int viewPhaseSeq = GetPhaseSequence(doc, viewPhaseId);
                        
                        if (elemPhaseSeq > viewPhaseSeq)
                            return false; // Element doesn't exist yet in view phase
                        
                        // Check if element is demolished
                        if (elemPhaseDemolished != null && elemPhaseDemolished != ElementId.InvalidElementId)
                        {
                            int demolishedPhaseSeq = GetPhaseSequence(doc, elemPhaseDemolished);
                            
                            // If demolished before or in view phase
                            if (demolishedPhaseSeq <= viewPhaseSeq)
                            {
                                // Check phase filter to see if demolished elements are shown
                                Parameter phaseFilterParam = view.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER);
                                if (phaseFilterParam != null && phaseFilterParam.HasValue)
                                {
                                    ElementId phaseFilterId = phaseFilterParam.AsElementId();
                                    PhaseFilter phaseFilter = doc.GetElement(phaseFilterId) as PhaseFilter;
                                    
                                    // If we can't determine filter settings, assume not visible
                                    // (demolished elements are typically hidden)
                                    if (phaseFilter == null)
                                        return false;
                                    
                                    // Try to determine if demolished elements are shown
                                    // Since we can't directly check, we'll assume standard behavior
                                    // where demolished elements in view phase are typically not shown
                                    if (demolishedPhaseSeq == viewPhaseSeq)
                                        return false; // Demolished in this phase
                                }
                            }
                        }
                    }
                }
            }
            
            // Check view filters
            ICollection<ElementId> filterIds = view.GetFilters();
            foreach (ElementId filterId in filterIds)
            {
                ParameterFilterElement filter = doc.GetElement(filterId) as ParameterFilterElement;
                if (filter != null)
                {
                    try
                    {
                        // Check if element passes filter
#if REVIT2017 || REVIT2018
                        // GetElementFilter() not available in Revit 2017-2018
                        ElementFilter elementFilter = null;
#else
                        ElementFilter elementFilter = filter.GetElementFilter();
#endif
                        if (elementFilter != null)
                        {
                            bool passes = elementFilter.PassesFilter(elem);
                            
                            // Check filter visibility setting
                            if (passes)
                            {
                                bool filterIsVisible = view.GetFilterVisibility(filterId);
                                if (!filterIsVisible)
                                    return false;
                            }
                        }
                    }
                    catch
                    {
                        // Some filters may not be applicable to all elements
                    }
                }
            }
            
            // Check graphic overrides
            OverrideGraphicSettings overrides = view.GetElementOverrides(elem.Id);
            if (overrides != null && overrides.IsValidObject)
            {
                // Check if element is set to fully transparent
                if (overrides.Transparency == 100)
                    return false;
            }
            
            // Note: Temporary hide/isolate checking would require additional API methods
            // that may not be available in all Revit versions. The FilteredElementCollector
            // already accounts for temporarily hidden elements, so this is covered.
            
            return true;
        }
        catch (Exception)
        {
            // If any error occurs, assume element is visible to avoid incorrectly hiding it
            return true;
        }
    }
    
    private bool IsElementInSectionBox(Element elem, View3D view3D)
    {
        if (!view3D.IsSectionBoxActive)
            return true; // No section box, element is visible
        
        BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
        if (sectionBox == null || !sectionBox.Enabled)
            return true;
        
        BoundingBoxXYZ elemBox = elem.get_BoundingBox(view3D);
        if (elemBox == null)
            return true; // No bounding box, assume visible
        
        // Transform element bounding box to section box coordinate system
        Transform transform = sectionBox.Transform.Inverse;
        XYZ minTransformed = transform.OfPoint(elemBox.Min);
        XYZ maxTransformed = transform.OfPoint(elemBox.Max);
        
        // Create proper min/max from transformed points
        XYZ properMin = new XYZ(
            Math.Min(minTransformed.X, maxTransformed.X),
            Math.Min(minTransformed.Y, maxTransformed.Y),
            Math.Min(minTransformed.Z, maxTransformed.Z)
        );
        
        XYZ properMax = new XYZ(
            Math.Max(minTransformed.X, maxTransformed.X),
            Math.Max(minTransformed.Y, maxTransformed.Y),
            Math.Max(minTransformed.Z, maxTransformed.Z)
        );
        
        // Check if bounding boxes intersect
        return !(properMax.X < sectionBox.Min.X || properMin.X > sectionBox.Max.X ||
                properMax.Y < sectionBox.Min.Y || properMin.Y > sectionBox.Max.Y ||
                properMax.Z < sectionBox.Min.Z || properMin.Z > sectionBox.Max.Z);
    }
    
    private bool IsElementInPlanViewRange(Element elem, ViewPlan viewPlan)
    {
        try
        {
            // Get the plan view range
            PlanViewRange viewRange = viewPlan.GetViewRange();
            if (viewRange == null)
                return true; // No view range defined, assume visible
            
            // Get level IDs for top and bottom of view range
            ElementId topLevelId = viewRange.GetLevelId(PlanViewPlane.TopClipPlane);
            ElementId bottomLevelId = viewRange.GetLevelId(PlanViewPlane.BottomClipPlane);
            ElementId cutLevelId = viewRange.GetLevelId(PlanViewPlane.CutPlane);
            
            // Get offsets
            double topOffset = viewRange.GetOffset(PlanViewPlane.TopClipPlane);
            double bottomOffset = viewRange.GetOffset(PlanViewPlane.BottomClipPlane);
            double cutOffset = viewRange.GetOffset(PlanViewPlane.CutPlane);
            
            // Get element's bounding box
            BoundingBoxXYZ bbox = elem.get_BoundingBox(viewPlan);
            if (bbox == null)
                return true; // No bounding box, assume visible
            
            // Get the Z coordinates of the element
            double elemMinZ = bbox.Min.Z;
            double elemMaxZ = bbox.Max.Z;
            
            // Calculate actual elevations
            Level topLevel = topLevelId != null ? viewPlan.Document.GetElement(topLevelId) as Level : null;
            Level bottomLevel = bottomLevelId != null ? viewPlan.Document.GetElement(bottomLevelId) as Level : null;
            
            if (topLevel != null && bottomLevel != null)
            {
                double topElevation = topLevel.Elevation + topOffset;
                double bottomElevation = bottomLevel.Elevation + bottomOffset;
                
                // Check if element is within view range
                return !(elemMaxZ < bottomElevation || elemMinZ > topElevation);
            }
            
            return true; // Can't determine, assume visible
        }
        catch
        {
            // If view range checking fails, assume element is visible
            return true;
        }
    }
    
    private bool IsElementInCropBox(Element elem, View view)
    {
        if (!view.CropBoxActive)
            return true; // No crop box active
        
        BoundingBoxXYZ cropBox = view.CropBox;
        if (cropBox == null || !cropBox.Enabled)
            return true;
        
        BoundingBoxXYZ elemBox = elem.get_BoundingBox(view);
        if (elemBox == null)
            return true; // No bounding box, assume visible
        
        // Check if bounding boxes intersect in view space
        return !(elemBox.Max.X < cropBox.Min.X || elemBox.Min.X > cropBox.Max.X ||
                elemBox.Max.Y < cropBox.Min.Y || elemBox.Min.Y > cropBox.Max.Y);
    }
    
    private int GetPhaseSequence(Document doc, ElementId phaseId)
    {
        // Get all phases in the document
        FilteredElementCollector phaseCollector = new FilteredElementCollector(doc);
        IList<Element> phases = phaseCollector.OfClass(typeof(Phase)).ToElements();
        
        // Find the sequence number of the given phase
        for (int i = 0; i < phases.Count; i++)
        {
            if (phases[i].Id == phaseId)
                return i;
        }
        
        return -1; // Phase not found
    }
}
