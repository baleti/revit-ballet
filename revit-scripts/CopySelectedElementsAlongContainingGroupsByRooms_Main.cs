using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using System.Windows.Forms;
using System.Threading;

[Transaction(TransactionMode.Manual)]
public partial class CopySelectedElementsAlongContainingGroupsByRooms : IExternalCommand
{
    // Track room validation statistics
    private int totalRoomsChecked = 0;
    private int roomsValidatedBySimilarity = 0;
    private int roomsInvalidatedByDissimilarity = 0;
    private int roomsAsDirectMembers = 0;
    
    // Progress tracking
    private CopyElementsProgressForm progressForm = null;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        bool operationCancelled = false;
        try
        {
            InitializeDiagnostics();

            // Get and validate selected elements
            List<Element> selectedElements = GetAndValidateSelection(uidoc, doc, ref message);
            if (selectedElements == null || selectedElements.Count == 0)
                return Result.Failed;

            LogSelectedElements(selectedElements);
            
            // Initialize progress form
            progressForm = new CopyElementsProgressForm();
            progressForm.Show();
            Application.DoEvents();
            progressForm.UpdateElementCounts(selectedElements.Count, 0, 0);
            progressForm.SetPhase("Initialization");
            progressForm.UpdateProgress(0, 100, "Preparing elements...");
            progressForm.UpdateElementCounts(selectedElements.Count, 0, 0);
            Application.DoEvents();

            // Pre-calculate test points for all selected elements
            foreach (Element elem in selectedElements)
            {
                _elementTestPointsCache[elem.Id] = GetElementTestPoints(elem);
            }

            if (CheckCancellation())
            {
                operationCancelled = true;
                return Result.Cancelled;
            }

            // Create bounding box for all selected elements for quick filtering
            BoundingBoxXYZ overallBB = GetOverallBoundingBox(selectedElements);
            if (overallBB == null)
            {
                message = "Could not determine bounding box of selected elements";
                return Result.Failed;
            }

            progressForm.SetPhase("Finding Groups");
            progressForm.UpdateProgress(10, 100, "Collecting all groups in document...");

            // Find all groups in the document
            FilteredElementCollector groupCollector = new FilteredElementCollector(doc);
            IList<Group> allGroups = groupCollector
                .OfClass(typeof(Group))
                .WhereElementIsNotElementType()
                .Cast<Group>()
                .ToList();

            progressForm.AddDetail($"Found {allGroups.Count} groups in document", CopyElementsProgressForm.DetailType.Info);

            // Build room cache for ALL rooms in the document
            progressForm.SetPhase("Building Room Cache");
            progressForm.UpdateProgress(20, 100, "Caching room data...");
            Application.DoEvents();
            
            BuildRoomCache(doc);

            // OPTIMIZATION: Clear caches from previous runs
            _floorToFloorHeightCache.Clear();
            _roomPointContainmentCache.Clear();
            _roomBoundaryCache.Clear();

            if (CheckCancellation())
            {
                operationCancelled = true;
                return Result.Cancelled;
            }

            // PRE-CALCULATE ALL GROUP BOUNDING BOXES AND BUILD SPATIAL INDEX
            progressForm.SetPhase("Building Spatial Index");
            progressForm.UpdateProgress(30, 100, "Pre-calculating group bounding boxes...");
            Application.DoEvents();
            
            PreCalculateGroupDataAndSpatialIndex(allGroups, doc);

            // Get potentially relevant groups using spatial index
            progressForm.UpdateProgress(40, 100, "Filtering spatially relevant groups...");
            List<Group> spatiallyRelevantGroups = GetSpatiallyRelevantGroups(overallBB);
            progressForm.AddDetail($"Filtered to {spatiallyRelevantGroups.Count} spatially relevant groups", 
                CopyElementsProgressForm.DetailType.Info);

            // Build comprehensive room-to-group mapping with priority handling and validation
            Dictionary<ElementId, Group> roomToSingleGroupMap = BuildRoomToGroupMapping(spatiallyRelevantGroups, doc, overallBB);

            // Map elements to their containing groups
            // OPTIMIZATION: Pre-filter groups by element elevation ranges
            double minElementZ = double.MaxValue;
            double maxElementZ = double.MinValue;
            
            foreach (Element elem in selectedElements)
            {
                List<XYZ> points = _elementTestPointsCache[elem.Id];
                foreach (XYZ pt in points)
                {
                    minElementZ = Math.Min(minElementZ, pt.Z);
                    maxElementZ = Math.Max(maxElementZ, pt.Z);
                }
            }
            
            // Filter groups by elevation
            List<Group> elevationFilteredGroups = new List<Group>();
            foreach (Group group in spatiallyRelevantGroups)
            {
                if (_groupBoundingBoxCache.ContainsKey(group.Id))
                {
                    BoundingBoxXYZ bb = _groupBoundingBoxCache[group.Id];
                    // Increased tolerance from 2.0 to 10.0 feet to ensure we don't miss groups
                    // This was causing groups on different floors to be filtered out incorrectly
                    double tolerance = 10.0;
                    if (!(bb.Max.Z + tolerance < minElementZ || bb.Min.Z - tolerance > maxElementZ))
                    {
                        elevationFilteredGroups.Add(group);
                    }
                }
            }
            if (CheckCancellation())
            {
                operationCancelled = true;
                return Result.Cancelled;
            }

            progressForm.SetPhase("Mapping Elements to Groups");
            progressForm.UpdateProgress(50, 100, $"Checking {elevationFilteredGroups.Count} groups...");
            Application.DoEvents();
            
            // Add more detailed progress during mapping
            if (elevationFilteredGroups.Count > 10)
            {
                progressForm.AddMappingProgress($"Building room-to-group mapping for {elevationFilteredGroups.Count} groups");
                progressForm.AddMappingProgress("Analyzing room boundaries and spatial containment...");
                Application.DoEvents();
            }

            Dictionary<ElementId, List<Group>> elementsInGroups = MapElementsToGroups(
                selectedElements, elevationFilteredGroups, overallBB, doc, roomToSingleGroupMap);

            // Log mapping statistics
            if (enableDiagnostics)
            {
                diagnosticLog.AppendLine($"\n=== ELEMENT TO GROUP MAPPING ===");
                diagnosticLog.AppendLine($"Total groups in document: {allGroups.Count}");
                diagnosticLog.AppendLine($"Spatially relevant groups: {spatiallyRelevantGroups.Count}");
                diagnosticLog.AppendLine($"Elevation filtered groups: {elevationFilteredGroups.Count}");
                diagnosticLog.AppendLine($"Elements mapped to groups: {elementsInGroups.Count} of {selectedElements.Count}");
            }

            if (elementsInGroups.Count == 0)
            {
                message = "No selected elements are contained within rooms of any groups. Check the diagnostic log for details.";
                progressForm.AddDetail("No elements found in group rooms", CopyElementsProgressForm.DetailType.Error);
                DiagnoseElementsNotInGroups(selectedElements, elementsInGroups, doc);
                SaveDiagnostics(true);
                progressForm.SetComplete(0, globalStopwatch.ElapsedMilliseconds, false);
                return Result.Failed;
            }

            if (CheckCancellation())
            {
                operationCancelled = true;
                return Result.Cancelled;
            }

            progressForm.SetPhase("Processing Copy Operations");
            // Build element to group types mapping for Comments parameter
            Dictionary<ElementId, List<string>> elementToGroupTypeNames = BuildElementToGroupTypeMapping(
                elementsInGroups, doc);

            // Group the containing groups by their group type
            Dictionary<ElementId, List<Group>> groupsByType = GroupContainingGroupsByType(elementsInGroups);

            progressForm.AddDetail($"Processing {groupsByType.Count} group types", CopyElementsProgressForm.DetailType.Info);
            progressForm.UpdateProgress(60, 100, "Starting copy operations...");
            Application.DoEvents();

            // Process copying
            var copyResult = ProcessCopying(doc, groupsByType, elementsInGroups, elementToGroupTypeNames);
            
            if (progressForm.IsCancelled)
            {
                operationCancelled = true;
            }

            LogFinalSummary(selectedElements.Count, elementsInGroups.Count, copyResult.TotalCopied);

            // Build and show result message
            /*
            ShowResults(doc, selectedElements, elementsInGroups, groupsByType,
                        elementToGroupTypeNames, spatiallyRelevantGroups, allGroups,
                        copyResult);
            */

            SaveDiagnostics();
            
            // Complete the progress form
            progressForm.UpdateElementCounts(selectedElements.Count, elementsInGroups.Count, copyResult.TotalCopied);
            progressForm.SetComplete(copyResult.TotalCopied, globalStopwatch.ElapsedMilliseconds, operationCancelled);

            return operationCancelled ? Result.Cancelled : Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Error: {ex.Message}";
            if (progressForm != null && !progressForm.IsDisposed)
            {
                progressForm.AddDetail($"Error: {ex.Message}", CopyElementsProgressForm.DetailType.Error);
                progressForm.SetComplete(0, globalStopwatch?.ElapsedMilliseconds ?? 0, true);
            }
            SaveDiagnostics(true);
            return Result.Failed;
        }
        finally
        {
            // Don't dispose the form immediately if not cancelled - let user review results
            if (operationCancelled && progressForm != null && !progressForm.IsDisposed)
            {
                progressForm.Dispose();
            }
        }
    }
    
    private bool CheckCancellation()
    {
        if (progressForm != null && progressForm.IsCancelled)
        {
            return true;
        }
        
        // Allow UI to update
        Application.DoEvents();
        Thread.Yield();
        
        return false;
    }

    private List<Element> GetAndValidateSelection(UIDocument uidoc, Document doc, ref string message)
    {
        List<Element> selectedElements = new List<Element>();
        List<Element> linkedElementsToCopy = new List<Element>();
        Dictionary<Element, Transform> linkedElementTransforms = new Dictionary<Element, Transform>();
        
        // Try GetReferences first for linked elements support
        IList<Reference> selectedRefs = new List<Reference>();
        bool hasReferences = false;
        
        try
        {
            selectedRefs = uidoc.GetReferences();
            hasReferences = selectedRefs != null && selectedRefs.Count > 0;
        }
        catch { }
        
        // If no references from GetReferences, try GetSelectionIds
        if (!hasReferences)
        {
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            if (selectedIds != null && selectedIds.Count > 0)
            {
                selectedRefs = new List<Reference>();
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem != null)
                    {
                        selectedRefs.Add(new Reference(elem));
                    }
                }
            }
        }

        
        // If still no selection, return error
        if (selectedRefs.Count == 0)
        {
            message = "Please select elements first";
            return null;
        }

        // Process references

        foreach (Reference reference in selectedRefs)
        {
            if (reference == null) continue;
            
            // Check if it's a linked element reference
            bool isLinkedElement = false;
            try {
                isLinkedElement = reference.LinkedElementId != ElementId.InvalidElementId;
            } catch { }
            
            if (reference.LinkedElementId != ElementId.InvalidElementId)
            {
                // This is a linked element reference
                RevitLinkInstance linkInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                if (linkInstance != null)
                {
                    Document linkedDoc = linkInstance.GetLinkDocument();
                    if (linkedDoc != null)
                    {
                        Element linkedElement = linkedDoc.GetElement(reference.LinkedElementId);
                        if (linkedElement != null && !(linkedElement is Group) && !(linkedElement is ElementType))
                        {
                            linkedElementsToCopy.Add(linkedElement);
                            linkedElementTransforms[linkedElement] = linkInstance.GetTotalTransform();
                        }
                    }
                }
            }
            else
            {
                // Regular element reference
                Element elem = doc.GetElement(reference.ElementId);
                if (elem != null && !(elem is Group) && !(elem is ElementType))
                {
                    selectedElements.Add(elem);
                }
            }
        }

        // Copy linked elements to current document if any
        if (linkedElementsToCopy.Count > 0)
        {
            using (Transaction tx = new Transaction(doc, "Copy Linked Elements"))
            {
                tx.Start();
                
                foreach (Element linkedElem in linkedElementsToCopy)
                {
                    Transform transform = linkedElementTransforms[linkedElem];
                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                        linkedElem.Document,
                        new List<ElementId> { linkedElem.Id },
                        doc,
                        transform,
                        new CopyPasteOptions());
                    
                    // Add copied elements to selected elements list
                    foreach (ElementId copiedId in copiedIds)
                    {
                        Element copiedElem = doc.GetElement(copiedId);
                        if (copiedElem != null)
                        {
                            selectedElements.Add(copiedElem);

                            // Mark as copied from linked model in Comments
                            Parameter commentsParam = copiedElem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                            if (commentsParam != null && !commentsParam.IsReadOnly)
                            {
                                string existingComment = commentsParam.AsString() ?? "";
                                string newComment = string.IsNullOrEmpty(existingComment) 
                                    ? "Copied from linked model" 
                                    : existingComment + ", Copied from linked model";
                                commentsParam.Set(newComment);
                            }
                        }
                    }
                }
                
                tx.Commit();
            }
        }

        if (selectedElements.Count == 0)
        {
            message = "No valid elements found in selection. Please select elements to copy.";
            return null;
        }

        return selectedElements;
    }

    private Dictionary<ElementId, List<Group>> MapElementsToGroups(
        List<Element> selectedElements,
        List<Group> spatiallyRelevantGroups,
        BoundingBoxXYZ overallBB,
        Document doc,
        Dictionary<ElementId, Group> roomToSingleGroupMap)
    {
        // Dictionary to store which groups contain which elements
        Dictionary<ElementId, List<Group>> elementToGroupsMap = new Dictionary<ElementId, List<Group>>();

        // Initialize the map
        foreach (Element elem in selectedElements)
        {
            elementToGroupsMap[elem.Id] = new List<Group>();
        }

        // Process only spatially relevant groups
        foreach (Group group in spatiallyRelevantGroups)
        {
            // Get cached bounding box
            BoundingBoxXYZ groupBB = _groupBoundingBoxCache[group.Id];

            if (groupBB == null)
            {
                continue;
            }

            if (!BoundingBoxesIntersect(overallBB, groupBB))
            {
                continue;
            }

            // Check which selected elements are contained in this group's rooms
            List<Element> containedElements = GetElementsContainedInGroupRoomsFiltered(
                group, selectedElements, doc, roomToSingleGroupMap);

            if (containedElements.Count > 0)
            {
                // Update the map
                foreach (Element elem in containedElements)
                {
                    elementToGroupsMap[elem.Id].Add(group);
                }
            }
        }

        // Sort groups consistently by name when there are multiple groups per element
        foreach (var kvp in elementToGroupsMap)
        {
            kvp.Value.Sort((g1, g2) =>
            {
                GroupType gt1 = doc.GetElement(g1.GetTypeId()) as GroupType;
                GroupType gt2 = doc.GetElement(g2.GetTypeId()) as GroupType;
                string name1 = gt1?.Name ?? "Unknown";
                string name2 = gt2?.Name ?? "Unknown";
                return string.Compare(name1, name2, StringComparison.Ordinal);
            });
        }

        // Filter out elements that aren't in any groups
        return elementToGroupsMap
            .Where(kvp => kvp.Value.Count > 0)
            .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
    }

    private Dictionary<ElementId, List<string>> BuildElementToGroupTypeMapping(
        Dictionary<ElementId, List<Group>> elementsInGroups,
        Document doc)
    {
        Dictionary<ElementId, List<string>> elementToGroupTypeNames = new Dictionary<ElementId, List<string>>();
        
        foreach (var kvp in elementsInGroups)
        {
            ElementId elementId = kvp.Key;
            List<Group> containingGroups = kvp.Value;

            List<string> groupTypeNames = new List<string>();
            HashSet<string> uniqueNames = new HashSet<string>();

            foreach (Group group in containingGroups)
            {
                GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
                if (gt != null && uniqueNames.Add(gt.Name))
                {
                    groupTypeNames.Add(gt.Name);
                }
            }

            elementToGroupTypeNames[elementId] = groupTypeNames;
        }

        return elementToGroupTypeNames;
    }

    private Dictionary<ElementId, List<Group>> GroupContainingGroupsByType(
        Dictionary<ElementId, List<Group>> elementsInGroups)
    {
        Dictionary<ElementId, List<Group>> groupsByType = new Dictionary<ElementId, List<Group>>();
        HashSet<Group> allContainingGroups = new HashSet<Group>();

        foreach (var kvp in elementsInGroups)
        {
            foreach (Group group in kvp.Value)
            {
                allContainingGroups.Add(group);
                ElementId typeId = group.GetTypeId();
                if (!groupsByType.ContainsKey(typeId))
                {
                    groupsByType[typeId] = new List<Group>();
                }
                if (!groupsByType[typeId].Contains(group))
                {
                    groupsByType[typeId].Add(group);
                }
            }
        }

        return groupsByType;
    }

    // Helper class for copy results
    private class CopyResult
    {
        public int TotalCopied { get; set; }
        public int GroupTypesProcessed { get; set; }
        public int GroupTypesSkippedNoRefElements { get; set; }
        public int TotalGroupInstancesProcessed { get; set; }
        public int TotalGroupInstancesSkipped { get; set; }
        public int TotalCopyOperations { get; set; }
    }
}
