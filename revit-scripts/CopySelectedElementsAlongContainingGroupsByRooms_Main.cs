using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;

[Transaction(TransactionMode.Manual)]
public partial class CopySelectedElementsAlongContainingGroupsByRooms : IExternalCommand
{
    // Track room validation statistics
    private int totalRoomsChecked = 0;
    private int roomsValidatedBySimilarity = 0;
    private int roomsInvalidatedByDissimilarity = 0;
    private int roomsAsDirectMembers = 0;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            InitializeDiagnostics();

            // Get and validate selected elements
            List<Element> selectedElements = GetAndValidateSelection(uidoc, doc, ref message);
            if (selectedElements == null || selectedElements.Count == 0)
                return Result.Failed;

            LogSelectedElements(selectedElements);

            // Pre-calculate test points for all selected elements
            foreach (Element elem in selectedElements)
            {
                _elementTestPointsCache[elem.Id] = GetElementTestPoints(elem);
            }

            // Create bounding box for all selected elements for quick filtering
            BoundingBoxXYZ overallBB = GetOverallBoundingBox(selectedElements);
            if (overallBB == null)
            {
                message = "Could not determine bounding box of selected elements";
                return Result.Failed;
            }

            // Find all groups in the document
            FilteredElementCollector groupCollector = new FilteredElementCollector(doc);
            IList<Group> allGroups = groupCollector
                .OfClass(typeof(Group))
                .WhereElementIsNotElementType()
                .Cast<Group>()
                .ToList();

            // Build room cache for ALL rooms in the document
            BuildRoomCache(doc);

            // OPTIMIZATION: Clear caches from previous runs
            _floorToFloorHeightCache.Clear();
            _roomPointContainmentCache.Clear();
            _roomBoundaryCache.Clear();

            // PRE-CALCULATE ALL GROUP BOUNDING BOXES AND BUILD SPATIAL INDEX
            PreCalculateGroupDataAndSpatialIndex(allGroups, doc);

            // Get potentially relevant groups using spatial index
            List<Group> spatiallyRelevantGroups = GetSpatiallyRelevantGroups(overallBB);

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
                DiagnoseElementsNotInGroups(selectedElements, elementsInGroups, doc);
                SaveDiagnostics(true);
                return Result.Failed;
            }

            // Build element to group types mapping for Comments parameter
            Dictionary<ElementId, List<string>> elementToGroupTypeNames = BuildElementToGroupTypeMapping(
                elementsInGroups, doc);

            // Group the containing groups by their group type
            Dictionary<ElementId, List<Group>> groupsByType = GroupContainingGroupsByType(elementsInGroups);

            // Process copying
            var copyResult = ProcessCopying(doc, groupsByType, elementsInGroups, elementToGroupTypeNames);

            LogFinalSummary(selectedElements.Count, elementsInGroups.Count, copyResult.TotalCopied);

            // Build and show result message
            ShowResults(doc, selectedElements, elementsInGroups, groupsByType,
                        elementToGroupTypeNames, spatiallyRelevantGroups, allGroups,
                        copyResult);

            SaveDiagnostics();

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Error: {ex.Message}";
            SaveDiagnostics(true);
            return Result.Failed;
        }
    }

    private List<Element> GetAndValidateSelection(UIDocument uidoc, Document doc, ref string message)
    {
        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();

        if (selectedIds.Count == 0)
        {
            message = "Please select elements first";
            return null;
        }

        // Get selected elements (exclude groups and element types)
        List<Element> selectedElements = new List<Element>();

        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem != null && !(elem is Group) && !(elem is ElementType))
            {
                selectedElements.Add(elem);
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
