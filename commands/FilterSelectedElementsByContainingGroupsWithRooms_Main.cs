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
public partial class FilterSelectedElementsByContainingGroupsWithRooms : IExternalCommand
{
    // Diagnostic data collection
    private StringBuilder _diagnostics = new StringBuilder();
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        try
        {
            // Get selected elements (handles both regular and linked references)
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            IList<Reference> selectedRefs = uidoc.GetReferences();
            
            // Start diagnostics
            _diagnostics.AppendLine("=== DIAGNOSTIC REPORT ===");
            _diagnostics.AppendLine($"Total selected IDs: {selectedIds.Count}");
            _diagnostics.AppendLine($"Total references (including linked): {selectedRefs?.Count ?? 0}");
            _diagnostics.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _diagnostics.AppendLine();
            
            // Get selected elements - include everything except Groups themselves and ElementTypes
            List<Element> selectedElements = new List<Element>();
            Dictionary<Element, Document> elementToDocMap = new Dictionary<Element, Document>();
            Dictionary<Element, RevitLinkInstance> elementToLinkMap = new Dictionary<Element, RevitLinkInstance>();
            
            _diagnostics.AppendLine("=== ELEMENT FILTERING ===");
            
            // First process regular selected elements from the host model
            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem != null && !(elem is RevitLinkInstance))
                {
                    string elemInfo = $"ID {id.IntegerValue}: {elem.GetType().Name}";
                    
                    if (elem is Group)
                    {
                        _diagnostics.AppendLine($"  {elemInfo} - EXCLUDED (is a Group instance)");
                    }
                    else if (elem is ElementType)
                    {
                        _diagnostics.AppendLine($"  {elemInfo} - EXCLUDED (is an ElementType)");
                    }
                    else if (elem is GroupType)
                    {
                        _diagnostics.AppendLine($"  {elemInfo} - EXCLUDED (is a GroupType)");
                    }
                    else
                    {
                        selectedElements.Add(elem);
                        elementToDocMap[elem] = doc;
                        _diagnostics.AppendLine($"  {elemInfo} - INCLUDED (Category: {elem.Category?.Name ?? "None"}) [HOST MODEL]");
                    }
                }
                else if (elem is RevitLinkInstance)
                {
                    _diagnostics.AppendLine($"  ID {id.IntegerValue}: RevitLinkInstance - checking for linked elements");
                }
                else if (elem == null)
                {
                    _diagnostics.AppendLine($"  ID {id.IntegerValue}: NULL element - EXCLUDED");
                }
            }
            
            // Now process linked references if any
            if (selectedRefs != null && selectedRefs.Count > 0)
            {
                _diagnostics.AppendLine("\n=== PROCESSING LINKED REFERENCES ===");
                
                foreach (Reference reference in selectedRefs)
                {
                    if (reference.LinkedElementId != ElementId.InvalidElementId)
                    {
                        // This is a linked element reference
                        RevitLinkInstance linkInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                        if (linkInstance != null)
                        {
                            Document linkedDoc = linkInstance.GetLinkDocument();
                            if (linkedDoc != null)
                            {
                                Element linkedElem = linkedDoc.GetElement(reference.LinkedElementId);
                                if (linkedElem != null)
                                {
                                    string elemInfo = $"Linked ID {reference.LinkedElementId.IntegerValue}: {linkedElem.GetType().Name}";
                                    
                                    if (linkedElem is Group)
                                    {
                                        _diagnostics.AppendLine($"  {elemInfo} - EXCLUDED (is a Group instance)");
                                    }
                                    else if (linkedElem is ElementType)
                                    {
                                        _diagnostics.AppendLine($"  {elemInfo} - EXCLUDED (is an ElementType)");
                                    }
                                    else if (linkedElem is GroupType)
                                    {
                                        _diagnostics.AppendLine($"  {elemInfo} - EXCLUDED (is a GroupType)");
                                    }
                                    else
                                    {
                                        selectedElements.Add(linkedElem);
                                        elementToDocMap[linkedElem] = linkedDoc;
                                        elementToLinkMap[linkedElem] = linkInstance;
                                        string linkName = linkInstance.Name ?? "Unknown Link";
                                        _diagnostics.AppendLine($"  {elemInfo} - INCLUDED (Category: {linkedElem.Category?.Name ?? "None"}) [LINKED: {linkName}]");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            _diagnostics.AppendLine($"\nFiltered elements count: {selectedElements.Count}");
            _diagnostics.AppendLine($"  - From host model: {selectedElements.Count(e => elementToDocMap[e] == doc)}");
            _diagnostics.AppendLine($"  - From linked models: {selectedElements.Count(e => elementToDocMap[e] != doc)}");
            
            if (selectedElements.Count == 0)
            {
                _diagnostics.AppendLine("\nNO VALID ELEMENTS AFTER FILTERING!");
                _diagnostics.AppendLine("Make sure you're selecting actual elements, not group instances.");
                
                // Show diagnostics before failing
                ShowDiagnostics();
                
                message = "No valid elements found in selection. Please select elements (not group instances).\n\nCheck the diagnostic report for details.";
                return Result.Failed;
            }
            
            // Process elements by their source document
            Dictionary<Document, List<Element>> elementsByDoc = new Dictionary<Document, List<Element>>();
            foreach (Element elem in selectedElements)
            {
                Document elemDoc = elementToDocMap[elem];
                if (!elementsByDoc.ContainsKey(elemDoc))
                {
                    elementsByDoc[elemDoc] = new List<Element>();
                }
                elementsByDoc[elemDoc].Add(elem);
            }
            
            // Build room cache and spatial index for each document
            _diagnostics.AppendLine("\n=== BUILDING SPATIAL INDICES ===");
            Dictionary<Document, SpatialIndex> roomIndicesByDoc = new Dictionary<Document, SpatialIndex>();
            Dictionary<Document, List<Room>> roomsByDoc = new Dictionary<Document, List<Room>>();
            
            foreach (Document docToProcess in elementsByDoc.Keys)
            {
                string docName = (docToProcess == doc) ? "HOST MODEL" : "LINKED MODEL";
                _diagnostics.AppendLine($"\nProcessing {docName}:");
                
                SpatialIndex roomIndex = new SpatialIndex();
                List<Room> rooms = BuildRoomCacheAndIndex(docToProcess, roomIndex);
                roomIndicesByDoc[docToProcess] = roomIndex;
                roomsByDoc[docToProcess] = rooms;
                _diagnostics.AppendLine($"  Indexed {rooms.Count} rooms");
            }
            
            // Pre-calculate element test points
            _diagnostics.AppendLine("\n=== PRE-CALCULATING ELEMENT DATA ===");
            Dictionary<ElementId, ElementData> elementDataCache = new Dictionary<ElementId, ElementData>();
            foreach (Element elem in selectedElements)
            {
                Document elemDoc = elementToDocMap[elem];
                elementDataCache[elem.Id] = new ElementData
                {
                    TestPoints = GetElementTestPoints(elem),
                    BoundingBox = elem.get_BoundingBox(null),
                    Level = elemDoc.GetElement(elem.LevelId) as Level
                };
            }
            
            // Find which rooms contain the selected elements using spatial index
            _diagnostics.AppendLine("\n=== ROOM CONTAINMENT CHECK (WITH SPATIAL INDEX) ===");
            Dictionary<ElementId, List<Room>> elementToRoomsMap = new Dictionary<ElementId, List<Room>>();
            
            foreach (var kvp in elementsByDoc)
            {
                Document elemDoc = kvp.Key;
                List<Element> docElements = kvp.Value;
                SpatialIndex roomIndex = roomIndicesByDoc[elemDoc];
                
                string docName = (elemDoc == doc) ? "HOST MODEL" : "LINKED MODEL";
                _diagnostics.AppendLine($"\nChecking elements in {docName}:");
                
                // Find rooms for elements in this document
                var roomsForElements = FindRoomsContainingElementsOptimized(docElements, elementDataCache, roomIndex);
                
                // Merge results
                foreach (var elemKvp in roomsForElements)
                {
                    elementToRoomsMap[elemKvp.Key] = elemKvp.Value;
                }
            }
            
            // Report findings
            int elementsInRooms = elementToRoomsMap.Count(kvp => kvp.Value.Count > 0);
            _diagnostics.AppendLine($"Elements in at least one room: {elementsInRooms}/{selectedElements.Count}");
            
            if (elementsInRooms == 0)
            {
                _diagnostics.AppendLine("\nNO ELEMENTS ARE IN ANY ROOMS!");
                ShowDiagnostics();
                
                TaskDialog.Show("No Rooms Found", 
                    "None of the selected elements are contained in any rooms.\n\n" +
                    "Check the diagnostic report for details.");
                return Result.Succeeded;
            }
            
            // Now find groups - for linked elements, look for groups in their linked document
            _diagnostics.AppendLine($"\n=== GROUP ANALYSIS ===");
            Dictionary<ElementId, List<Group>> elementToGroupsMap = new Dictionary<ElementId, List<Group>>();
            
            foreach (var kvp in elementsByDoc)
            {
                Document elemDoc = kvp.Key;
                List<Element> docElements = kvp.Value;
                
                string docName = (elemDoc == doc) ? "HOST MODEL" : "LINKED MODEL";
                _diagnostics.AppendLine($"\nAnalyzing groups in {docName}:");
                
                // Get all groups in this document
                FilteredElementCollector groupCollector = new FilteredElementCollector(elemDoc);
                IList<Group> allGroups = groupCollector
                    .OfClass(typeof(Group))
                    .WhereElementIsNotElementType()
                    .Cast<Group>()
                    .ToList();
                
                _diagnostics.AppendLine($"  Total groups in document: {allGroups.Count}");
                
                // Pre-calculate group bounding boxes for this document
                PreCalculateGroupBoundingBoxes(allGroups, elemDoc);
                
                // Build comprehensive room-to-group mapping with priority handling (using boundary analysis)
                Dictionary<ElementId, Group> roomToSingleGroupMap = BuildRoomToGroupMapping(allGroups, elemDoc);
                
                // Find groups for elements in this document
                foreach (Element elem in docElements)
                {
                    if (elementToRoomsMap.ContainsKey(elem.Id))
                    {
                        List<Room> rooms = elementToRoomsMap[elem.Id];
                        HashSet<Group> uniqueGroups = new HashSet<Group>();
                        
                        foreach (Room room in rooms)
                        {
                            if (roomToSingleGroupMap.ContainsKey(room.Id))
                            {
                                uniqueGroups.Add(roomToSingleGroupMap[room.Id]);
                            }
                        }
                        
                        elementToGroupsMap[elem.Id] = uniqueGroups.ToList();
                    }
                    else
                    {
                        elementToGroupsMap[elem.Id] = new List<Group>();
                    }
                }
            }
            
            // Report on elements with multiple groups (edge cases)
            var multiGroupElements = elementToGroupsMap.Where(kvp => kvp.Value.Count > 1).ToList();
            if (multiGroupElements.Any())
            {
                _diagnostics.AppendLine($"\n=== ELEMENTS IN MULTIPLE GROUPS (EDGE CASES) ===");
                foreach (var kvp in multiGroupElements)
                {
                    Element elem = doc.GetElement(kvp.Key);
                    _diagnostics.AppendLine($"Element {elem.Id.IntegerValue} is in {kvp.Value.Count} groups:");
                    foreach (Group group in kvp.Value)
                    {
                        GroupType gt = doc.GetElement(group.GetTypeId()) as GroupType;
                        _diagnostics.AppendLine($"  - {gt?.Name ?? "Unknown"}");
                    }
                }
            }
            
            // Find maximum number of groups any single element is contained in
            int maxGroupsPerElement = 0;
            foreach (var groups in elementToGroupsMap.Values)
            {
                maxGroupsPerElement = Math.Max(maxGroupsPerElement, groups.Count);
            }
            
            _diagnostics.AppendLine($"\n=== SUMMARY ===");
            _diagnostics.AppendLine($"Elements in rooms: {elementsInRooms}");
            _diagnostics.AppendLine($"Elements in groups (via rooms): {elementToGroupsMap.Count(kvp => kvp.Value.Count > 0)}");
            _diagnostics.AppendLine($"Maximum groups per element: {maxGroupsPerElement}");
            _diagnostics.AppendLine($"Elements in multiple groups (edge cases): {multiGroupElements.Count}");
            
            // Show diagnostics
            ShowDiagnostics();
            
            if (maxGroupsPerElement == 0)
            {
                TaskDialog.Show("No Groups Found", 
                    "The selected elements are in rooms, but those rooms are not part of any groups.\n\n" +
                    "Check the diagnostic report for details.");
                return Result.Succeeded;
            }
            
            // Prepare data for DataGrid
            List<Dictionary<string, object>> gridEntries = PrepareGridEntries(
                selectedElements, elementToRoomsMap, elementToGroupsMap, 
                maxGroupsPerElement, elementToDocMap, elementToLinkMap);
            
            // Prepare property names
            List<string> propertyNames = new List<string> { "Type", "Family", "Id", "Comments", "Room(s)", "Group" };
            
            // Only add additional group columns if we have edge cases with multiple groups
            if (maxGroupsPerElement > 1)
            {
                for (int i = 2; i <= maxGroupsPerElement; i++)
                {
                    propertyNames.Add($"Alt Group {i}");
                }
            }
            
            // Show DataGrid to user
            List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(
                gridEntries,
                propertyNames,
                false,
                null
            );
            
            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                // User cancelled or selected nothing
                return Result.Cancelled;
            }
            
            // Determine which elements to select based on selected entries
            HashSet<Element> elementsToSelect = new HashSet<Element>();
            
            foreach (var entry in selectedEntries)
            {
                // Get the element from the hidden key
                if (entry.ContainsKey("_Element") && entry["_Element"] is Element)
                {
                    Element elem = entry["_Element"] as Element;
                    elementsToSelect.Add(elem);
                }
            }
            
            // Set selection to chosen elements
            if (elementsToSelect.Count > 0)
            {
                List<ElementId> elementIds = elementsToSelect.Select(e => e.Id).ToList();
                uidoc.SetSelectionIds(elementIds);
            }
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            _diagnostics.AppendLine($"\n=== EXCEPTION ===");
            _diagnostics.AppendLine($"Message: {ex.Message}");
            _diagnostics.AppendLine($"Stack Trace: {ex.StackTrace}");
            ShowDiagnostics();
            return Result.Failed;
        }
    }
    
    // Find which groups contain elements through their rooms
    private Dictionary<ElementId, List<Group>> FindGroupsForElements(
        Dictionary<ElementId, List<Room>> elementToRoomsMap,
        Dictionary<ElementId, Group> roomToGroupMap)
    {
        Dictionary<ElementId, List<Group>> result = new Dictionary<ElementId, List<Group>>();
        
        foreach (var kvp in elementToRoomsMap)
        {
            ElementId elemId = kvp.Key;
            List<Room> rooms = kvp.Value;
            
            HashSet<Group> uniqueGroups = new HashSet<Group>();
            
            foreach (Room room in rooms)
            {
                if (roomToGroupMap.ContainsKey(room.Id))
                {
                    uniqueGroups.Add(roomToGroupMap[room.Id]);
                }
            }
            
            result[elemId] = uniqueGroups.ToList();
        }
        
        return result;
    }
    
    // Prepare grid entries for display
    private List<Dictionary<string, object>> PrepareGridEntries(
        List<Element> selectedElements,
        Dictionary<ElementId, List<Room>> elementToRoomsMap,
        Dictionary<ElementId, List<Group>> elementToGroupsMap,
        int maxGroupsPerElement,
        Dictionary<Element, Document> elementToDocMap,
        Dictionary<Element, RevitLinkInstance> elementToLinkMap)
    {
        List<Dictionary<string, object>> gridEntries = new List<Dictionary<string, object>>();
        
        foreach (Element elem in selectedElements)
        {
            Dictionary<string, object> entry = new Dictionary<string, object>();
            
            // Store the actual element object for later retrieval
            entry["_Element"] = elem;
            
            // Get the document this element belongs to
            Document elemDoc = elementToDocMap[elem];
            
            // Type (Category)
            entry["Type"] = elem.Category?.Name ?? "Unknown";
            
            // Family
            ElementType elemType = elemDoc.GetElement(elem.GetTypeId()) as ElementType;
            FamilySymbol famSymbol = elemType as FamilySymbol;
            if (famSymbol != null)
            {
                entry["Family"] = famSymbol.FamilyName ?? "Unknown";
            }
            else
            {
                entry["Family"] = elemType?.Name ?? "Unknown";
            }
            
            // Id - include link info if from linked model
            string idText = elem.Id.IntegerValue.ToString();
            if (elementToLinkMap.ContainsKey(elem))
            {
                RevitLinkInstance link = elementToLinkMap[elem];
                string linkName = link.Name ?? "Link";
                idText = $"{idText} [{linkName}]";
            }
            entry["Id"] = idText;
            
            // Comments
            Parameter commentsParam = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            entry["Comments"] = commentsParam?.AsString() ?? "";
            
            // Room(s) containing this element
            List<Room> containingRooms = elementToRoomsMap.ContainsKey(elem.Id) 
                ? elementToRoomsMap[elem.Id] 
                : new List<Room>();
            
            if (containingRooms.Count > 0)
            {
                List<string> roomNames = containingRooms.Select(r => r.Name ?? $"Room {r.Id.IntegerValue}").ToList();
                entry["Room(s)"] = string.Join(", ", roomNames);
            }
            else
            {
                entry["Room(s)"] = "None";
            }
            
            // Get groups containing this element
            List<Group> containingGroups = elementToGroupsMap.ContainsKey(elem.Id)
                ? elementToGroupsMap[elem.Id]
                : new List<Group>();
            
            // Sort groups by name for consistent ordering
            containingGroups = containingGroups.OrderBy(g => 
            {
                Document groupDoc = g.Document;
                GroupType gt = groupDoc.GetElement(g.GetTypeId()) as GroupType;
                return gt?.Name ?? "Unknown";
            }).ToList();
            
            // Primary group column
            if (containingGroups.Count > 0)
            {
                Group group = containingGroups[0];
                Document groupDoc = group.Document;
                GroupType groupType = groupDoc.GetElement(group.GetTypeId()) as GroupType;
                string groupName = groupType?.Name ?? "Unknown";
                entry["Group"] = groupName;
            }
            else
            {
                entry["Group"] = "";
            }
            
            // Additional group columns only if needed (edge cases)
            if (maxGroupsPerElement > 1)
            {
                for (int i = 1; i < maxGroupsPerElement; i++)
                {
                    string columnName = $"Alt Group {i + 1}";
                    if (i < containingGroups.Count)
                    {
                        Group group = containingGroups[i];
                        Document groupDoc = group.Document;
                        GroupType groupType = groupDoc.GetElement(group.GetTypeId()) as GroupType;
                        string groupName = groupType?.Name ?? "Unknown";
                        entry[columnName] = groupName;
                    }
                    else
                    {
                        entry[columnName] = "";
                    }
                }
            }
            
            gridEntries.Add(entry);
        }
        
        return gridEntries;
    }
    
    // Show diagnostics dialog
    private void ShowDiagnostics()
    {
        TaskDialog diagnosticDialog = new TaskDialog("Diagnostics Report");
        diagnosticDialog.MainInstruction = "Group Detection Diagnostics";
        diagnosticDialog.MainContent = "The diagnostic report has been generated. Click 'Show Report' to view it.";
        diagnosticDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Show Report");
        diagnosticDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Continue");
        
        TaskDialogResult dialogResult = diagnosticDialog.Show();
        
        if (dialogResult == TaskDialogResult.CommandLink1)
        {
            // Show the diagnostics in a form that can be copied
            System.Windows.Forms.Form form = new System.Windows.Forms.Form();
            form.Text = "Diagnostic Report - Copy All Text";
            form.Size = new System.Drawing.Size(900, 700);
            
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            textBox.Multiline = true;
            textBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            textBox.ReadOnly = true;
            textBox.Font = new System.Drawing.Font("Consolas", 9);
            textBox.Text = _diagnostics.ToString();
            textBox.Dock = System.Windows.Forms.DockStyle.Fill;
            textBox.WordWrap = false;
            
            form.Controls.Add(textBox);
            form.ShowDialog();
        }
    }
}
