using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// Static class to manage selection mode using SelectionSets
public static class SelectionModeManager
{
    private static readonly string AppDataPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "revit-scripts"
    );
    
    private static readonly string ModeFilePath = Path.Combine(AppDataPath, "SelectionMode");
    private static readonly string LinkedReferencesFilePath = Path.Combine(AppDataPath, "SelectionSet-LinkedModelReferences");
    private const string TempSelectionSetName = "temp";
    
    public enum SelectionMode
    {
        RevitUI,
        SelectionSet
    }
    
    // Structure to store linked reference information
    private class LinkedReferenceInfo
    {
        public int LinkInstanceId { get; set; }
        public int LinkedElementId { get; set; }
    }
    
    static SelectionModeManager()
    {
        // Ensure directory exists
        Directory.CreateDirectory(AppDataPath);
    }
    
    public static SelectionMode CurrentMode
    {
        get
        {
            if (File.Exists(ModeFilePath))
            {
                string mode = File.ReadAllText(ModeFilePath).Trim();
                if (mode == "SelectionSet")
                    return SelectionMode.SelectionSet;
            }
            return SelectionMode.RevitUI;
        }
        set
        {
            File.WriteAllText(ModeFilePath, value.ToString());
        }
    }
    
    // Get or create the temp selection set
    private static SelectionFilterElement GetOrCreateTempSelectionSet(Document doc)
    {
        // Try to find existing selection set
        var collector = new FilteredElementCollector(doc)
            .OfClass(typeof(SelectionFilterElement));
        
        SelectionFilterElement tempSet = collector
            .Cast<SelectionFilterElement>()
            .FirstOrDefault(s => s.Name == TempSelectionSetName);
        
        if (tempSet == null)
        {
            // Create new selection set
            if (doc.IsModifiable)
            {
                // We're already in a transaction, just create directly
                tempSet = SelectionFilterElement.Create(doc, TempSelectionSetName);
            }
            else
            {
                // No active transaction, create one
                using (Transaction tx = new Transaction(doc, "Create Temp Selection Set"))
                {
                    tx.Start();
                    tempSet = SelectionFilterElement.Create(doc, TempSelectionSetName);
                    tx.Commit();
                }
            }
        }
        
        return tempSet;
    }
    
    // Save linked references to file
    private static void SaveLinkedReferences(List<LinkedReferenceInfo> linkedRefs)
    {
        if (linkedRefs == null || !linkedRefs.Any())
        {
            // Delete file if no linked references
            if (File.Exists(LinkedReferencesFilePath))
                File.Delete(LinkedReferencesFilePath);
            return;
        }
        
        var lines = linkedRefs.Select(r => $"{r.LinkInstanceId},{r.LinkedElementId}");
        File.WriteAllLines(LinkedReferencesFilePath, lines);
    }
    
    // Load linked references from file
    private static List<LinkedReferenceInfo> LoadLinkedReferences()
    {
        var result = new List<LinkedReferenceInfo>();
        
        if (!File.Exists(LinkedReferencesFilePath))
            return result;
        
        try
        {
            var lines = File.ReadAllLines(LinkedReferencesFilePath);
            foreach (var line in lines)
            {
                var parts = line.Split(',');
                if (parts.Length == 2 && 
                    int.TryParse(parts[0], out int linkId) && 
                    int.TryParse(parts[1], out int elemId))
                {
                    result.Add(new LinkedReferenceInfo 
                    { 
                        LinkInstanceId = linkId, 
                        LinkedElementId = elemId 
                    });
                }
            }
        }
        catch { /* Return empty list on error */ }
        
        return result;
    }
    
    // Extension methods for UIDocument to provide selection functionality
    public static ICollection<ElementId> GetSelectionIds(this UIDocument uidoc)
    {
        if (CurrentMode == SelectionMode.SelectionSet)
        {
            var doc = uidoc.Document;
            var tempSet = GetOrCreateTempSelectionSet(doc);
            return tempSet.GetElementIds();
        }
        return uidoc.Selection.GetElementIds();
    }
    
    public static void SetSelectionIds(this UIDocument uidoc, ICollection<ElementId> elementIds)
    {
        if (CurrentMode == SelectionMode.SelectionSet)
        {
            var doc = uidoc.Document;
            var tempSet = GetOrCreateTempSelectionSet(doc);
            
            // Check if we're already in a transaction
            if (doc.IsModifiable)
            {
                // We're already in a transaction, just modify directly
                tempSet.SetElementIds(elementIds);
            }
            else
            {
                // No active transaction, create one
                using (Transaction tx = new Transaction(doc, "Update Temp Selection Set"))
                {
                    tx.Start();
                    tempSet.SetElementIds(elementIds);
                    tx.Commit();
                }
            }
            
            // Clear linked references when setting only element IDs
            SaveLinkedReferences(null);
        }
        else
        {
            uidoc.Selection.SetElementIds(elementIds);
        }
    }
    
    // Get current references (including linked)
    public static IList<Reference> GetReferences(this UIDocument uidoc)
    {
        if (CurrentMode == SelectionMode.SelectionSet)
        {
            var doc = uidoc.Document;
            var references = new List<Reference>();
            
            // Load linked references from file
            var linkedRefs = LoadLinkedReferences();
            foreach (var linkedRef in linkedRefs)
            {
                try
                {
                    var linkInstance = doc.GetElement(new ElementId(linkedRef.LinkInstanceId)) as RevitLinkInstance;
                    if (linkInstance != null)
                    {
                        var linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            var linkedElement = linkedDoc.GetElement(new ElementId(linkedRef.LinkedElementId));
                            if (linkedElement != null)
                            {
                                var elemRef = new Reference(linkedElement);
                                var linkRef = elemRef.CreateLinkReference(linkInstance);
                                if (linkRef != null)
                                    references.Add(linkRef);
                            }
                        }
                    }
                }
                catch { /* Skip problematic references */ }
            }
            
            return references;
        }
        return uidoc.Selection.GetReferences();
    }
    
    public static IList<Reference> GetSelectionReferences(this UIDocument uidoc)
    {
        if (CurrentMode == SelectionMode.SelectionSet)
        {
            // For picking new objects, we can't use SelectionSet mode
            // This would require switching to UI mode temporarily
            throw new InvalidOperationException("Cannot pick new objects in SelectionSet mode. Please switch to RevitUI mode.");
        }
        return uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element);
    }
    
    public static void SetReferences(this UIDocument uidoc, IList<Reference> references)
    {
        if (CurrentMode == SelectionMode.SelectionSet)
        {
            var doc = uidoc.Document;
            var regularElementIds = new List<ElementId>();
            var linkedRefs = new List<LinkedReferenceInfo>();
            
            // Separate regular and linked references
            foreach (var reference in references)
            {
                if (reference.LinkedElementId != ElementId.InvalidElementId)
                {
                    // This is a linked element reference
                    linkedRefs.Add(new LinkedReferenceInfo
                    {
                        LinkInstanceId = reference.ElementId.IntegerValue,
                        LinkedElementId = reference.LinkedElementId.IntegerValue
                    });
                }
                else
                {
                    // Regular element reference
                    regularElementIds.Add(reference.ElementId);
                }
            }
            
            // Store regular elements in selection set
            var tempSet = GetOrCreateTempSelectionSet(doc);
            
            // Check if we're already in a transaction
            if (doc.IsModifiable)
            {
                // We're already in a transaction, just modify directly
                tempSet.SetElementIds(regularElementIds);
            }
            else
            {
                // No active transaction, create one
                using (Transaction tx = new Transaction(doc, "Update Temp Selection Set"))
                {
                    tx.Start();
                    tempSet.SetElementIds(regularElementIds);
                    tx.Commit();
                }
            }
            
            // Store linked references in file
            SaveLinkedReferences(linkedRefs);
        }
        else
        {
            uidoc.Selection.SetReferences(references);
        }
    }
    
    // Add to existing selection
    public static void AddToSelection(this UIDocument uidoc, ICollection<ElementId> elementIds)
    {
        var currentSelection = uidoc.GetSelectionIds();
        var combined = new HashSet<ElementId>(currentSelection);
        foreach (var id in elementIds)
        {
            combined.Add(id);
        }
        uidoc.SetSelectionIds(combined.ToList());
    }
    
    // Clear the temp selection set and linked references
    public static void ClearTempSelectionSet(Document doc)
    {
        var tempSet = GetOrCreateTempSelectionSet(doc);
        
        // Check if we're already in a transaction
        if (doc.IsModifiable)
        {
            // We're already in a transaction, just modify directly
            tempSet.SetElementIds(new List<ElementId>());
        }
        else
        {
            // No active transaction, create one
            using (Transaction tx = new Transaction(doc, "Clear Temp Selection Set"))
            {
                tx.Start();
                tempSet.SetElementIds(new List<ElementId>());
                tx.Commit();
            }
        }
        
        // Clear linked references file
        if (File.Exists(LinkedReferencesFilePath))
            File.Delete(LinkedReferencesFilePath);
    }
}

// Simple toggle command - switches between modes quickly
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ToggleSelectionMode : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var currentMode = SelectionModeManager.CurrentMode;
        var newMode = currentMode == SelectionModeManager.SelectionMode.RevitUI 
            ? SelectionModeManager.SelectionMode.SelectionSet 
            : SelectionModeManager.SelectionMode.RevitUI;
        
        SelectionModeManager.CurrentMode = newMode;
        
        // Clear temp selection set when switching to SelectionSet mode
        if (newMode == SelectionModeManager.SelectionMode.SelectionSet)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;
            SelectionModeManager.ClearTempSelectionSet(doc);
        }
        
        return Result.Succeeded;
    }
}

// DataGrid command - shows current mode and allows selection
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SwitchSelectionMode : IExternalCommand
{
    // Wrapper class for displaying selection modes
    public class SelectionModeWrapper
    {
        public string Mode { get; set; }
        public SelectionModeManager.SelectionMode EnumValue { get; set; }
    }
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var currentMode = SelectionModeManager.CurrentMode;
        
        // Create list of available modes
        var modes = new List<SelectionModeWrapper>
        {
            new SelectionModeWrapper { Mode = "RevitUI", EnumValue = SelectionModeManager.SelectionMode.RevitUI },
            new SelectionModeWrapper { Mode = "SelectionSet", EnumValue = SelectionModeManager.SelectionMode.SelectionSet }
        };
        
        // Define property names for DataGrid
        var propertyNames = new List<string> { "Mode" };
        
        // Determine initial selection index
        int selectedIndex = modes.FindIndex(m => m.EnumValue == currentMode);
        var initialSelectionIndices = new List<int> { selectedIndex };
        
        // Display modes using DataGrid
        var selectedModes = CustomGUIs.DataGrid(modes, propertyNames, initialSelectionIndices);
        
        if (selectedModes == null || selectedModes.Count == 0)
        {
            // User cancelled - keep current mode
            return Result.Cancelled;
        }
        
        var chosenMode = selectedModes.First().EnumValue;
        
        // Only change if different from current
        if (chosenMode != currentMode)
        {
            SelectionModeManager.CurrentMode = chosenMode;
            
            // Clear temp selection set when switching to SelectionSet mode
            if (chosenMode == SelectionModeManager.SelectionMode.SelectionSet)
            {
                var doc = commandData.Application.ActiveUIDocument.Document;
                SelectionModeManager.ClearTempSelectionSet(doc);
            }
        }
        
        return Result.Succeeded;
    }
}
