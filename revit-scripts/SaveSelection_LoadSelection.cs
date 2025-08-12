using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// Structure to store linked reference information for selection sets
internal class SelectionSetLinkedReferenceInfo
{
    public int LinkInstanceId { get; set; }
    public int LinkedElementId { get; set; }
}

// Static class to manage selection sets with linked element support
internal static class SelectionSetManager
{
    private static readonly string AppDataPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "revit-scripts",
        "SelectionSets"
    );
    
    static SelectionSetManager()
    {
        // Ensure directory exists - create all parent directories if needed
        if (!Directory.Exists(AppDataPath))
        {
            Directory.CreateDirectory(AppDataPath);
        }
    }
    
    // Get project name from document
    private static string GetProjectName(Document doc)
    {
        // Get project name from document path or use document title
        if (!string.IsNullOrEmpty(doc.PathName))
        {
            return Path.GetFileNameWithoutExtension(doc.PathName);
        }
        else
        {
            // For unsaved documents, use title or a default name
            return string.IsNullOrEmpty(doc.Title) ? "Untitled" : doc.Title;
        }
    }
    
    // Get the file path for linked references associated with a selection set
    private static string GetLinkedReferencesFilePath(Document doc, string selectionSetName)
    {
        // Ensure directory exists before creating file path
        if (!Directory.Exists(AppDataPath))
        {
            Directory.CreateDirectory(AppDataPath);
        }
        
        string projectName = GetProjectName(doc);
        // Sanitize file name to remove invalid characters
        string safeProjectName = SanitizeFileName(projectName);
        string safeSelectionSetName = SanitizeFileName(selectionSetName);
        
        return Path.Combine(AppDataPath, $"{safeProjectName}-{safeSelectionSetName}");
    }
    
    // Sanitize file name by removing invalid characters
    private static string SanitizeFileName(string fileName)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = fileName;
        foreach (char c in invalidChars)
        {
            sanitized = sanitized.Replace(c, '_');
        }
        return sanitized;
    }
    
    // Save linked references to file
    public static void SaveLinkedReferences(Document doc, string selectionSetName, List<SelectionSetLinkedReferenceInfo> linkedRefs)
    {
        // Ensure directory exists before saving
        if (!Directory.Exists(AppDataPath))
        {
            Directory.CreateDirectory(AppDataPath);
        }
        
        string filePath = GetLinkedReferencesFilePath(doc, selectionSetName);
        
        if (linkedRefs == null || !linkedRefs.Any())
        {
            // Delete file if no linked references
            if (File.Exists(filePath))
                File.Delete(filePath);
            return;
        }
        
        var lines = linkedRefs.Select(r => $"{r.LinkInstanceId},{r.LinkedElementId}");
        File.WriteAllLines(filePath, lines);
    }
    
    // Load linked references from file
    public static List<SelectionSetLinkedReferenceInfo> LoadLinkedReferences(Document doc, string selectionSetName)
    {
        var result = new List<SelectionSetLinkedReferenceInfo>();
        string filePath = GetLinkedReferencesFilePath(doc, selectionSetName);
        
        if (!File.Exists(filePath))
            return result;
        
        try
        {
            var lines = File.ReadAllLines(filePath);
            foreach (var line in lines)
            {
                var parts = line.Split(',');
                if (parts.Length == 2 && 
                    int.TryParse(parts[0], out int linkId) && 
                    int.TryParse(parts[1], out int elemId))
                {
                    result.Add(new SelectionSetLinkedReferenceInfo 
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
    
    // Get all existing selection set names
    public static List<string> GetSelectionSetNames(Document doc)
    {
        var collector = new FilteredElementCollector(doc)
            .OfClass(typeof(SelectionFilterElement));
        
        return collector
            .Cast<SelectionFilterElement>()
            .Select(s => s.Name)
            .Where(name => name != "temp") // Exclude temp selection set
            .OrderBy(name => name)
            .ToList();
    }
    
    // Delete linked references file when selection set is deleted
    public static void DeleteLinkedReferences(Document doc, string selectionSetName)
    {
        string filePath = GetLinkedReferencesFilePath(doc, selectionSetName);
        if (File.Exists(filePath))
            File.Delete(filePath);
    }
}

// Simple WinForm for entering selection set name
internal class SelectionSetNameForm : System.Windows.Forms.Form
{
    private System.Windows.Forms.TextBox nameTextBox;
    private System.Windows.Forms.Label warningLabel;
    private System.Windows.Forms.Button okButton;
    private System.Windows.Forms.Button cancelButton;
    private List<string> existingNames;
    
    public string SelectionSetName => nameTextBox.Text.Trim();
    
    public SelectionSetNameForm(List<string> existingSelectionSetNames)
    {
        existingNames = existingSelectionSetNames;
        InitializeComponents();
    }
    
    private void InitializeComponents()
    {
        // Form settings
        this.Text = "Save Selection Set";
        this.Size = new System.Drawing.Size(400, 180);
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        
        // Name label
        var nameLabel = new System.Windows.Forms.Label
        {
            Text = "Selection Set Name:",
            Location = new System.Drawing.Point(12, 15),
            Size = new System.Drawing.Size(120, 23),
            TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        };
        
        // Name textbox
        nameTextBox = new System.Windows.Forms.TextBox
        {
            Location = new System.Drawing.Point(140, 15),
            Size = new System.Drawing.Size(230, 23),
            Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        };
        nameTextBox.TextChanged += NameTextBox_TextChanged;
        
        // Warning label
        warningLabel = new System.Windows.Forms.Label
        {
            Text = "",
            Location = new System.Drawing.Point(12, 45),
            Size = new System.Drawing.Size(360, 40),
            ForeColor = System.Drawing.Color.Red,
            Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        };
        
        // OK button
        okButton = new System.Windows.Forms.Button
        {
            Text = "OK",
            Location = new System.Drawing.Point(215, 100),
            Size = new System.Drawing.Size(75, 23),
            Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right,
            Enabled = false
        };
        okButton.Click += (sender, e) => this.DialogResult = System.Windows.Forms.DialogResult.OK;
        
        // Cancel button
        cancelButton = new System.Windows.Forms.Button
        {
            Text = "Cancel",
            Location = new System.Drawing.Point(295, 100),
            Size = new System.Drawing.Size(75, 23),
            Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        };
        cancelButton.Click += (sender, e) => this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        
        // Add controls
        this.Controls.AddRange(new System.Windows.Forms.Control[] { nameLabel, nameTextBox, warningLabel, okButton, cancelButton });
        
        // Set accept and cancel buttons
        this.AcceptButton = okButton;
        this.CancelButton = cancelButton;
    }
    
    private void NameTextBox_TextChanged(object sender, EventArgs e)
    {
        string name = nameTextBox.Text.Trim();
        
        if (string.IsNullOrEmpty(name))
        {
            warningLabel.Text = "";
            okButton.Enabled = false;
        }
        else if (existingNames.Contains(name, StringComparer.OrdinalIgnoreCase))
        {
            warningLabel.Text = $"Warning: Selection set '{name}' already exists.\nIt will be overwritten if you continue.";
            okButton.Enabled = true;
        }
        else
        {
            warningLabel.Text = "";
            okButton.Enabled = true;
        }
    }
}

// Wrapper class for displaying selection sets in DataGrid
internal class SelectionSetWrapper
{
    public string Name { get; set; }
}

// SaveSelection command
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SaveSelection : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = commandData.Application.ActiveUIDocument;
        var doc = uidoc.Document;
        
        try
        {
            // Get current selection using the abstraction
            var currentSelection = uidoc.GetReferences();
            
            if (currentSelection == null || currentSelection.Count == 0)
            {
                TaskDialog.Show("Save Selection", "No elements selected. Please select elements before saving.");
                return Result.Failed;
            }
            
            // Get existing selection set names
            var existingNames = SelectionSetManager.GetSelectionSetNames(doc);
            
            // Show name dialog
            using (var nameForm = new SelectionSetNameForm(existingNames))
            {
                if (nameForm.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return Result.Cancelled;
                
                string selectionSetName = nameForm.SelectionSetName;
                
                if (string.IsNullOrWhiteSpace(selectionSetName))
                {
                    TaskDialog.Show("Save Selection", "Invalid selection set name.");
                    return Result.Failed;
                }
                
                using (Transaction tx = new Transaction(doc, $"Save Selection Set: {selectionSetName}"))
                {
                    tx.Start();
                    
                    // Separate regular and linked element references
                    var regularElementIds = new List<ElementId>();
                    var linkedRefs = new List<SelectionSetLinkedReferenceInfo>();
                    
                    foreach (var reference in currentSelection)
                    {
                        if (reference.LinkedElementId != ElementId.InvalidElementId)
                        {
                            // This is a linked element reference
                            linkedRefs.Add(new SelectionSetLinkedReferenceInfo
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
                    
                    // Find or create the selection set
                    SelectionFilterElement selectionSet = null;
                    var collector = new FilteredElementCollector(doc)
                        .OfClass(typeof(SelectionFilterElement));
                    
                    selectionSet = collector
                        .Cast<SelectionFilterElement>()
                        .FirstOrDefault(s => s.Name.Equals(selectionSetName, StringComparison.OrdinalIgnoreCase));
                    
                    if (selectionSet != null)
                    {
                        // Update existing selection set
                        selectionSet.SetElementIds(regularElementIds);
                    }
                    else
                    {
                        // Create new selection set
                        selectionSet = SelectionFilterElement.Create(doc, selectionSetName);
                        selectionSet.SetElementIds(regularElementIds);
                    }
                    
                    tx.Commit();
                    
                    // Save linked references to file (outside transaction)
                    SelectionSetManager.SaveLinkedReferences(doc, selectionSetName, linkedRefs);
                }
                
                return Result.Succeeded;
            }
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}

// LoadSelection command
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class LoadSelection : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = commandData.Application.ActiveUIDocument;
        var doc = uidoc.Document;
        
        try
        {
            // Get existing selection set names
            var selectionSetNames = SelectionSetManager.GetSelectionSetNames(doc);
            
            if (!selectionSetNames.Any())
            {
                TaskDialog.Show("Load Selection", "No saved selection sets found.");
                return Result.Failed;
            }
            
            // Create wrappers for DataGrid
            var selectionSetWrappers = selectionSetNames.Select(name => new SelectionSetWrapper { Name = name }).ToList();
            
            // Let user choose a selection set using DataGrid
            var propertyNames = new List<string> { "Name" };
            var selectedWrappers = CustomGUIs.DataGrid(selectionSetWrappers, propertyNames, new List<int> { 0 });
            
            if (selectedWrappers == null || !selectedWrappers.Any())
                return Result.Cancelled;
            
            string selectedName = selectedWrappers.First().Name;
            
            // Find the selection set
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(SelectionFilterElement));
            
            var selectionSet = collector
                .Cast<SelectionFilterElement>()
                .FirstOrDefault(s => s.Name.Equals(selectedName, StringComparison.OrdinalIgnoreCase));
            
            if (selectionSet == null)
            {
                TaskDialog.Show("Load Selection", $"Selection set '{selectedName}' not found.");
                return Result.Failed;
            }
            
            // Get regular element IDs from selection set
            var regularElementIds = selectionSet.GetElementIds();
            
            // Load linked references from file
            var linkedRefs = SelectionSetManager.LoadLinkedReferences(doc, selectedName);
            
            // Build complete reference list
            var allReferences = new List<Reference>();
            
            // Add regular elements
            foreach (var elemId in regularElementIds)
            {
                var elem = doc.GetElement(elemId);
                if (elem != null)
                {
                    allReferences.Add(new Reference(elem));
                }
            }
            
            // Add linked elements
            int validLinkedRefs = 0;
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
                                {
                                    allReferences.Add(linkRef);
                                    validLinkedRefs++;
                                }
                            }
                        }
                    }
                }
                catch { /* Skip problematic references */ }
            }
            
            // Set the selection using the abstraction
            uidoc.SetReferences(allReferences);
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
