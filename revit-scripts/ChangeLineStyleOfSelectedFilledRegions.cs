using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;

// Alias Revit's View to avoid ambiguity with System.Windows.Forms.View.
using RevitView = Autodesk.Revit.DB.View;

namespace RevitAddin
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ChangeLineStyleOfSelectedFilledRegions : IExternalCommand
    {
        // File path for storing preferences
        private static readonly string AppDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "revit-scripts"
        );
        private static readonly string PreferencesFilePath = Path.Combine(AppDataPath, "ChangeLineStyleOfSelectedFilledRegions");
        
        // Get project name (sanitized for file storage)
        private static string GetProjectName(Document doc)
        {
            // Get project name from project information
            string projectName = doc.ProjectInformation.Name;
            if (string.IsNullOrWhiteSpace(projectName))
                projectName = "Untitled";
            
            // Remove any characters that might cause issues in the file
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                projectName = projectName.Replace(c, '_');
            }
            projectName = projectName.Replace('[', '_').Replace(']', '_');
            
            return projectName.Trim();
        }
        
        // Structure to store preferences
        private class Preferences
        {
            public string LineStyleName { get; set; }
        }
        
        // Save preferences to file
        private void SavePreferences(Document doc, string lineStyleName)
        {
            try
            {
                // Ensure directory exists
                Directory.CreateDirectory(AppDataPath);
                
                string projectName = GetProjectName(doc);
                List<string> outputLines = new List<string>();
                bool projectFound = false;
                bool inCurrentProject = false;
                
                // Read existing file and update the current project's section
                if (File.Exists(PreferencesFilePath))
                {
                    var lines = File.ReadAllLines(PreferencesFilePath);
                    
                    for (int i = 0; i < lines.Length; i++)
                    {
                        string line = lines[i].Trim();
                        
                        // Check if we're entering a project section
                        if (line.StartsWith("[") && line.EndsWith("]"))
                        {
                            // If we were in our project section and didn't write yet, write before moving to next
                            if (inCurrentProject && !projectFound)
                            {
                                outputLines.Add(lineStyleName ?? "");
                                projectFound = true;
                            }
                            
                            inCurrentProject = (line == $"[{projectName}]");
                            outputLines.Add(line);
                            
                            if (inCurrentProject)
                            {
                                // Skip the next line (old setting) if it exists
                                int linesToSkip = 0;
                                if (i + 1 < lines.Length && !lines[i + 1].Trim().StartsWith("["))
                                    linesToSkip++;
                                i += linesToSkip;
                                
                                // Write new setting
                                outputLines.Add(lineStyleName ?? "");
                                projectFound = true;
                            }
                        }
                        else if (!inCurrentProject)
                        {
                            // Copy lines from other projects
                            outputLines.Add(line);
                        }
                    }
                }
                
                // If project wasn't found, add it at the end
                if (!projectFound)
                {
                    if (outputLines.Count > 0 && !string.IsNullOrWhiteSpace(outputLines.Last()))
                        outputLines.Add(""); // Add blank line for readability
                    
                    outputLines.Add($"[{projectName}]");
                    outputLines.Add(lineStyleName ?? "");
                }
                
                File.WriteAllLines(PreferencesFilePath, outputLines);
            }
            catch { /* Ignore errors saving preferences */ }
        }
        
        // Load preferences from file
        private Preferences LoadPreferences(Document doc)
        {
            var prefs = new Preferences();
            
            try
            {
                if (File.Exists(PreferencesFilePath))
                {
                    string projectName = GetProjectName(doc);
                    var lines = File.ReadAllLines(PreferencesFilePath);
                    bool inCurrentProject = false;
                    
                    for (int i = 0; i < lines.Length; i++)
                    {
                        string line = lines[i].Trim();
                        
                        if (line == $"[{projectName}]")
                        {
                            inCurrentProject = true;
                            continue;
                        }
                        else if (line.StartsWith("[") && line.EndsWith("]"))
                        {
                            inCurrentProject = false;
                            continue;
                        }
                        
                        if (inCurrentProject && !string.IsNullOrWhiteSpace(line))
                        {
                            if (string.IsNullOrEmpty(prefs.LineStyleName))
                            {
                                prefs.LineStyleName = line;
                                break; // We have the value, exit
                            }
                        }
                    }
                }
            }
            catch { /* Return default empty preferences on error */ }
            
            return prefs;
        }
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get document and UIDocument.
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Require that at least one element is selected.
            ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
            if (selIds == null || selIds.Count == 0)
            {
                MessageBox.Show("Please select one or more Filled Region elements first.", "No Selection");
                return Result.Failed;
            }

            // Filter to get only FilledRegion elements from selection
            List<FilledRegion> selectedFilledRegions = new List<FilledRegion>();
            foreach (ElementId id in selIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is FilledRegion fr)
                {
                    selectedFilledRegions.Add(fr);
                }
            }

            if (selectedFilledRegions.Count == 0)
            {
                MessageBox.Show("No valid Filled Region elements selected. Please select one or more Filled Regions.", "No Valid Filled Regions");
                return Result.Failed;
            }

            // Load preferences
            var preferences = LoadPreferences(doc);
            
            // --- Prompt for Line Style selection ---
            // Get the Lines category
            Category linesCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            
            List<Dictionary<string, object>> lineStyleData = new List<Dictionary<string, object>>();
            Dictionary<string, GraphicsStyle> nameToLineStyleMap = new Dictionary<string, GraphicsStyle>();
            
            // Get line styles from the Lines category's subcategories
            foreach (Category lineSubCategory in linesCategory.SubCategories)
            {
                GraphicsStyle gs = lineSubCategory.GetGraphicsStyle(GraphicsStyleType.Projection);
                if (gs != null)
                {
                    string name = lineSubCategory.Name;
                    
                    nameToLineStyleMap[name] = gs;
                    lineStyleData.Add(new Dictionary<string, object>
                    {
                        { "Line Style", name }
                    });
                }
            }
            
            // Sort by name
            lineStyleData = lineStyleData.OrderBy(d => d["Line Style"].ToString()).ToList();
            
            // Find preferred line style index after sorting
            int preferredLineStyleIndex = -1;
            if (!string.IsNullOrEmpty(preferences.LineStyleName))
            {
                preferredLineStyleIndex = lineStyleData.FindIndex(d => 
                    d["Line Style"].ToString() == preferences.LineStyleName);
            }
            
            // Define columns for line style selection
            List<string> lineStyleColumns = new List<string> { "Line Style" };
            
            // Set initial selection if we have a preference
            List<int> initialLineStyleSelection = null;
            if (preferredLineStyleIndex >= 0)
            {
                initialLineStyleSelection = new List<int> { preferredLineStyleIndex };
            }
            
            // Show line style selection dialog
            List<Dictionary<string, object>> selectedLineStyles = CustomGUIs.DataGrid(
                lineStyleData,
                lineStyleColumns,
                false,  // Don't span all screens
                initialLineStyleSelection
            );
            
            // Check if user selected a line style
            GraphicsStyle selectedLineStyle = null;
            string selectedLineStyleName = null;
            if (selectedLineStyles != null && selectedLineStyles.Count > 0)
            {
                selectedLineStyleName = selectedLineStyles[0]["Line Style"].ToString();
                nameToLineStyleMap.TryGetValue(selectedLineStyleName, out selectedLineStyle);
            }
            else
            {
                message = "No line style selected.";
                return Result.Cancelled;
            }
            
            if (selectedLineStyle == null)
            {
                message = "Unable to resolve the selected line style.";
                return Result.Failed;
            }

            // Save preferences for next time
            SavePreferences(doc, selectedLineStyleName);

            // --- Change line styles of selected filled regions ---
            using (Transaction trans = new Transaction(doc, "Change Line Styles of Filled Regions"))
            {
                trans.Start();
                
                foreach (FilledRegion fr in selectedFilledRegions)
                {
                    try
                    {
                        fr.SetLineStyleId(selectedLineStyle.Id);
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Warning", $"Error changing line style for Filled Region {fr.Id}: {ex.Message}");
                    }
                }
                
                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
