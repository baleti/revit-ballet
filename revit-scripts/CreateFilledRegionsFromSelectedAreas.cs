using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using Autodesk.Revit.DB.Architecture;

// Alias Revit's View to avoid ambiguity with System.Windows.Forms.View.
using RevitView = Autodesk.Revit.DB.View;

namespace RevitAddin
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CreateFilledRegionsFromSelectedAreas : IExternalCommand
    {
        // File path for storing preferences
        private static readonly string AppDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "revit-scripts"
        );
        private static readonly string PreferencesFilePath = Path.Combine(AppDataPath, "CreateFilledRegionsFromSelectedAreas");
        
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
            public string FilledRegionTypeName { get; set; }
            public string LineStyleName { get; set; }
        }
        
        // Save preferences to file
        private void SavePreferences(Document doc, string filledRegionTypeName, string lineStyleName)
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
                                outputLines.Add(filledRegionTypeName ?? "");
                                outputLines.Add(lineStyleName ?? "");
                                projectFound = true;
                            }
                            
                            inCurrentProject = (line == $"[{projectName}]");
                            outputLines.Add(line);
                            
                            if (inCurrentProject)
                            {
                                // Skip the next two lines (old settings) if they exist
                                int linesToSkip = 0;
                                if (i + 1 < lines.Length && !lines[i + 1].Trim().StartsWith("["))
                                    linesToSkip++;
                                if (i + 2 < lines.Length && !lines[i + 2].Trim().StartsWith("["))
                                    linesToSkip++;
                                i += linesToSkip;
                                
                                // Write new settings
                                outputLines.Add(filledRegionTypeName ?? "");
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
                    outputLines.Add(filledRegionTypeName ?? "");
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
                            if (string.IsNullOrEmpty(prefs.FilledRegionTypeName))
                                prefs.FilledRegionTypeName = line;
                            else if (string.IsNullOrEmpty(prefs.LineStyleName))
                            {
                                prefs.LineStyleName = line;
                                break; // We have both values, exit
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
            RevitView activeView = doc.ActiveView;

            // Check if the active view supports filled regions
            if (activeView.ViewType == ViewType.Schedule)
            {
                MessageBox.Show("Filled regions cannot be created in 3D views or schedules. Please switch to a plan, section, or elevation view.", "Invalid View Type");
                return Result.Failed;
            }

            // Require that at least one element is selected.
            ICollection<ElementId> selIds = uidoc.GetSelectionIds();
            if (selIds == null || selIds.Count == 0)
            {
                MessageBox.Show("Please select one or more Area elements first.", "No Selection");
                return Result.Failed;
            }

            // Filter to get only Area elements from selection
            List<Area> selectedAreas = new List<Area>();
            foreach (ElementId id in selIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is Area area)
                {
                    // Check if area is placed and has valid boundaries
                    if (area.Area > 0)
                    {
                        selectedAreas.Add(area);
                    }
                }
            }

            if (selectedAreas.Count == 0)
            {
                MessageBox.Show("No valid Area elements selected. Please select one or more placed Areas.", "No Valid Areas");
                return Result.Failed;
            }

            // Load preferences
            var preferences = LoadPreferences(doc);
            
            // --- Prompt for Filled Region Type selection ---
            // Collect all available FilledRegionTypes.
            FilteredElementCollector regionTypeCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType));
            List<Dictionary<string, object>> regionTypeData = new List<Dictionary<string, object>>();
            // Map filled region type name to FilledRegionType object.
            Dictionary<string, FilledRegionType> nameToRegionTypeMap = new Dictionary<string, FilledRegionType>();
            
            int preferredRegionTypeIndex = -1;
            int currentIndex = 0;
            
            foreach (FilledRegionType frt in regionTypeCollector.Cast<FilledRegionType>())
            {
                string name = frt.Name;
                // Add to mapping and data grid list.
                nameToRegionTypeMap[name] = frt;
                regionTypeData.Add(new Dictionary<string, object>
                {
                    { "Name", name }
                });
                
                // Check if this matches the preferred type
                if (!string.IsNullOrEmpty(preferences.FilledRegionTypeName) && 
                    name == preferences.FilledRegionTypeName)
                {
                    preferredRegionTypeIndex = currentIndex;
                }
                currentIndex++;
            }

            // Define the columns for the filled region type selection.
            List<string> regionTypeColumns = new List<string> { "Name" };
            
            // Set initial selection if we have a preference
            List<int> initialRegionTypeSelection = null;
            if (preferredRegionTypeIndex >= 0)
            {
                initialRegionTypeSelection = new List<int> { preferredRegionTypeIndex };
            }
            
            List<Dictionary<string, object>> selectedRegionTypes = CustomGUIs.DataGrid(
                regionTypeData,
                regionTypeColumns,
                false,  // Don't span all screens.
                initialRegionTypeSelection
            );

            // Ensure a selection was made.
            if (selectedRegionTypes == null || selectedRegionTypes.Count == 0)
            {
                message = "No filled region type selected.";
                return Result.Cancelled;
            }

            // For simplicity, take the first selected filled region type.
            string selectedRegionTypeName = selectedRegionTypes[0]["Name"].ToString();
            if (!nameToRegionTypeMap.TryGetValue(selectedRegionTypeName, out FilledRegionType selectedRegionType))
            {
                message = "Unable to resolve the selected filled region type.";
                return Result.Failed;
            }

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
            // If no line style selected, we'll use the default (null)
            
            // Save preferences for next time
            SavePreferences(doc, selectedRegionTypeName, selectedLineStyleName);

            // --- Create filled regions from area boundaries ---
            List<ElementId> createdFilledRegionIds = new List<ElementId>();
            
            using (Transaction trans = new Transaction(doc, "Create Filled Regions from Areas"))
            {
                trans.Start();
                
                int successCount = 0;
                int failCount = 0;
                
                foreach (Area area in selectedAreas)
                {
                    try
                    {
                        // Get the boundary segments of the area
                        IList<IList<BoundarySegment>> segments = area.GetBoundarySegments(new SpatialElementBoundaryOptions());
                        
                        if (segments == null || segments.Count == 0)
                        {
                            failCount++;
                            continue;
                        }

                        // Create curve loops from boundary segments
                        List<CurveLoop> curveLoops = new List<CurveLoop>();
                        
                        foreach (IList<BoundarySegment> segmentList in segments)
                        {
                            if (segmentList.Count == 0)
                                continue;
                                
                            CurveLoop curveLoop = new CurveLoop();
                            
                            foreach (BoundarySegment segment in segmentList)
                            {
                                Curve curve = segment.GetCurve();
                                if (curve != null)
                                {
                                    curveLoop.Append(curve);
                                }
                            }
                            
                            // Only add if we have a valid curve loop
                            if (curveLoop.Count() > 0)
                            {
                                curveLoops.Add(curveLoop);
                            }
                        }

                        if (curveLoops.Count == 0)
                        {
                            failCount++;
                            continue;
                        }

                        // Create the filled region in the current view
                        FilledRegion filledRegion = FilledRegion.Create(doc, selectedRegionType.Id, activeView.Id, curveLoops);
                        
                        if (filledRegion != null)
                        {
                            // Apply line style if one was selected
                            if (selectedLineStyle != null)
                            {
                                filledRegion.SetLineStyleId(selectedLineStyle.Id);
                            }
                            
                            // Add to created list
                            createdFilledRegionIds.Add(filledRegion.Id);
                            successCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        // Log the error but continue with other areas
                        TaskDialog.Show("Warning", $"Error creating filled region for Area {area.Id}: {ex.Message}");
                    }
                }
                
                trans.Commit();
                
                // Set the created filled regions as the current selection
                if (createdFilledRegionIds.Count > 0)
                {
                    uidoc.SetSelectionIds(createdFilledRegionIds);
                }
            }

            return Result.Succeeded;
        }
    }
}
