using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using RevitBallet.Commands;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ImportElementsFromFile : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                // Get the application and current document
                UIApplication uiApp = commandData.Application;
                Application app = uiApp.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document currentDoc = uiDoc.Document;

                // Prompt user to select Revit file(s)
                System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog
                {
                    Filter = "Revit Files (*.rvt)|*.rvt",
                    Title = "Select Revit File(s) to Import Elements From",
                    Multiselect = true
                };

                if (openFileDialog.ShowDialog(new RevitWindow(Helpers.GetMainWindowHandle(commandData.Application))) != System.Windows.Forms.DialogResult.OK)
                {
                    return Result.Cancelled;
                }

                string[] filePaths = openFileDialog.FileNames;
                List<ElementId> importedElementIds = new List<ElementId>();
                int totalSuccessCount = 0;
                int totalFailureCount = 0;
                int totalSkippedCount = 0;
                int filesProcessed = 0;
                List<string> failureDetails = new List<string>();

                // Process each selected file
                foreach (string filePath in filePaths)
                {
                    Document sourceDoc = null;
                    try
                    {
                        OpenOptions openOptions = new OpenOptions
                        {
                            Audit = false,
                            DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
                        };

                        ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
                        sourceDoc = app.OpenDocumentFile(modelPath, openOptions);

                        // Collect elements from source document
                        List<ElementId> sourceElementIds = new List<ElementId>();

                        // Collect model elements
                        FilteredElementCollector modelCollector = new FilteredElementCollector(sourceDoc)
                            .WhereElementIsNotElementType()
                            .WhereElementIsViewIndependent();

                        foreach (Element elem in modelCollector)
                        {
                            // Skip base points - they cannot be copied
                            if (elem is BasePoint)
                                continue;

                            // Filter for 3D elements with geometry
                            if (elem.Location != null &&
                                elem.Category != null &&
                                elem.get_BoundingBox(null) != null &&
                                IsValidModelElement(elem))
                            {
                                sourceElementIds.Add(elem.Id);
                            }
                        }

                        // Also collect views (legends, schedules, drafting views)
                        FilteredElementCollector viewCollector = new FilteredElementCollector(sourceDoc)
                            .OfClass(typeof(Autodesk.Revit.DB.View));

                        int skippedCount = 0;
                        foreach (Autodesk.Revit.DB.View view in viewCollector)
                        {
                            // Only copy legend views, schedules, and drafting views (not standard views)
                            if (view.ViewType == ViewType.Legend ||
                                view.ViewType == ViewType.Schedule ||
                                view.ViewType == ViewType.DraftingView)
                            {
                                if (!view.IsTemplate)
                                {
                                    // Skip if view with same name and type already exists in target document
                                    if (!ViewExistsInDocument(currentDoc, view.Name, view.ViewType))
                                    {
                                        sourceElementIds.Add(view.Id);
                                    }
                                    else
                                    {
                                        skippedCount++;
                                    }
                                }
                            }
                        }

                        totalSkippedCount += skippedCount;

                        if (sourceElementIds.Count == 0 && skippedCount == 0)
                        {
                            continue; // Skip files with no valid elements at all
                        }

                        // If all elements were skipped, count the file as processed and continue
                        if (sourceElementIds.Count == 0 && skippedCount > 0)
                        {
                            filesProcessed++;
                            continue;
                        }

                        // Copy all elements in a single batch operation
                        // Revit automatically handles dependent types and materials
                        int successCount = 0;
                        int failureCount = 0;
                        string diagnosticPath = System.IO.Path.Combine(
                            RevitBallet.Commands.PathHelper.RuntimeDirectory,
                            "diagnostics",
                            $"ImportElements-{System.DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");
                        var diagnosticLines = new List<string>();
                        diagnosticLines.Add($"=== Import Elements at {System.DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ===");
                        diagnosticLines.Add($"Source file: {System.IO.Path.GetFileName(filePath)}");
                        diagnosticLines.Add($"Elements to copy: {sourceElementIds.Count}");

                        using (Transaction trans = new Transaction(currentDoc, "Import Elements"))
                        {
                            trans.Start();

                            try
                            {
                                // Batch copy all elements - use existing families when names match
                                // Let Revit handle dependent types/materials automatically
                                var startTime = System.DateTime.Now;
                                diagnosticLines.Add($"[{startTime:HH:mm:ss.fff}] Starting batch copy of {sourceElementIds.Count} elements");

                                CopyPasteOptions options = new CopyPasteOptions();
                                options.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler());

                                ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                    sourceDoc,
                                    sourceElementIds,
                                    currentDoc,
                                    Transform.Identity,
                                    options);

                                var endTime = System.DateTime.Now;
                                var duration = (endTime - startTime).TotalMilliseconds;
                                diagnosticLines.Add($"[{endTime:HH:mm:ss.fff}] Batch copy completed in {duration}ms");

                                if (copiedIds != null && copiedIds.Count > 0)
                                {
                                    successCount = sourceElementIds.Count;
                                    diagnosticLines.Add($"Batch copy successful: {copiedIds.Count} elements copied");

                                    // Add instance elements to selection list
                                    foreach (ElementId copiedId in copiedIds)
                                    {
                                        Element copiedElem = currentDoc.GetElement(copiedId);
                                        if (copiedElem != null &&
                                            !(copiedElem is ElementType) &&
                                            !(copiedElem is Material) &&
                                            copiedElem.Category != null)
                                        {
                                            importedElementIds.Add(copiedId);
                                        }
                                    }
                                }
                                else
                                {
                                    failureCount = sourceElementIds.Count;
                                    failureDetails.Add("Batch copy returned null or empty collection");
                                    diagnosticLines.Add("ERROR: Batch copy returned null or empty collection");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Batch copy failed - fall back to individual copying for detailed error reporting
                                diagnosticLines.Add($"ERROR: Batch copy failed - {ex.Message}");
                                diagnosticLines.Add($"Falling back to individual copy for {sourceElementIds.Count} elements");
                                var fallbackStart = System.DateTime.Now;

                                foreach (ElementId elemId in sourceElementIds)
                                {
                                    Element elem = sourceDoc.GetElement(elemId);
                                    string elemInfo = GetElementDescription(elem);

                                    try
                                    {
                                        CopyPasteOptions options = new CopyPasteOptions();
                                        options.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler());

                                        ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                            sourceDoc,
                                            new List<ElementId> { elemId },
                                            currentDoc,
                                            Transform.Identity,
                                            options);

                                        if (copiedIds != null && copiedIds.Count > 0)
                                        {
                                            successCount++;

                                            foreach (ElementId copiedId in copiedIds)
                                            {
                                                Element copiedElem = currentDoc.GetElement(copiedId);
                                                if (copiedElem != null &&
                                                    !(copiedElem is ElementType) &&
                                                    !(copiedElem is Material) &&
                                                    copiedElem.Category != null)
                                                {
                                                    importedElementIds.Add(copiedId);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            failureCount++;
                                            failureDetails.Add($"{elemInfo} - Copy returned null");
                                        }
                                    }
                                    catch (Exception elemEx)
                                    {
                                        failureCount++;
                                        failureDetails.Add($"{elemInfo} - {elemEx.Message}");
                                    }
                                }

                                var fallbackEnd = System.DateTime.Now;
                                var fallbackDuration = (fallbackEnd - fallbackStart).TotalMilliseconds;
                                diagnosticLines.Add($"Fallback completed in {fallbackDuration}ms");
                                diagnosticLines.Add($"Fallback results: {successCount} succeeded, {failureCount} failed");
                            }

                            trans.Commit();
                        }

                        // Save diagnostics
                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
                        System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);

                        totalSuccessCount += successCount;
                        totalFailureCount += failureCount;
                        filesProcessed++;
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Error Opening File",
                            $"Failed to open file:\n{filePath}\n\nError: {ex.Message}");
                    }
                    finally
                    {
                        if (sourceDoc != null && sourceDoc.IsValidObject)
                        {
                            sourceDoc.Close(false);
                        }
                    }
                }

                // Set selection to all imported elements
                if (importedElementIds.Count > 0)
                {
                    uiDoc.SetSelectionIds(importedElementIds);
                }

                // Show result only if there were issues or everything was skipped
                if (filesProcessed == 0)
                {
                    TaskDialog.Show("Import Failed", "No files were processed successfully.");
                    return Result.Failed;
                }
                else if (totalSuccessCount == 0 && totalFailureCount == 0 && totalSkippedCount > 0)
                {
                    // All elements were skipped (duplicates)
                    TaskDialog.Show("Import Complete",
                        $"All {totalSkippedCount} element(s) from {filesProcessed} file(s) already exist in the target document.");
                    return Result.Succeeded;
                }
                else if (totalFailureCount > 0 && totalSuccessCount == 0)
                {
                    TaskDialog.Show("Import Failed",
                        $"All elements failed to copy from {filesProcessed} file(s).");
                    return Result.Failed;
                }
                else if (totalFailureCount > 0 || totalSkippedCount > 0)
                {
                    // Build detailed message
                    string resultMessage = $"Import completed from {filesProcessed} file(s):\n" +
                        $"- {totalSuccessCount} element(s) copied successfully\n";

                    if (totalSkippedCount > 0)
                    {
                        resultMessage += $"- {totalSkippedCount} element(s) skipped (already exist)\n";
                    }

                    if (totalFailureCount > 0)
                    {
                        resultMessage += $"- {totalFailureCount} element(s) failed to copy\n\n" +
                            "Failed elements:\n";

                        // Limit to first 10 failures to avoid overly long dialog
                        int displayCount = Math.Min(failureDetails.Count, 10);
                        for (int i = 0; i < displayCount; i++)
                        {
                            resultMessage += $"  {i + 1}. {failureDetails[i]}\n";
                        }

                        if (failureDetails.Count > 10)
                        {
                            resultMessage += $"\n  ... and {failureDetails.Count - 10} more";
                        }
                    }

                    TaskDialog.Show("Import Complete", resultMessage);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"An error occurred:\n{ex.Message}");
                return Result.Failed;
            }
        }

        private bool IsValidModelElement(Element elem)
        {
            // Check if element is a valid 3D model element
            if (elem == null || elem.Category == null)
                return false;

            // Get category id
            int catId = (int)elem.Category.Id.AsLong();

            // Exclude non-model categories
            HashSet<int> excludedCategories = new HashSet<int>
            {
                (int)BuiltInCategory.OST_Views,
                (int)BuiltInCategory.OST_Sheets,
                (int)BuiltInCategory.OST_ProjectInformation,
#if !REVIT2017
                (int)BuiltInCategory.OST_Schedules,
#endif
                (int)BuiltInCategory.OST_RasterImages,
                (int)BuiltInCategory.OST_Materials,
                (int)BuiltInCategory.OST_Lines,
                (int)BuiltInCategory.OST_RvtLinks,
                (int)BuiltInCategory.OST_Cameras,
                (int)BuiltInCategory.OST_Elev,
                (int)BuiltInCategory.OST_Sections,
                (int)BuiltInCategory.OST_Grids,
                (int)BuiltInCategory.OST_Levels,
                (int)BuiltInCategory.OST_CLines,
                (int)BuiltInCategory.OST_TextNotes,
                (int)BuiltInCategory.OST_Tags,
                (int)BuiltInCategory.OST_DetailComponents,
                (int)BuiltInCategory.OST_Dimensions
            };

            return !excludedCategories.Contains(catId);
        }

        private bool ViewExistsInDocument(Document doc, string viewName, ViewType viewType)
        {
            // Check if a view with the same name and type already exists
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.View));

            foreach (Autodesk.Revit.DB.View existingView in collector)
            {
                if (existingView.ViewType == viewType &&
                    existingView.Name.Equals(viewName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private string GetElementDescription(Element elem)
        {
            if (elem == null)
                return "Unknown element";

            string category = elem.Category?.Name ?? "No Category";
            string name = elem.Name ?? "Unnamed";
            string type = elem.GetType().Name;

            // Get element type name if it's an instance
            if (elem is FamilyInstance fi)
            {
                type = "FamilyInstance";
                Element typeElem = elem.Document.GetElement(fi.GetTypeId());
                if (typeElem != null)
                {
                    name = $"{typeElem.Name} [{name}]";
                }
            }
            else if (elem is ElementType)
            {
                type = "ElementType";
            }
            else if (elem is Material)
            {
                type = "Material";
            }

            return $"{type} - {category} - {name} (ID: {elem.Id})";
        }

        // Handler for duplicate type names - automatically uses existing families
        public class DuplicateTypeNamesHandler : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                // Use destination types if they exist (avoids repeated dialogs)
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }
    }
}
