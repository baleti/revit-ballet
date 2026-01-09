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

                // Prompt user to select a Revit file
                System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog
                {
                    Filter = "Revit Files (*.rvt)|*.rvt",
                    Title = "Select a Revit File to Import Elements From"
                };

                if (openFileDialog.ShowDialog(new RevitWindow(Helpers.GetMainWindowHandle(commandData.Application))) != System.Windows.Forms.DialogResult.OK)
                {
                    return Result.Cancelled;
                }

                string filePath = openFileDialog.FileName;
                List<ElementId> importedElementIds = new List<ElementId>();

                // Open the source document
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

                    foreach (Autodesk.Revit.DB.View view in viewCollector)
                    {
                        // Only copy legend views, schedules, and drafting views (not standard views)
                        if (view.ViewType == ViewType.Legend ||
                            view.ViewType == ViewType.Schedule ||
                            view.ViewType == ViewType.DraftingView)
                        {
                            if (!view.IsTemplate)
                            {
                                sourceElementIds.Add(view.Id);
                            }
                        }
                    }

                    if (sourceElementIds.Count == 0)
                    {
                        TaskDialog.Show("Import Elements",
                            $"No valid elements found in file:\n{filePath}");
                        return Result.Succeeded;
                    }

                    // Collect dependent elements (types, materials, etc.)
                    HashSet<ElementId> allElementIds = new HashSet<ElementId>(sourceElementIds);

                    foreach (ElementId id in sourceElementIds)
                    {
                        Element elem = sourceDoc.GetElement(id);

                        // Add element type
                        ElementId typeId = elem.GetTypeId();
                        if (typeId != ElementId.InvalidElementId)
                        {
                            allElementIds.Add(typeId);
                        }

                        // Add materials
                        ICollection<ElementId> materialIds = elem.GetMaterialIds(false);
                        foreach (ElementId matId in materialIds)
                        {
                            allElementIds.Add(matId);
                        }
                    }

                    // Copy elements to current document (try each individually to handle failures gracefully)
                    int successCount = 0;
                    int failureCount = 0;

                    using (Transaction trans = new Transaction(currentDoc, "Import Elements"))
                    {
                        trans.Start();

                        foreach (ElementId elemId in allElementIds)
                        {
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

                                    // Add to selection list if it's not a type or material
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
                                }
                            }
                            catch (Exception)
                            {
                                failureCount++;
                            }
                        }

                        trans.Commit();
                    }

                    // Set selection to imported elements
                    if (importedElementIds.Count > 0)
                    {
                        uiDoc.SetSelectionIds(importedElementIds);
                    }

                    // Show result if some elements failed
                    if (failureCount > 0 && successCount == 0)
                    {
                        TaskDialog.Show("Import Failed", "All elements failed to copy.");
                        return Result.Failed;
                    }
                    else if (failureCount > 0)
                    {
                        TaskDialog.Show("Partial Success",
                            $"Import completed with issues:\n" +
                            $"- {successCount} element(s) copied successfully\n" +
                            $"- {failureCount} element(s) failed to copy");
                    }
                }
                finally
                {
                    if (sourceDoc != null && sourceDoc.IsValidObject)
                    {
                        sourceDoc.Close(false);
                    }
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

        // Handler for duplicate type names
        public class DuplicateTypeNamesHandler : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                // Use destination types if they exist
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }
    }
}
