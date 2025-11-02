using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;

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
                Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Revit Files (*.rvt)|*.rvt",
                    Title = "Select a Revit File to Import Elements From"
                };

                if (openFileDialog.ShowDialog() != true)
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

                    // Collect all 3D model elements from source
                    // Use multiple collectors for different element classes
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

                    if (sourceElementIds.Count == 0)
                    {
                        TaskDialog.Show("Import Elements",
                            $"No valid 3D elements found in file:\n{filePath}");
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

                    // Copy elements to current document
                    using (Transaction trans = new Transaction(currentDoc, "Import Elements"))
                    {
                        trans.Start();

                        try
                        {
                            // Copy elements from source document
                            CopyPasteOptions options = new CopyPasteOptions();
                            options.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler());

                            ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                sourceDoc,
                                allElementIds.ToList(),
                                currentDoc,
                                Transform.Identity,
                                options);

                            // Filter out types and materials from the imported list for selection
                            foreach (ElementId id in copiedIds)
                            {
                                Element elem = currentDoc.GetElement(id);
                                if (elem != null && 
                                    !(elem is ElementType) && 
                                    !(elem is Material) &&
                                    elem.Category != null)
                                {
                                    importedElementIds.Add(id);
                                }
                            }

                            trans.Commit();
                        }
                        catch (Exception copyEx)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Copy Failed", 
                                $"Failed to copy elements: {copyEx.Message}\n\n" +
                                "This might be due to:\n" +
                                "- Name conflicts\n" +
                                "- Different Revit versions\n" +
                                "- Missing families or types\n" +
                                "- Corrupted elements");
                            return Result.Failed;
                        }
                    }

                    // Set selection to imported elements
                    if (importedElementIds.Count > 0)
                    {
                        uiDoc.SetSelectionIds(importedElementIds);
                        
                        TaskDialog.Show("Import Complete",
                            $"Successfully imported {importedElementIds.Count} element(s) from:\n{System.IO.Path.GetFileName(filePath)}");
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
            int catId = elem.Category.Id.IntegerValue;

            // Exclude non-model categories
            HashSet<int> excludedCategories = new HashSet<int>
            {
                (int)BuiltInCategory.OST_Views,
                (int)BuiltInCategory.OST_Sheets,
                (int)BuiltInCategory.OST_ProjectInformation,
                (int)BuiltInCategory.OST_Schedules,
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
