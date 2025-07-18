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
    public class ImportScopeBoxesFromFile : IExternalCommand
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
                    Title = "Select a Revit File to Import Scope Boxes From"
                };

                if (openFileDialog.ShowDialog() != true)
                {
                    return Result.Cancelled;
                }

                string filePath = openFileDialog.FileName;
                List<ElementId> importedScopeBoxIds = new List<ElementId>();

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

                    // Collect all scope boxes from source
                    List<ElementId> sourceScopeBoxIds = new FilteredElementCollector(sourceDoc)
                        .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                        .WhereElementIsNotElementType()
                        .ToElementIds()
                        .ToList();

                    if (sourceScopeBoxIds.Count == 0)
                    {
                        TaskDialog.Show("Import Scope Boxes",
                            $"No scope boxes found in file:\n{filePath}");
                        return Result.Succeeded;
                    }

                    // Use Copy/Paste approach which works better for scope boxes
                    using (Transaction trans = new Transaction(currentDoc, "Import Scope Boxes"))
                    {
                        trans.Start();

                        try
                        {
                            // Copy scope boxes from source document
                            CopyPasteOptions options = new CopyPasteOptions();
                            options.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler());

                            ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                sourceDoc,
                                sourceScopeBoxIds,
                                currentDoc,
                                Transform.Identity,
                                options);

                            importedScopeBoxIds = copiedIds.ToList();
                        }
                        catch (Exception copyEx)
                        {
                            // If copy fails, show specific error
                            trans.RollBack();
                            TaskDialog.Show("Copy Failed", 
                                $"Failed to copy scope boxes: {copyEx.Message}\n\n" +
                                "This might be due to:\n" +
                                "- Name conflicts\n" +
                                "- Different Revit versions\n" +
                                "- Corrupted elements\n\n" +
                                "Try using Revit's built-in Copy/Paste functionality instead.");
                            return Result.Failed;
                        }

                        // Verify the copied elements are scope boxes
                        List<ElementId> validScopeBoxIds = new List<ElementId>();
                        foreach (ElementId id in importedScopeBoxIds)
                        {
                            Element elem = currentDoc.GetElement(id);
                            if (elem != null && elem.Category != null && 
                                elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_VolumeOfInterest)
                            {
                                validScopeBoxIds.Add(id);
                            }
                        }
                        importedScopeBoxIds = validScopeBoxIds;

                        trans.Commit();
                    }

                    // Set selection to imported scope boxes
                    if (importedScopeBoxIds.Count > 0)
                    {
                        uiDoc.SetSelectionIds(importedScopeBoxIds);
                        
                        TaskDialog.Show("Import Complete",
                            $"Successfully imported {importedScopeBoxIds.Count} scope box(es) from:\n{System.IO.Path.GetFileName(filePath)}");
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

        // Handler for duplicate type names
        public class DuplicateTypeNamesHandler : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                // Use destination types if they exist, otherwise abort
                // This prevents errors when scope box names already exist
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }
    }
}
