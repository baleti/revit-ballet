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
    public class ExportSelectedElementsToFile : IExternalCommand
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

                // Get selected elements
                ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
                
                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("Export Elements", 
                        "No elements selected. Please select elements to export.");
                    return Result.Cancelled;
                }

                // Filter out non-3D elements and element types
                List<ElementId> validElementIds = new List<ElementId>();
                foreach (ElementId id in selectedIds)
                {
                    Element elem = currentDoc.GetElement(id);
                    if (elem != null && 
                        !(elem is ElementType) && 
                        elem.Location != null &&
                        elem.Category != null &&
                        elem.get_BoundingBox(null) != null)
                    {
                        validElementIds.Add(id);
                    }
                }

                if (validElementIds.Count == 0)
                {
                    TaskDialog.Show("Export Elements", 
                        "No valid 3D element instances found in selection.");
                    return Result.Cancelled;
                }

                // Prompt user for save location
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Revit Files (*.rvt)|*.rvt",
                    Title = "Save Exported Elements As",
                    FileName = $"ExportedElements_{DateTime.Now:yyyyMMdd_HHmmss}.rvt"
                };

                if (saveFileDialog.ShowDialog() != true)
                {
                    return Result.Cancelled;
                }

                string filePath = saveFileDialog.FileName;

                // Create new document
                Document newDoc = null;
                try
                {
                    // Attempt to use the default project template if available; otherwise create an empty project without one
                    string templatePath = app.DefaultProjectTemplate;
                    if (!string.IsNullOrEmpty(templatePath))
                    {
                        newDoc = app.NewProjectDocument(templatePath);
                    }
                    else
                    {
                        // Create an empty project document without a template (defaults to metric units; change to UnitSystem.Imperial if preferred)
                        newDoc = app.NewProjectDocument(UnitSystem.Metric);
                    }


                    // Collect dependent elements (like types, materials, etc.)
                    HashSet<ElementId> allElementIds = new HashSet<ElementId>(validElementIds);
                    
                    // Add element types
                    foreach (ElementId id in validElementIds)
                    {
                        Element elem = currentDoc.GetElement(id);
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

                    // Copy elements to new document
                    using (Transaction trans = new Transaction(newDoc, "Import Elements"))
                    {
                        trans.Start();

                        try
                        {
                            // Copy elements with their dependencies
                            CopyPasteOptions options = new CopyPasteOptions();
                            options.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler());

                            ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                currentDoc,
                                allElementIds.ToList(),
                                newDoc,
                                Transform.Identity,
                                options);

                            trans.Commit();

                            // Save the new document
                            SaveAsOptions saveOptions = new SaveAsOptions
                            {
                                OverwriteExistingFile = true,
                                Compact = true
                            };

                            newDoc.SaveAs(filePath, saveOptions);

                            TaskDialog.Show("Export Complete",
                                $"Successfully exported {validElementIds.Count} element(s) to:\n{System.IO.Path.GetFileName(filePath)}");
                        }
                        catch (Exception ex)
                        {
                            trans.RollBack();
                            throw new Exception($"Failed to copy elements: {ex.Message}");
                        }
                    }
                }
                finally
                {
                    if (newDoc != null && newDoc.IsValidObject)
                    {
                        newDoc.Close(false);
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
                // Create new types to avoid conflicts
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }
    }
}
