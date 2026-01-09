// Export selected elements to a new Revit file
// Automatically resolves viewports to their underlying views
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

[Transaction(TransactionMode.Manual)]
public class ExportElementsToRvt : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the selected elements
        ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "Please select at least one element to export.");
            return Result.Failed;
        }

        // Resolve viewports to their underlying views
        List<ElementId> elementsToExport = new List<ElementId>();
        foreach (ElementId selectedId in selectedIds)
        {
            Element selectedElement = doc.GetElement(selectedId);

            // If it's a viewport, get the underlying view
            if (selectedElement is Viewport viewport)
            {
                elementsToExport.Add(viewport.ViewId);
            }
            else
            {
                elementsToExport.Add(selectedId);
            }
        }

        if (elementsToExport.Count == 0)
        {
            TaskDialog.Show("Error", "No valid elements to export after analysis.");
            return Result.Failed;
        }

        // Prompt user to choose location and file name
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "Revit File (*.rvt)|*.rvt";
        saveFileDialog.Title = "Save As";
        saveFileDialog.ShowDialog(new RevitWindow(Helpers.GetMainWindowHandle(commandData.Application)));

        // Check if the user cancelled the operation
        if (saveFileDialog.FileName == "")
            return Result.Cancelled;

        // Create a new Revit document
        var revitApp = uiApp.Application;
        Document newDoc = null;

        try
        {
            newDoc = revitApp.NewProjectDocument(UnitSystem.Imperial);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Failed to create new document: {ex.Message}");
            return Result.Failed;
        }

        // Copy the elements to the new document
        int successCount = 0;
        int failureCount = 0;
        string lastError = "";

        foreach (ElementId elementId in elementsToExport)
        {
            try
            {
                using (Transaction trans = new Transaction(newDoc, "Copy Elements"))
                {
                    trans.Start();

                    CopyPasteOptions copyOptions = new CopyPasteOptions();
                    ICollection<ElementId> copiedElementIds = ElementTransformUtils.CopyElements(
                        doc,
                        new List<ElementId> { elementId },
                        newDoc,
                        Transform.Identity,
                        copyOptions);

                    trans.Commit();

                    if (copiedElementIds != null && copiedElementIds.Count > 0)
                    {
                        successCount++;
                    }
                    else
                    {
                        failureCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                failureCount++;
                lastError = ex.InnerException?.Message ?? ex.Message;
            }
        }

        // Save the new document as .rvt
        string filePath = saveFileDialog.FileName;

        try
        {
            SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true };
            newDoc.SaveAs(filePath, saveAsOptions);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Failed to save document: {ex.Message}");
            return Result.Failed;
        }

        // Show error if some elements failed
        if (failureCount > 0 && successCount == 0)
        {
            TaskDialog.Show("Export Failed",
                $"Failed to export elements.\n\n" +
                $"Error: {lastError}");
            return Result.Failed;
        }
        else if (failureCount > 0)
        {
            TaskDialog.Show("Partial Success",
                $"Export completed with issues:\n" +
                $"- {successCount} element(s) copied successfully\n" +
                $"- {failureCount} element(s) failed to copy\n\n" +
                $"Last error: {lastError}");
        }

        return Result.Succeeded;
    }
}
