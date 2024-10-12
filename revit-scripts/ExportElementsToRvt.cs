// doesn't work ModificationOutsideTransactoionException
// also doesn't work for view specific objects
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class ExportElementsToRvt : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the selected elements
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "Please select at least one element to export.");
            return Result.Failed;
        }

        // Prompt user to choose location and file name
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "Revit File (*.rvt)|*.rvt";
        saveFileDialog.Title = "Save As";
        saveFileDialog.ShowDialog();

        // Check if the user cancelled the operation
        if (saveFileDialog.FileName == "")
            return Result.Cancelled;

        // Create a new Revit document
        var revitApp = uiApp.Application;
        Document newDoc = revitApp.NewProjectDocument(UnitSystem.Imperial);

        // Copy the selected elements to the new document
        List<ElementId> copiedIds = new List<ElementId>();
        foreach (ElementId selectedId in selectedIds)
        {
            Element selectedElement = doc.GetElement(selectedId);
            CopyPasteOptions copyOptions = new CopyPasteOptions();
            ElementId copiedId = ElementTransformUtils.CopyElements(doc, new List<ElementId> { selectedId }, newDoc, Transform.Identity, copyOptions).FirstOrDefault();
            copiedIds.Add(copiedId);
        }

        // Save the new document as .rvt
        string filePath = saveFileDialog.FileName;
        SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true };
        newDoc.SaveAs(filePath, saveAsOptions);

        TaskDialog.Show("Success", "Selected objects have been exported to a new .rvt file.");

        return Result.Succeeded;
    }
}
