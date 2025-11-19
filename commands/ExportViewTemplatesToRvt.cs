using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class ExportViewTemplatesToRvt : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get all view templates in the document
        FilteredElementCollector collector = new FilteredElementCollector(doc);
        List<Autodesk.Revit.DB.View> viewTemplates = collector.OfClass(typeof(Autodesk.Revit.DB.View))
                                                    .Cast<Autodesk.Revit.DB.View>()
                                                    .Where(v => v.IsTemplate)
                                                    .ToList();

        List<string> properties = new List<string> { "Title" };

        var selectedViews = CustomGUIs.DataGrid<Autodesk.Revit.DB.View>(viewTemplates, properties);

        if (selectedViews.Count == 0)
            return Result.Cancelled;

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
        Document newDoc = revitApp.NewProjectDocument(UnitSystem.Metric);

        // Start a transaction
        using (Transaction transaction = new Transaction(newDoc, "Export View Templates"))
        {
            transaction.Start();

            // Copy the selected elements to the new document
            List<ElementId> copiedIds = new List<ElementId>();
            foreach (Autodesk.Revit.DB.View selectedViewTemplate in selectedViews)
            {
                CopyPasteOptions copyOptions = new CopyPasteOptions();
                copyOptions.SetDuplicateTypeNamesHandler(new DuplicateTypesHandler());
                ElementId copiedId = ElementTransformUtils.CopyElements(doc, new List<ElementId> { selectedViewTemplate.Id }, newDoc, Transform.Identity, copyOptions).FirstOrDefault();
                copiedIds.Add(copiedId);
            }

            // Commit the transaction
            transaction.Commit();
        }

        // Save the new document as .rvt
        string filePath = saveFileDialog.FileName;
        SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true };
        newDoc.SaveAs(filePath, saveAsOptions);
        newDoc.Close(false); // Pass false to indicate that changes should not be saved

        return Result.Succeeded;
    }
}
