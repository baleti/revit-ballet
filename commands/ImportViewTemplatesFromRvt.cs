using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class ImportViewTemplatesFromRvt : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Prompt user to select the .rvt file containing view templates
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "Revit File (*.rvt)|*.rvt";
        openFileDialog.Title = "Select .rvt File";
        openFileDialog.ShowDialog();

        // Check if the user cancelled the operation
        if (openFileDialog.FileName == "")
            return Result.Cancelled;

        // Open the selected .rvt file
        string filePath = openFileDialog.FileName;
        Document sourceDoc = uiApp.Application.OpenDocumentFile(filePath);

        // Get all view templates from the source document
        FilteredElementCollector collector = new FilteredElementCollector(sourceDoc);
        List<Autodesk.Revit.DB.View> viewTemplates = collector.OfClass(typeof(Autodesk.Revit.DB.View))
                                                    .Cast<Autodesk.Revit.DB.View>()
                                                    .Where(v => v.IsTemplate)
                                                    .ToList();
        
        if (viewTemplates.Count == 0)
        {
            TaskDialog.Show("No View Templates", "The selected .rvt file does not contain any view templates.");
            return Result.Failed;
        }

        // Select view templates to import
        List<string> properties = new List<string> { "Title" };
        var selectedViews = CustomGUIs.DataGrid<Autodesk.Revit.DB.View>(viewTemplates, properties);

        if (selectedViews.Count == 0)
            return Result.Cancelled;

        // Start a transaction to import view templates
        using (Transaction transaction = new Transaction(doc, "Import View Templates"))
        {
            transaction.Start();

            // Copy the selected view templates into the active document
            foreach (Autodesk.Revit.DB.View selectedViewTemplate in selectedViews)
            {
                CopyPasteOptions copyOptions = new CopyPasteOptions();
                copyOptions.SetDuplicateTypeNamesHandler(new DuplicateTypesHandler());
                ElementTransformUtils.CopyElements(sourceDoc, new List<ElementId> { selectedViewTemplate.Id }, doc, Transform.Identity, copyOptions);
            }

            // Commit the transaction
            transaction.Commit();
        }

        // Close the source document
        sourceDoc.Close(false); // Pass false to indicate that changes should not be saved

        return Result.Succeeded;
    }
}
