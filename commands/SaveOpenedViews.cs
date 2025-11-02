using System;
using System.IO;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class SaveOpenedViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        // Get the currently opened UIViews
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        var uiViews = uidoc.GetOpenUIViews();

        // Ask user to choose file location and name
        SaveFileDialog saveFileDialog = new SaveFileDialog
        {
            Title = "Save Opened UIViews",
            DefaultExt = ".txt",
            Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
            FileName = "revit-saved-UIViews.txt"
        };

        DialogResult result = saveFileDialog.ShowDialog();
        if (result != DialogResult.OK)
        {
            // User cancelled the dialog
            return Result.Cancelled;
        }

        string filePath = saveFileDialog.FileName;

        // Collecting the data to write
        using (StreamWriter writer = new StreamWriter(filePath))
        {
            foreach (UIView uiview in uiViews)
            {
                Autodesk.Revit.DB.View view = doc.GetElement(uiview.ViewId) as Autodesk.Revit.DB.View;
                if (view != null)
                {
                    writer.WriteLine(view.Title);
                }
            }
        }

        return Result.Succeeded;
    }
}
