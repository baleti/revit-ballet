using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

[Transaction(TransactionMode.Manual)]
public class OpenSavedViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        // Ask user to choose file location and name
        OpenFileDialog openFileDialog = new OpenFileDialog
        {
            Title = "Open Saved UIViews",
            DefaultExt = ".txt",
            Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
            FileName = "revit-saved-UIViews.txt"
        };

        DialogResult result = openFileDialog.ShowDialog();
        if (result != DialogResult.OK)
        {
            // User cancelled the dialog
            return Result.Cancelled;
        }

        string filePath = openFileDialog.FileName;

        // Read the data from the file
        List<Dictionary<string, object>> savedViewTitles = new List<Dictionary<string, object>>();
        using (StreamReader reader = new StreamReader(filePath))
        {
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                var entry = new Dictionary<string, object> { { "Title", line } };
                savedViewTitles.Add(entry);
            }
        }

        // Display the saved views using CustomGUIs.DataGrid
        var selectedViews = CustomGUIs.DataGrid(savedViewTitles, new List<string> { "Title" }, false);

        // Open the selected views in Revit
        UIApplication uiapp = commandData.Application;
        Document doc = uiapp.ActiveUIDocument.Document;
        UIDocument uidoc = uiapp.ActiveUIDocument;

        foreach (var viewEntry in selectedViews)
        {
            if (viewEntry["Title"] == null)
                continue;

            string viewTitle = viewEntry["Title"].ToString();
            Autodesk.Revit.DB.View view = new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .FirstOrDefault(v => v.Title.Equals(viewTitle, StringComparison.OrdinalIgnoreCase));

            if (view != null)
            {
                uidoc.RequestViewChange(view);
            }
        }

        return Result.Succeeded;
    }
}
