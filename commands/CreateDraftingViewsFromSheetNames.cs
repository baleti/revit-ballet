using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CreateDraftingViewsFromSheetNames : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the application and document from command data
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;
        Autodesk.Revit.ApplicationServices.Application app = uiApp.Application;

        // List to store document names
        List<Document> projectNames = new List<Document>();

        // Iterate over all loaded documents
        foreach (Document document in app.Documents)
        {
            // Check if the document is not a link and it's a project document
            if (!document.IsLinked && document.IsFamilyDocument == false)
            {
                projectNames.Add(document);
            }
        }

        List<string> propertyNames = new List<string>() { "Title" };

        var selectedProject  = CustomGUIs.DataGrid(projectNames, propertyNames, null, "Pick Project to Copy Sheet Names from").FirstOrDefault();

        if (selectedProject == null)
            return Result.Failed;

        // Now find all sheets in the selected project
        FilteredElementCollector collector = new FilteredElementCollector(selectedProject);
        ICollection<Element> sheets = collector.OfClass(typeof(ViewSheet)).ToElements();

        // List to store the sheets for selection
        List<ViewSheet> sheetList = sheets.Cast<ViewSheet>().ToList();

        // List of properties to display for sheets
        List<string> sheetPropertyNames = new List<string>() { "SheetNumber", "Name" };

        // Use DataGrid again to display sheets for further processing
        List<ViewSheet> selectedSheets = CustomGUIs.DataGrid(sheetList, sheetPropertyNames, null, "Select Sheets from the Project to Copy Names from");

        if (selectedSheets.Count == 0)
            return Result.Failed;

        // Get the first available ViewFamilyType for a drafting view
        ViewFamilyType draftingViewFamilyType = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewFamilyType))
            .Cast<ViewFamilyType>()
            .FirstOrDefault(x => x.ViewFamily == ViewFamily.Drafting);

        if (draftingViewFamilyType == null)
        {
            // Handle the case where no drafting view family type is found
            message = "No drafting view family type found.";
            return Result.Failed;
        }
        // Create drafting views in the current document for each sheet selected
        Transaction trans = new Transaction(uiDoc.Document, "Create Drafting Views");
        trans.Start();
        try
        {
            // Fetch existing view names to avoid duplicates
            HashSet<string> existingViewNames = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Select(v => v.Name)
            );

            foreach (ViewSheet sheet in selectedSheets)
            {
                if (!existingViewNames.Contains(sheet.Name))
                {
                    ViewDrafting draftingView = ViewDrafting.Create(doc, draftingViewFamilyType.Id);
                    if (draftingView != null)
                    {
                        draftingView.Name = sheet.Name; // Set the name of the drafting view to match the sheet name
                    }
                }
            }
        }
        catch (Exception ex)
        {
            message = ex.Message;
            trans.RollBack();
            return Result.Failed;
        }
        trans.Commit();

        return Result.Succeeded;
    }
}
