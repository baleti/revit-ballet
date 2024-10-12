using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class PasteSelectedElementToViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        // Get the current document
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the selected elements in the active view
        ICollection<ElementId> selectedElementIds = uidoc.Selection.GetElementIds();
        if (!selectedElementIds.Any())
        {
            TaskDialog.Show("Error","No elements selected. Please select elements to copy.");
            return Result.Failed;
        }

        // Get all views in the document, excluding the current view
        View currentView = doc.ActiveView;
        var views = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate && v.Id != currentView.Id)
                    .OrderBy(v => v.Title)
                    .ToList();

        // Use the CustomGUIs.DataGrid to prompt the user to select target views
        List<View> selectedViews = CustomGUIs.DataGrid(
            views, 
            new List<string> { "Title" }, 
            Title: "Select Views to Paste Elements"
        );

        if (!selectedViews.Any())
        {
            message = "No target views selected.";
            return Result.Failed;
        }

        // Start a transaction to copy and paste the elements
        using (Transaction transaction = new Transaction(doc, "Paste in Same Place to Views"))
        {
            transaction.Start();

            CopyPasteOptions options = new CopyPasteOptions();
            foreach (View targetView in selectedViews)
            {
                try
                {
                    ElementTransformUtils.CopyElements(
                        currentView, 
                        selectedElementIds, 
                        targetView, 
                        Transform.Identity, 
                        options);
                }
                catch (Exception ex)
                {
                    message = $"Error copying elements to view {targetView.Title}: {ex.Message}";
                    transaction.RollBack();
                    return Result.Failed;
                }
            }

            transaction.Commit();
        }

        return Result.Succeeded;
    }
}
