using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class CopySelectedElementsToViews : IExternalCommand
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
        ICollection<ElementId> selectedElementIds = uidoc.GetSelectionIds();
        if (!selectedElementIds.Any())
        {
            TaskDialog.Show("Error", "No elements selected. Please select elements to copy.");
            return Result.Failed;
        }
        
        // Get all views in the document, including the current view
        View currentView = doc.ActiveView;
        var views = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate)
                    .OrderBy(v => v.Title)
                    .ToList();
        
        // Convert views to dictionary format for DataGrid
        var viewDictionaries = new List<Dictionary<string, object>>();
        int currentViewIndex = -1;
        
        for (int i = 0; i < views.Count; i++)
        {
            var view = views[i];
            var dict = new Dictionary<string, object>
            {
                { "Title", view.Title },
                { "ViewObject", view }  // Store the actual View object for later retrieval
            };
            viewDictionaries.Add(dict);
            
            // Track the index of the current view
            if (view.Id == currentView.Id)
            {
                currentViewIndex = i;
            }
        }
        
        // Set initial selection to the current view
        List<int> initialSelectionIndices = currentViewIndex >= 0 
            ? new List<int> { currentViewIndex } 
            : null;
        
        // Use the CustomGUIs.DataGrid to prompt the user to select target views
        List<Dictionary<string, object>> selectedViewDicts = CustomGUIs.DataGrid(
            viewDictionaries,
            new List<string> { "Title" },
            spanAllScreens: false,
            initialSelectionIndices: initialSelectionIndices
        );
        
        if (!selectedViewDicts.Any())
        {
            message = "No target views selected.";
            return Result.Failed;
        }
        
        // Extract the View objects from the selected dictionaries
        List<View> selectedViews = selectedViewDicts
            .Select(dict => dict["ViewObject"] as View)
            .Where(v => v != null && v.Id != currentView.Id)  // Exclude current view from copy targets
            .ToList();
        
        if (!selectedViews.Any())
        {
            message = "No valid target views selected (current view was excluded).";
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
