using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[TransactionAttribute(TransactionMode.Manual)]
public class SelectViewsFromSets : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Find all ViewSheetSets within the document
        FilteredElementCollector coll = new FilteredElementCollector(doc);
        var viewSheetSets = coll.OfClass(typeof(ViewSheetSet)).Cast<ViewSheetSet>().ToList();

        if (!viewSheetSets.Any())
        {
            TaskDialog.Show("No View Sets", "No view sheet sets found in the document.");
            return Result.Cancelled;
        }

        // Prepare data for the data grid and map view set names to ViewSheetSet objects
        List<Dictionary<string, object>> viewSetData = new List<Dictionary<string, object>>();
        Dictionary<string, ViewSheetSet> nameToViewSetMap = new Dictionary<string, ViewSheetSet>();

        foreach (ViewSheetSet viewSet in viewSheetSets)
        {
            // Get the views in this set
            var viewsInSet = viewSet.Views.Cast<View>().ToList();
            int viewCount = viewsInSet.Count;
            
            // Get sheet count (count ViewSheet types specifically)
            int sheetCount = viewsInSet.Count(v => v is ViewSheet);
            
            // Create a description of the set contents
            string description = $"{viewCount} view(s), {sheetCount} sheet(s)";

            // Map the name to the ViewSheetSet for later retrieval
            nameToViewSetMap[viewSet.Name] = viewSet;
            
            Dictionary<string, object> viewSetInfo = new Dictionary<string, object>
            {
                { "Name", viewSet.Name },
                { "Description", description },
                { "View Count", viewCount }
            };
            viewSetData.Add(viewSetInfo);
        }

        // Define the column headers
        List<string> columns = new List<string> { "Name", "Description", "View Count" };

        // Show the selection dialog (using your custom GUI)
        List<Dictionary<string, object>> selectedViewSets = CustomGUIs.DataGrid(
            viewSetData,
            columns,
            false  // Don't span all screens
        );

        // If the user made a selection, add the views from those sets to the current selection
        if (selectedViewSets != null && selectedViewSets.Any())
        {
            // Get the current selection
            ICollection<ElementId> currentSelectionIds = uidoc.GetSelectionIds();
            
            // Get all views from the selected view sets
            HashSet<ElementId> newViewIds = new HashSet<ElementId>();
            
            foreach (var selectedSet in selectedViewSets)
            {
                string setName = selectedSet["Name"].ToString();
                if (nameToViewSetMap.TryGetValue(setName, out ViewSheetSet viewSet))
                {
                    // Add all views from this set
                    foreach (View view in viewSet.Views.Cast<View>())
                    {
                        newViewIds.Add(view.Id);
                    }
                }
            }
            
            // Combine current selection with new view IDs
            foreach (ElementId id in currentSelectionIds)
            {
                newViewIds.Add(id);
            }
            
            // Update the selection with the combined set of elements
            uidoc.SetSelectionIds(newViewIds);
            
            return Result.Succeeded;
        }

        return Result.Cancelled;
    }
}
