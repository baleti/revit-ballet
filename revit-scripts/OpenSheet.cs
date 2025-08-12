using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class OpenSheet : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        // Get the current Revit application and document
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the active view
        View activeView = doc.ActiveView;
        string activeViewName = activeView.Name;

        // Check if the active view is placed on a sheet
        FilteredElementCollector collector = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet));

        ViewSheet sheetContainingView = collector
            .Cast<ViewSheet>()
            .FirstOrDefault(sheet => sheet.GetAllPlacedViews().Contains(activeView.Id));

        // If the view is placed on a sheet, switch to the sheet and close the previous view
        if (sheetContainingView != null)
        {
            // Switch to the sheet view
            uiDoc.RequestViewChange(sheetContainingView);

            // Close the previous view
            CloseViewByName(uiDoc, activeViewName);

            return Result.Succeeded;
        }
        else
        {
            // Inform the user that the active view is not placed on any sheet
            TaskDialog.Show("Open Sheet", "The active view is not placed on any sheet.");
            return Result.Failed;
        }
    }

    private void CloseViewByName(UIDocument uiDoc, string viewName)
    {
        // Get the list of all open views
        var openUIViews = uiDoc.GetOpenUIViews();

        // Find the view with the matching name
        var viewToClose = openUIViews
            .Select(uiView => uiDoc.Document.GetElement(uiView.ViewId) as View)
            .FirstOrDefault(view => view.Name == viewName);

        // Close the matching view if found
        if (viewToClose != null)
        {
            var uiViewToClose = openUIViews.FirstOrDefault(uiView => uiView.ViewId == viewToClose.Id);
            if (uiViewToClose != null)
            {
                uiViewToClose.Close();
            }
        }
    }
}
