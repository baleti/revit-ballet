using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace RevitAddin
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ListChildViewsCommand : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Get the current active view
            View activeView = doc.ActiveView;

            // Filter for all views in the document
            FilteredElementCollector viewCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(View));

            List<View> childViews = new List<View>();

            foreach (View view in viewCollector)
            {
                // Check if the view is a child of the active view
                if (view.GetPrimaryViewId() == activeView.Id)
                {
                    childViews.Add(view);
                }
            }

            // Output the results
            TaskDialog.Show("Child Views",
                "Number of child views: " + childViews.Count + "\n" +
                string.Join("\n", childViews.Select(v => v.Name)));

            return Result.Succeeded;
        }
    }
}
