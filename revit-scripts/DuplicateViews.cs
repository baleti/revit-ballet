using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class DuplicateViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

        // Get all views in the project
        List<View> views = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .ToList();

        List<string> properties = new List<string> { "Title", "ViewType" };

        // Get the selected views from your custom GUI function
        List<View> selectedViews = CustomGUIs.DataGrid(views, properties);

        // Start a transaction
        using (Transaction trans = new Transaction(doc, "Duplicate Area Plans"))
        {
            trans.Start();

            try
            {
                // Loop through each selected view
                foreach (View view in selectedViews)
                {
                    if (view != null && view.ViewType == ViewType.AreaPlan)
                    {
                        // Duplicate the view
                        ElementId duplicatedViewId = view.Duplicate(ViewDuplicateOption.Duplicate);

                        // Optionally, rename the duplicated view
                        View duplicatedView = doc.GetElement(duplicatedViewId) as View;
                        if (duplicatedView != null)
                        {
                            string timestamp = DateTime.Now.ToString("yyyyMMddHHmm");
                            duplicatedView.Name = $"{view.Name} - Copy {timestamp}";
                        }
                    }
                }

                // Commit the transaction
                trans.Commit();
            }
            catch (Exception ex)
            {
                // If something goes wrong, roll back the transaction
                trans.RollBack();
                message = ex.Message;
                return Result.Failed;
            }
        }

        return Result.Succeeded;
    }
}
