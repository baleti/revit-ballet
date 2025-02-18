using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class DuplicateSelectedViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

        // Use the currently selected elements.
        List<View> selectedViews = new List<View>();

        foreach (ElementId id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element is View view)
            {
                selectedViews.Add(view);
            }
            else if (element is Viewport viewport)
            {
                // If a viewport is selected, add its corresponding view.
                View viewFromViewport = doc.GetElement(viewport.ViewId) as View;
                if (viewFromViewport != null)
                {
                    selectedViews.Add(viewFromViewport);
                }
            }
        }

        // Start a transaction.
        using (Transaction trans = new Transaction(doc, "Duplicate Selected Views"))
        {
            trans.Start();

            try
            {
                // Loop through each selected view.
                foreach (View view in selectedViews)
                {
                    // Duplicate only if the view is not null and is an AreaPlan.
                    if (view != null && view.ViewType == ViewType.AreaPlan)
                    {
                        // Duplicate the view.
                        ElementId duplicatedViewId = view.Duplicate(ViewDuplicateOption.Duplicate);

                        // Optionally, rename the duplicated view.
                        View duplicatedView = doc.GetElement(duplicatedViewId) as View;
                        if (duplicatedView != null)
                        {
                            string timestamp = DateTime.Now.ToString("yyyyMMddHHmm");
                            duplicatedView.Name = $"{view.Name} - Copy {timestamp}";
                        }
                    }
                }

                // Commit the transaction.
                trans.Commit();
            }
            catch (Exception ex)
            {
                // Roll back the transaction if an error occurs.
                trans.RollBack();
                message = ex.Message;
                return Result.Failed;
            }
        }

        return Result.Succeeded;
    }
}
