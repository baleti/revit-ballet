using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectViewTemplatesOfSelectedViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData,
                          ref string          message,
                          ElementSet          elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument   uidoc  = uiapp.ActiveUIDocument;
        Document     doc    = uidoc.Document;

        // ───────────────────────────────────────────────────────────────
        // 1. Get the user’s current selection (views or viewports).
        //    If nothing is selected, fall back to the active view.
        // ───────────────────────────────────────────────────────────────
        ICollection<ElementId> pickedIds = uidoc.Selection.GetElementIds();

        if (pickedIds == null || !pickedIds.Any())
        {
            pickedIds = new List<ElementId> { uidoc.ActiveView.Id };
        }

        // ───────────────────────────────────────────────────────────────
        // 2. For each picked element, find the associated view
        //    and collect its template (if any).
        // ───────────────────────────────────────────────────────────────
        HashSet<ElementId> templateIds = new HashSet<ElementId>();

        foreach (ElementId id in pickedIds)
        {
            if (id == ElementId.InvalidElementId) continue;

            Element element = doc.GetElement(id);
            if (element == null) continue;

            View view = element as View;

            if (view == null && element is Viewport vp)
                view = doc.GetElement(vp.ViewId) as View;

            if (view == null) continue;

            ElementId tplId = view.ViewTemplateId;
            if (tplId != ElementId.InvalidElementId)
                templateIds.Add(tplId);
        }

        // ───────────────────────────────────────────────────────────────
        // 3. Replace the current selection with the found templates.
        // ───────────────────────────────────────────────────────────────
        if (templateIds.Any())
        {
            uidoc.Selection.SetElementIds(templateIds);
            return Result.Succeeded;
        }

        TaskDialog.Show("Select View Templates",
                        "No view template was found on the selected items or the active view.");
        return Result.Cancelled;
    }
}
