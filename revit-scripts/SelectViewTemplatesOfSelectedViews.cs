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
        // 1. Get whatever the user has selected in Revit.
        // ───────────────────────────────────────────────────────────────
        ICollection<ElementId> pickedIds = uidoc.Selection.GetElementIds();

        if (pickedIds == null || !pickedIds.Any())
        {
            TaskDialog.Show("Select View Templates",
                            "Please select one or more views or viewports first.");
            return Result.Cancelled;
        }

        // ───────────────────────────────────────────────────────────────
        // 2. For each picked element, find the associated view
        //    and collect its template (if any).
        // ───────────────────────────────────────────────────────────────
        HashSet<ElementId> templateIds = new HashSet<ElementId>();

        foreach (ElementId id in pickedIds)
        {
            Element element = doc.GetElement(id);
            if (element == null) continue;

            View view = null;

            switch (element)
            {
                case View v:
                    view = v;
                    break;

                case Viewport vp:
                    view = doc.GetElement(vp.ViewId) as View;
                    break;
            }

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
                        "None of the selected views or viewports use a view template.");
        return Result.Cancelled;
    }
}
