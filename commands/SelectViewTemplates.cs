using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SelectViewTemplates : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument   uidoc = uiapp.ActiveUIDocument;
        Document     doc   = uidoc.Document;

        // ───────────────────────────────────────────────────────────────
        // 1. Collect all **view templates**
        // ───────────────────────────────────────────────────────────────
        IList<View> viewTemplates = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v => v.IsTemplate)
            .ToList();

        // ───────────────────────────────────────────────────────────────
        // 2. Build a **usage count** lookup (how many views use each template)
        // ───────────────────────────────────────────────────────────────
        Dictionary<ElementId, int> usageCount = new Dictionary<ElementId, int>();

        foreach (View v in new FilteredElementCollector(doc)
                                 .OfClass(typeof(View))
                                 .Cast<View>())
        {
            ElementId tid = v.ViewTemplateId;
            if (tid == ElementId.InvalidElementId) continue;

            usageCount[tid] = usageCount.TryGetValue(tid, out int n) ? n + 1 : 1;
        }

        // ───────────────────────────────────────────────────────────────
        // 3. Prepare rows for the DataGrid and a title→template map
        // ───────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> templateData  = new List<Dictionary<string, object>>();
        Dictionary<string, View>         titleToTplMap = new Dictionary<string, View>();

        foreach (View tpl in viewTemplates)
        {
            string title = tpl.Title;              // Human-readable name in the PB
            int    used  = usageCount.TryGetValue(tpl.Id, out int cnt) ? cnt : 0;

            titleToTplMap[title] = tpl;

            templateData.Add(new Dictionary<string, object>
            {
                { "Title",      title },
                { "View Type",  tpl.ViewType.ToString() },
                { "Used By",    used }          // number of views that reference this template
            });
        }

        // Column order for the grid
        List<string> columns = new List<string> { "Title", "View Type", "Used By" };

        // ───────────────────────────────────────────────────────────────
        // 4. Show the pick list and capture the user’s choice(s)
        // ───────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> picked =
            CustomGUIs.DataGrid(templateData, columns, spanAllScreens: false);

        // ───────────────────────────────────────────────────────────────
        // 5. If anything was picked, merge into the current selection
        // ───────────────────────────────────────────────────────────────
        if (picked != null && picked.Any())
        {
            ICollection<ElementId> currentSel = uidoc.GetSelectionIds();

            foreach (Dictionary<string, object> row in picked)
            {
                ElementId id = titleToTplMap[row["Title"].ToString()].Id;
                if (!currentSel.Contains(id))
                    currentSel.Add(id);
            }

            uidoc.SetSelectionIds(currentSel);
            return Result.Succeeded;
        }

        return Result.Cancelled;
    }
}
