using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class OpenViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document   doc   = uidoc.Document;
        View       activeView = uidoc.ActiveView;

        // ─────────────────────────────────────────────────────────────
        // 1. Collect every non-template, non-browser view (incl. sheets)
        // ─────────────────────────────────────────────────────────────
        List<View> allViews = new FilteredElementCollector(doc)
            .OfClass(typeof(View))
            .Cast<View>()
            .Where(v =>
                   !v.IsTemplate &&
                   v.ViewType != ViewType.ProjectBrowser &&
                   v.ViewType != ViewType.SystemBrowser)
            .OrderBy(v => v.Title)
            .ToList();

        // ─────────────────────────────────────────────────────────────
        // 2. Prepare data for the grid
        //    • one Dictionary per row
        //    • map title → view so we can open it later
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> gridData =
            new List<Dictionary<string, object>>();
        Dictionary<string, View> titleToView = new Dictionary<string, View>();

        foreach (View v in allViews)
        {
            string sheetFolder =
                v.LookupParameter("Sheet Folder")?.AsString() ?? string.Empty;

            gridData.Add(new Dictionary<string, object>
            {
                { "Title",       v.Title },
                { "SheetFolder", sheetFolder }
            });

            // (assumes titles are unique; adjust if your projects break that rule)
            titleToView[v.Title] = v;
        }

        // Column headers in the order you want them shown
        List<string> columns = new List<string> { "Title", "SheetFolder" };

        // ─────────────────────────────────────────────────────────────
        // 3. Show the grid.  (false = do NOT span across all screens)
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> selectedRows =
            CustomGUIs.DataGrid(gridData, columns, false);

        if (selectedRows != null && selectedRows.Any())
        {
            // Open every selected view (sheet or model view)
            foreach (Dictionary<string, object> row in selectedRows)
            {
                string title = row["Title"].ToString();
                if (titleToView.TryGetValue(title, out View view))
                {
                    uidoc.RequestViewChange(view);
                }
            }
        }

        // match original behaviour – always return Succeeded
        return Result.Succeeded;
    }
}
