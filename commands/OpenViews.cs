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
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> gridData =
            new List<Dictionary<string, object>>();
        Dictionary<string, View> titleToView =
            new Dictionary<string, View>();

        foreach (View v in allViews)
        {
            string sheetFolder =
                v.LookupParameter("Sheet Folder")?.AsString() ?? string.Empty;

            gridData.Add(new Dictionary<string, object>
            {
                { "Title",       v.Title },
                { "SheetFolder", sheetFolder }
            });

            // (assumes titles are unique; adjust if needed)
            titleToView[v.Title] = v;
        }

        // Column headers (order determines column order)
        List<string> columns = new List<string> { "Title", "SheetFolder" };

        // ─────────────────────────────────────────────────────────────
        // 3. Figure out which row should be pre-selected
        // ─────────────────────────────────────────────────────────────
        int selectedIndex = -1;

        if (activeView is ViewSheet)
        {
            selectedIndex = allViews.FindIndex(v => v.Id == activeView.Id);
        }
        else
        {
            Viewport vp = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .FirstOrDefault(vpt => vpt.ViewId == activeView.Id);

            if (vp != null)
            {
                ViewSheet containingSheet = doc.GetElement(vp.SheetId) as ViewSheet;
                if (containingSheet != null)
                    selectedIndex = allViews.FindIndex(v => v.Id == containingSheet.Id);
            }

            if (selectedIndex == -1) // not on a sheet
                selectedIndex = allViews.FindIndex(v => v.Id == activeView.Id);
        }

        List<int> initialSelectionIndices = selectedIndex >= 0
            ? new List<int> { selectedIndex }
            : new List<int>();

        // ─────────────────────────────────────────────────────────────
        // 4. Show the grid
        // ─────────────────────────────────────────────────────────────
        List<Dictionary<string, object>> selectedRows =
            CustomGUIs.DataGrid(gridData, columns, spanAllScreens: false, initialSelectionIndices);

        // ─────────────────────────────────────────────────────────────
        // 5. Open every selected view (sheet or model view)
        // ─────────────────────────────────────────────────────────────
        if (selectedRows != null && selectedRows.Any())
        {
            foreach (Dictionary<string, object> row in selectedRows)
            {
                string title = row["Title"].ToString();
                View view;
                if (titleToView.TryGetValue(title, out view))
                    uidoc.RequestViewChange(view);
            }
        }

        return Result.Succeeded;
    }
}
