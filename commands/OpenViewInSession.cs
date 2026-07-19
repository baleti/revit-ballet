using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("")]
public class OpenViewInSession : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document activeDoc = uidoc?.Document;

        if (activeDoc == null)
        {
            TaskDialog.Show("Error", "No active document.");
            return Result.Failed;
        }

        View activeView = uidoc.ActiveView;

        List<Dictionary<string, object>> gridData;
        List<string> columns;

        if (ViewDataCache.TryGetSessionCache(uiApp.Application, "views", out gridData, out columns))
        {
            // Cache hit — nothing to rebuild
        }
        else
        {
            gridData = new List<Dictionary<string, object>>();
            var browserColumnsByDoc = new Dictionary<Document, List<BrowserOrganizationHelper.BrowserColumn>>();

            foreach (Document doc in uiApp.Application.Documents)
            {
                if (doc.IsLinked || doc.IsFamilyDocument)
                    continue;

                List<View> views = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v =>
                        !(v is ViewSheet) &&
                        !v.IsTemplate &&
                        v.ViewType != ViewType.ProjectBrowser &&
                        v.ViewType != ViewType.SystemBrowser)
                    .ToList();

                List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                    BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, views);
                browserColumnsByDoc[doc] = browserColumns;

                var viewIdToViewport = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .GroupBy(vp => vp.ViewId)
                    .ToDictionary(g => g.Key, g => g.First());

                foreach (View v in views)
                {
                    var dict = new Dictionary<string, object>();
                    dict["Document"] = doc.Title;
                    BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, v, doc, browserColumns);
                    dict["Name"] = v.Name;

                    Viewport viewport;
                    if (viewIdToViewport.TryGetValue(v.Id, out viewport))
                    {
                        ViewSheet containingSheet = doc.GetElement(viewport.SheetId) as ViewSheet;
                        dict["Sheet Number"] = containingSheet != null ? containingSheet.SheetNumber : "";
                        dict["Sheet Title"] = containingSheet != null ? containingSheet.Name : "";
                    }
                    else
                    {
                        dict["Sheet Number"] = "";
                        dict["Sheet Title"] = "";
                    }

                    dict["ElementIdObject"] = v.Id;
                    dict["__OriginalObject"] = v;
                    dict["__Document"] = doc;
                    gridData.Add(dict);
                }
            }

            if (gridData.Count == 0)
            {
                TaskDialog.Show("Info", "No views found in any open documents.");
                return Result.Failed;
            }

            // Sort per document group by browser columns, then combine
            var groupedByDocument = gridData.GroupBy(row => row["Document"]?.ToString() ?? "").ToList();
            gridData.Clear();

            foreach (var docGroup in groupedByDocument.OrderBy(g => g.Key))
            {
                var viewsInDoc = docGroup.ToList();
                Document doc = viewsInDoc.First()["__Document"] as Document;

                if (doc != null && browserColumnsByDoc.TryGetValue(doc, out var bc) &&
                    bc != null && bc.Count > 0)
                    viewsInDoc = BrowserOrganizationHelper.SortByBrowserColumns(viewsInDoc, bc);
                else
                    viewsInDoc = viewsInDoc.OrderBy(row =>
                        (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View v)
                            ? v.Title : "").ToList();

                gridData.AddRange(viewsInDoc);
            }

            columns = new List<string> { "Document" };

            var allBrowserColumnNames = new HashSet<string>();
            foreach (var bc in browserColumnsByDoc.Values)
                if (bc != null)
                    foreach (var col in bc) allBrowserColumnNames.Add(col.Name);
            columns.AddRange(allBrowserColumnNames.OrderBy(n => n));

            columns.Add("Name");
            columns.Add("Sheet Number");
            columns.Add("Sheet Title");

            ViewDataCache.SaveSessionCache(uiApp.Application, "views", gridData, columns);
        }

        // Pre-select active view in active doc
        ElementId targetViewId = !(activeView is ViewSheet) ? activeView.Id : null;
        int selectedIndex = targetViewId != null
            ? gridData.FindIndex(row =>
                row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is View rv &&
                row.ContainsKey("__Document") && row["__Document"] is Document rd &&
                rv.Id == targetViewId && rd.Equals(activeDoc))
            : -1;

        List<int> initialSelectionIndices = selectedIndex >= 0
            ? new List<int> { selectedIndex }
            : new List<int>();

        CustomGUIs.SetCurrentUIDocument(uidoc);
        List<Dictionary<string, object>> selectedRows =
            CustomGUIs.DataGrid(gridData, columns, false, initialSelectionIndices);

        bool editsWereApplied = false;
        if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
        {
            CustomGUIs.ApplyCellEditsToEntities();
            editsWereApplied = true;
            ViewDataCache.InvalidateAll("views");
        }

        if (!editsWereApplied && selectedRows != null && selectedRows.Any())
        {
            Document currentDoc = activeDoc;
            UIDocument currentUidoc = uidoc;

            foreach (var row in selectedRows)
            {
                View view = row["__OriginalObject"] as View;
                Document viewDoc = row["__Document"] as Document;
                if (view == null || viewDoc == null) continue;

                if (!viewDoc.Equals(currentDoc))
                {
                    if (string.IsNullOrEmpty(viewDoc.PathName))
                    {
                        TaskDialog.Show("Unsaved Document",
                            $"Cannot switch to unsaved document '{viewDoc.Title}'.");
                        continue;
                    }
                    try
                    {
                        uiApp.Application.OpenDocumentFile(viewDoc.PathName);
                        currentUidoc = uiApp.OpenAndActivateDocument(viewDoc.PathName);
                        currentDoc = viewDoc;
                    }
                    catch (System.Exception ex)
                    {
                        TaskDialog.Show("Error",
                            $"Failed to switch to document '{viewDoc.Title}':\n\n{ex.Message}");
                        continue;
                    }
                }

                currentUidoc.RequestViewChange(view);
            }
        }

        return Result.Succeeded;
    }
}
