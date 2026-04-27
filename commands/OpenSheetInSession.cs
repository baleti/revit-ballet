using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("")]
public class OpenSheetInSession : IExternalCommand
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

        if (ViewDataCache.TryGetSessionCache(uiApp.Application, "sheets", out gridData, out columns))
        {
            // Cache hit
        }
        else
        {
            gridData = new List<Dictionary<string, object>>();
            var browserColumnsByDoc = new Dictionary<Document, List<BrowserOrganizationHelper.BrowserColumn>>();

            foreach (Document doc in uiApp.Application.Documents)
            {
                if (doc.IsLinked || doc.IsFamilyDocument)
                    continue;

                List<ViewSheet> sheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .ToList();

                List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                    BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, sheets.Cast<View>().ToList());
                browserColumnsByDoc[doc] = browserColumns;

                foreach (ViewSheet sheet in sheets)
                {
                    var dict = new Dictionary<string, object>();
                    dict["Document"] = doc.Title;
                    BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, sheet, doc, browserColumns);
                    dict["SheetNumber"] = sheet.SheetNumber;
                    dict["Name"] = sheet.Name;
                    dict["ElementIdObject"] = sheet.Id;
                    dict["__OriginalObject"] = sheet;
                    dict["__Document"] = doc;
                    gridData.Add(dict);
                }
            }

            if (gridData.Count == 0)
            {
                TaskDialog.Show("Info", "No sheets found in any open documents.");
                return Result.Failed;
            }

            var groupedByDocument = gridData.GroupBy(row => row["Document"]?.ToString() ?? "").ToList();
            gridData.Clear();

            foreach (var docGroup in groupedByDocument.OrderBy(g => g.Key))
            {
                var sheetsInDoc = docGroup.ToList();
                Document doc = sheetsInDoc.First()["__Document"] as Document;

                if (doc != null && browserColumnsByDoc.TryGetValue(doc, out var bc) &&
                    bc != null && bc.Count > 0)
                    sheetsInDoc = BrowserOrganizationHelper.SortByBrowserColumns(sheetsInDoc, bc, tiebreakerColumn: "SheetNumber");
                else
                    sheetsInDoc = sheetsInDoc.OrderBy(row => row["SheetNumber"]?.ToString() ?? "").ToList();

                gridData.AddRange(sheetsInDoc);
            }

            columns = new List<string> { "Document" };

            var allBrowserColumnNames = new HashSet<string>();
            foreach (var bc in browserColumnsByDoc.Values)
                if (bc != null)
                    foreach (var col in bc) allBrowserColumnNames.Add(col.Name);
            columns.AddRange(allBrowserColumnNames.OrderBy(n => n));

            columns.Add("SheetNumber");
            columns.Add("Name");

            ViewDataCache.SaveSessionCache(uiApp.Application, "sheets", gridData, columns);
        }

        int selectedIndex = -1;
        ElementId targetSheetId = null;

        if (activeView is ViewSheet)
        {
            targetSheetId = activeView.Id;
        }
        else
        {
            Viewport vp = new FilteredElementCollector(activeDoc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .FirstOrDefault(vpt => vpt.ViewId == activeView.Id);
            if (vp != null)
            {
                ViewSheet containingSheet = activeDoc.GetElement(vp.SheetId) as ViewSheet;
                if (containingSheet != null)
                    targetSheetId = containingSheet.Id;
            }
        }

        if (targetSheetId != null)
        {
            selectedIndex = gridData.FindIndex(row =>
                row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is ViewSheet sheet &&
                row.ContainsKey("__Document") && row["__Document"] is Document doc &&
                sheet.Id == targetSheetId && doc.Equals(activeDoc));
        }

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
            ViewDataCache.InvalidateAll("sheets");
        }

        if (!editsWereApplied && selectedRows != null && selectedRows.Any())
        {
            Document currentDoc = activeDoc;
            UIDocument currentUidoc = uidoc;

            foreach (var row in selectedRows)
            {
                ViewSheet sheet = row["__OriginalObject"] as ViewSheet;
                Document sheetDoc = row["__Document"] as Document;
                if (sheet == null || sheetDoc == null) continue;

                if (!sheetDoc.Equals(currentDoc))
                {
                    if (string.IsNullOrEmpty(sheetDoc.PathName))
                    {
                        TaskDialog.Show("Unsaved Document",
                            $"Cannot switch to unsaved document '{sheetDoc.Title}'.");
                        continue;
                    }
                    try
                    {
                        uiApp.Application.OpenDocumentFile(sheetDoc.PathName);
                        currentUidoc = uiApp.OpenAndActivateDocument(sheetDoc.PathName);
                        currentDoc = sheetDoc;
                    }
                    catch (System.Exception ex)
                    {
                        TaskDialog.Show("Error",
                            $"Failed to switch to document '{sheetDoc.Title}':\n\n{ex.Message}");
                        continue;
                    }
                }

                currentUidoc.RequestViewChange(sheet);
            }
        }

        return Result.Succeeded;
    }
}
