// PlaceSelectedViewsOnSheets.cs – Revit 2024, C# 7.3
// Refactored to use DataGrid with editable Sheets and Views columns
// Natural‑sort enabled (e.g. "1", "1A", "10" → 1, 1A, 10)

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using DB = Autodesk.Revit.DB;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitAddin.Commands
{
    /// <summary>
    /// Prompts the user to map views to sheets using an editable DataGrid with two columns.
    /// Validates that mappings are 1:1 (except for legends) and that all names exist.
    /// Places each view at the centre of the title‑block on its target sheet.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PlaceSelectedViewsOnSheets : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, DB.ElementSet _) // elements not used
        {
            UIApplication uiapp = commandData.Application;
            UIDocument    uidoc = uiapp.ActiveUIDocument;
            DB.Document   doc   = uidoc.Document;

            try
            {
                // ────────────────────────────────────────────────────────────────
                // 1. Collect views and sheets from current selection
                // ────────────────────────────────────────────────────────────────
                var selViews = new List<DB.View>();
                var selSheets = new List<DB.ViewSheet>();

                foreach (var id in uidoc.GetSelectionIds())
                {
                    var elem = doc.GetElement(id);

                    // Collect views (non-sheet, non-template, non-schedule)
                    if (elem is DB.View view &&
                        !(view is DB.ViewSheet) &&
                        !view.IsTemplate &&
                        view.ViewType != DB.ViewType.Schedule)
                    {
                        selViews.Add(view);
                    }
                    // Collect sheets
                    else if (elem is DB.ViewSheet sheet)
                    {
                        selSheets.Add(sheet);
                    }
                    // Collect views from viewports
                    else if (elem is DB.Viewport viewport)
                    {
                        var vpView = doc.GetElement(viewport.ViewId) as DB.View;
                        if (vpView != null &&
                            !vpView.IsTemplate &&
                            vpView.ViewType != DB.ViewType.Schedule)
                        {
                            selViews.Add(vpView);
                        }
                    }
                }

                // Require at least one view or sheet selected
                if (selViews.Count == 0 && selSheets.Count == 0)
                {
                    TaskDialog.Show("Place Views on Sheets",
                        "Please select views and/or sheets in the Project Browser.\n\n" +
                        "The dialog will let you map the selected views to the selected sheets.");
                    return Result.Cancelled;
                }

                // ────────────────────────────────────────────────────────────────
                // 2. Get all sheets and views for validation
                // ────────────────────────────────────────────────────────────────
                var allSheets = new DB.FilteredElementCollector(doc)
                    .OfClass(typeof(DB.ViewSheet))
                    .Cast<DB.ViewSheet>()
                    .OrderBy(s => s.SheetNumber, NaturalSortComparer.Instance)
                    .ToList();

                var allViews = new DB.FilteredElementCollector(doc)
                    .OfClass(typeof(DB.View))
                    .Cast<DB.View>()
                    .Where(v => !(v is DB.ViewSheet) && !v.IsTemplate && v.ViewType != DB.ViewType.Schedule)
                    .OrderBy(v => v.Name, NaturalSortComparer.Instance)
                    .ToList();

                // ────────────────────────────────────────────────────────────────
                // 3. Create mapping grid with editable Sheets and Views columns
                // ────────────────────────────────────────────────────────────────
                var sortedViews  = selViews.OrderBy(v => v.Name, NaturalSortComparer.Instance).ToList();
                var sortedSheets = selSheets.OrderBy(s => s.SheetNumber, NaturalSortComparer.Instance).ToList();

                // Validation loop
                List<Dictionary<string, object>> mappingRows = null;
                string validationError = null;

                do
                {
                    // Populate or repopulate the grid
                    if (mappingRows == null)
                    {
                        // First time: create initial mappings
                        mappingRows = new List<Dictionary<string, object>>();
                        int count = Math.Max(sortedViews.Count, sortedSheets.Count);

                        for (int i = 0; i < count; i++)
                        {
                            mappingRows.Add(new Dictionary<string, object>
                            {
                                {"Sheets", i < sortedSheets.Count ? sortedSheets[i].SheetNumber + " - " + sortedSheets[i].Name : ""},
                                {"Views",  i < sortedViews.Count ? sortedViews[i].Name : ""}
                            });
                        }
                    }

                    // Show validation error if exists
                    if (!string.IsNullOrEmpty(validationError))
                    {
                        TaskDialog.Show("Validation Error", validationError);
                        validationError = null;
                    }

                    // Show DataGrid - returnAllEntries=true ensures we get all rows regardless of selection
                    var columns = new List<string> { "Sheets", "Views" };
                    var result = CustomGUIs.DataGrid(mappingRows, columns, spanAllScreens: false,
                        initialSelectionIndices: null, onDeleteEntries: null,
                        allowCreateFromSearch: false, commandName: null, returnAllEntries: true);

                    if (result == null || result.Count == 0)
                        return Result.Cancelled;

                    // Update mappingRows with user edits (DataGrid returns all entries)
                    mappingRows = result;

                    // ────────────────────────────────────────────────────────────
                    // 4. Validate mappings
                    // ────────────────────────────────────────────────────────────
                    var mappings = new List<(DB.View view, DB.ViewSheet sheet)>();
                    var sheetCounts = new Dictionary<string, int>();
                    var viewCounts = new Dictionary<string, int>();
                    var missingSheets = new List<string>();
                    var missingViews = new List<string>();
                    var legends = new HashSet<string>();

                    foreach (var row in mappingRows)
                    {
                        string sheetText = row.ContainsKey("Sheets") && row["Sheets"] != null ? row["Sheets"].ToString().Trim() : "";
                        string viewText = row.ContainsKey("Views") && row["Views"] != null ? row["Views"].ToString().Trim() : "";

                        // Skip rows that don't have both values (user deleted one or both)
                        if (string.IsNullOrEmpty(sheetText) || string.IsNullOrEmpty(viewText))
                            continue;

                        // Count occurrences for duplicate detection
                        if (!sheetCounts.ContainsKey(sheetText))
                            sheetCounts[sheetText] = 0;
                        sheetCounts[sheetText]++;

                        if (!viewCounts.ContainsKey(viewText))
                            viewCounts[viewText] = 0;
                        viewCounts[viewText]++;
                    }

                    // Build actual mappings and check for existence
                    foreach (var row in mappingRows)
                    {
                        string sheetText = row.ContainsKey("Sheets") && row["Sheets"] != null ? row["Sheets"].ToString().Trim() : "";
                        string viewText = row.ContainsKey("Views") && row["Views"] != null ? row["Views"].ToString().Trim() : "";

                        // Skip rows that don't have both sheet and view (user deleted one or both)
                        if (string.IsNullOrEmpty(sheetText) || string.IsNullOrEmpty(viewText))
                            continue;

                        // Find sheet (match by number or number+name)
                        DB.ViewSheet sheet = allSheets.FirstOrDefault(s =>
                            sheetText.Equals(s.SheetNumber, StringComparison.OrdinalIgnoreCase) ||
                            sheetText.Equals(s.SheetNumber + " - " + s.Name, StringComparison.OrdinalIgnoreCase) ||
                            sheetText.StartsWith(s.SheetNumber + " - ", StringComparison.OrdinalIgnoreCase));

                        // Find view
                        DB.View view = allViews.FirstOrDefault(v =>
                            v.Name.Equals(viewText, StringComparison.OrdinalIgnoreCase));

                        // Track missing elements
                        if (sheet == null && !missingSheets.Contains(sheetText))
                            missingSheets.Add(sheetText);

                        if (view == null && !missingViews.Contains(viewText))
                            missingViews.Add(viewText);

                        // Track legends (can be placed on multiple sheets)
                        if (view != null && view.ViewType == DB.ViewType.Legend)
                            legends.Add(viewText);

                        if (view != null && sheet != null)
                            mappings.Add((view, sheet));
                    }

                    // Check for validation errors
                    if (!string.IsNullOrEmpty(validationError))
                        continue;

                    if (missingSheets.Count > 0 || missingViews.Count > 0)
                    {
                        var errorMsg = "The following items do not exist:\n\n";
                        if (missingSheets.Count > 0)
                            errorMsg += "Sheets:\n  " + string.Join("\n  ", missingSheets) + "\n\n";
                        if (missingViews.Count > 0)
                            errorMsg += "Views:\n  " + string.Join("\n  ", missingViews) + "\n";
                        validationError = errorMsg;
                        continue;
                    }

                    // Check for duplicate mappings (sheets and non-legend views must be 1:1)
                    var duplicateSheets = sheetCounts.Where(kvp => kvp.Value > 1).Select(kvp => kvp.Key).ToList();
                    var duplicateViews = viewCounts.Where(kvp => kvp.Value > 1 && !legends.Contains(kvp.Key)).Select(kvp => kvp.Key).ToList();

                    if (duplicateSheets.Count > 0 || duplicateViews.Count > 0)
                    {
                        var errorMsg = "The following items are mapped multiple times (only legends can be placed on multiple sheets):\n\n";
                        if (duplicateSheets.Count > 0)
                            errorMsg += "Sheets:\n  " + string.Join("\n  ", duplicateSheets) + "\n\n";
                        if (duplicateViews.Count > 0)
                            errorMsg += "Views:\n  " + string.Join("\n  ", duplicateViews) + "\n";
                        validationError = errorMsg;
                        continue;
                    }

                    // Validation passed - if no mappings, just return (user deleted everything)
                    if (mappings.Count == 0)
                        return Result.Cancelled;

                    // ────────────────────────────────────────────────────────────
                    // 5. Place the views silently (no sheet activation)
                    // ────────────────────────────────────────────────────────────
                    DB.View originalActive = uidoc.ActiveView;

                    using (var tx = new DB.Transaction(doc, "Place Views on Sheets"))
                    {
                        tx.Start();

                        foreach (var (view, sheet) in mappings)
                        {
                            if (view == null || sheet == null)
                                continue;

                            // Legends can be placed multiple times, other views cannot
                            if (view.ViewType != DB.ViewType.Legend && IsViewPlaced(doc, view))
                                continue;

                            DB.XYZ pt = GetTitleBlockCenter(doc, sheet) ?? SheetCentre(sheet);
                            DB.Viewport.Create(doc, sheet.Id, view.Id, pt);
                        }

                        tx.Commit();
                    }

                    // Restore the user's original view (prevents UI from jumping).
                    if (originalActive != null && uidoc.ActiveView.Id != originalActive.Id)
                        uidoc.ActiveView = originalActive;

                    return Result.Succeeded;

                } while (!string.IsNullOrEmpty(validationError));

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ──────────────────────────────────────────────────────────────────
        // Helper utilities
        // ──────────────────────────────────────────────────────────────────
        private static bool IsViewPlaced(DB.Document d, DB.View v)
        {
            return new DB.FilteredElementCollector(d)
                .OfClass(typeof(DB.Viewport))
                .Cast<DB.Viewport>()
                .Any(vp => vp.ViewId == v.Id);
        }

        private static DB.XYZ GetTitleBlockCenter(DB.Document d, DB.ViewSheet sh)
        {
            var tb = new DB.FilteredElementCollector(d)
                        .OwnedByView(sh.Id)
                        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
                        .Cast<DB.FamilyInstance>()
                        .FirstOrDefault();
            if (tb == null) return null;
            DB.BoundingBoxXYZ bb = tb.get_BoundingBox(sh);
            return (bb.Min + bb.Max) * 0.5;
        }

        private static DB.XYZ SheetCentre(DB.ViewSheet sh)
        {
            return new DB.XYZ(
                (sh.Outline.Max.U + sh.Outline.Min.U) / 2,
                (sh.Outline.Max.V + sh.Outline.Min.V) / 2,
                0);
        }
    }

    // ──────────────────────────────────────────────────────────────────────
    // Natural‑sort comparer – splits digit/non‑digit runs and compares
    // numerically where possible
    // ──────────────────────────────────────────────────────────────────────
    internal class NaturalSortComparer : IComparer<string>
    {
        public static readonly NaturalSortComparer Instance = new NaturalSortComparer();

        public int Compare(string a, string b)
        {
            if (ReferenceEquals(a, b)) return 0;
            if (a == null) return -1;
            if (b == null) return 1;

            int ia = 0, ib = 0;
            while (ia < a.Length && ib < b.Length)
            {
                bool da = char.IsDigit(a[ia]);
                bool db = char.IsDigit(b[ib]);

                if (da && db)
                {
                    long va = 0;
                    while (ia < a.Length && char.IsDigit(a[ia]))
                        va = va * 10 + (a[ia++] - '0');

                    long vb = 0;
                    while (ib < b.Length && char.IsDigit(b[ib]))
                        vb = vb * 10 + (b[ib++] - '0');

                    int cmp = va.CompareTo(vb);
                    if (cmp != 0) return cmp;
                }
                else
                {
                    char ca = char.ToUpperInvariant(a[ia++]);
                    char cb = char.ToUpperInvariant(b[ib++]);
                    int cmp = ca.CompareTo(cb);
                    if (cmp != 0) return cmp;
                }
            }
            return a.Length - b.Length;
        }
    }
}
