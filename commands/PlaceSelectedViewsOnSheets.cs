// PlaceSelectedViewsOnSheets.cs – Revit 2024, C# 7.3
// Author: ChatGPT (o3)
// Natural‑sort enabled (e.g. "1", "1A", "10" → 1, 1A, 10)

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using DB = Autodesk.Revit.DB;
using WinForms = System.Windows.Forms;

namespace RevitAddin.Commands
{
    /// <summary>
    /// Prompts the user to map the currently selected *views* (not sheets) to chosen sheets
    /// and places each view at the centre of the title‑block on its target sheet, without
    /// opening those sheets in the UI.
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
                // 1. Collect *non‑sheet* views currently selected               
                // ────────────────────────────────────────────────────────────────
                IList<DB.View> selViews = uidoc.GetSelectionIds()
                    .Select(id => doc.GetElement(id) as DB.View)
                    .Where(v => v != null &&
                                 !(v is DB.ViewSheet) &&
                                 !v.IsTemplate &&
                                 v.ViewType != DB.ViewType.Schedule &&
                                 v.ViewType != DB.ViewType.Legend)
                    .ToList();
                if (selViews.Count == 0)
                {
                    TaskDialog.Show("Place Views on Sheets",
                        "Please select one or more *non‑sheet* views in the Project Browser before running this command.");
                    return Result.Cancelled;
                }

                // ────────────────────────────────────────────────────────────────
                // 2. Pick sheets via CustomGUIs.DataGrid                         
                // ────────────────────────────────────────────────────────────────
                IList<DB.ViewSheet> allSheets = new DB.FilteredElementCollector(doc)
                                                   .OfClass(typeof(DB.ViewSheet))
                                                   .Cast<DB.ViewSheet>()
                                                   .OrderBy(s => s.SheetNumber, NaturalSortComparer.Instance)
                                                   .ToList();

                var sheetRows = allSheets.Select(s => new Dictionary<string, object>
                {
                    {"Number", s.SheetNumber},
                    {"Name",   s.Name}
                }).ToList();
                var columns = new List<string> { "Number", "Name" };

                var selectedRows = CustomGUIs.DataGrid(sheetRows, columns, false);
                if (selectedRows == null || selectedRows.Count == 0)
                    return Result.Cancelled;

                IList<DB.ViewSheet> chosenSheets = selectedRows
                    .Select(r => allSheets.First(s => s.SheetNumber == r["Number"].ToString()))
                    .ToList();

                // ────────────────────────────────────────────────────────────────
                // 3. Sort views & sheets in natural order, pair 1‑to‑1          
                // ────────────────────────────────────────────────────────────────
                var sortedViews  = selViews.OrderBy(v => v.Name, NaturalSortComparer.Instance).ToList();
                var sortedSheets = chosenSheets.OrderBy(s => s.SheetNumber, NaturalSortComparer.Instance).ToList();

                using (var mappingForm = new MappingForm(sortedViews, sortedSheets))
                {
                    if (mappingForm.ShowDialog() != WinForms.DialogResult.OK)
                        return Result.Cancelled;

                    var map = mappingForm.Mappings;
                    if (map.Count == 0)
                        return Result.Cancelled;

                    // Remember current active view so we can restore it afterwards.
                    DB.View originalActive = uidoc.ActiveView;

                    // ────────────────────────────────────────────────────────────
                    // 4. Place the views silently (no sheet activation)          
                    // ────────────────────────────────────────────────────────────
                    using (var tx = new DB.Transaction(doc, "Place Selected Views on Sheets"))
                    {
                        tx.Start();

                        foreach (var (view, sheet) in map)
                        {
                            if (view == null || sheet == null) continue;
                            if (IsViewPlaced(doc, view))       continue;

                            DB.XYZ pt = GetTitleBlockCenter(doc, sheet) ?? SheetCentre(sheet);
                            DB.Viewport.Create(doc, sheet.Id, view.Id, pt);
                        }

                        tx.Commit();
                    }

                    // Restore the user's original view (prevents UI from jumping).
                    if (originalActive != null && uidoc.ActiveView.Id != originalActive.Id)
                        uidoc.ActiveView = originalActive;
                }

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

    // ──────────────────────────────────────────────────────────────────────
    // Mapping dialog – lets user tweak view→sheet pairs; supports Enter/Esc
    // and auto‑sized height to fit grid rows (unless > 90 % screen height)
    // ──────────────────────────────────────────────────────────────────────
    internal class MappingForm : WinForms.Form
    {
        private readonly WinForms.DataGridView _grid = new WinForms.DataGridView
        {
            Dock = WinForms.DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            RowHeadersVisible = false,
            AutoSizeRowsMode = WinForms.DataGridViewAutoSizeRowsMode.AllCells
        };

        private readonly WinForms.Button _ok     = new WinForms.Button { Text = "OK",     Dock = WinForms.DockStyle.Right, DialogResult = WinForms.DialogResult.OK };
        private readonly WinForms.Button _cancel = new WinForms.Button { Text = "Cancel", Dock = WinForms.DockStyle.Right, DialogResult = WinForms.DialogResult.Cancel };

        private readonly IList<DB.View> _views;
        private readonly IList<DB.ViewSheet> _sheets;
        public IList<(DB.View View, DB.ViewSheet Sheet)> Mappings { get; } = new List<(DB.View, DB.ViewSheet)>();

        public MappingForm(IList<DB.View> views, IList<DB.ViewSheet> sheets)
        {
            KeyPreview = true; // capture Enter/Esc before controls
            _views  = views;
            _sheets = sheets;

            Text  = "Map Views → Sheets";
            Width = 600;

            BuildGrid();
            AutoSizeAndLayout();
        }

        private void BuildGrid()
        {
            _grid.Columns.Add(new WinForms.DataGridViewTextBoxColumn
            {
                HeaderText   = "View",
                ReadOnly     = true,
                AutoSizeMode = WinForms.DataGridViewAutoSizeColumnMode.Fill
            });

            var sheetColumn = new WinForms.DataGridViewComboBoxColumn
            {
                HeaderText   = "Sheet",
                AutoSizeMode = WinForms.DataGridViewAutoSizeColumnMode.Fill,
                FlatStyle    = WinForms.FlatStyle.Flat,
                DropDownWidth = 250
            };
            sheetColumn.Items.AddRange(_sheets.Select(SheetDisplay).Cast<object>().ToArray());
            _grid.Columns.Add(sheetColumn);

            int pairCount = Math.Min(_views.Count, _sheets.Count);
            for (int i = 0; i < pairCount; ++i)
                _grid.Rows.Add(_views[i].Name, SheetDisplay(_sheets[i]));
            for (int i = pairCount; i < _views.Count; ++i)
                _grid.Rows.Add(_views[i].Name, null);
        }

        private void AutoSizeAndLayout()
        {
            var buttons = new WinForms.Panel { Dock = WinForms.DockStyle.Bottom, Height = 40 };
            buttons.Controls.AddRange(new WinForms.Control[] { _cancel, _ok });
            Controls.AddRange(new WinForms.Control[] { _grid, buttons });

            int desiredGridHeight = _grid.ColumnHeadersHeight + _grid.RowTemplate.Height * _grid.Rows.Count + 2;
            int desiredFormHeight = desiredGridHeight + buttons.Height + 50; // padding
            int screenHeight      = WinForms.Screen.PrimaryScreen.WorkingArea.Height;
            Height = desiredFormHeight < screenHeight ? desiredFormHeight : (int)(screenHeight * 0.9);
        }

        protected override bool ProcessCmdKey(ref WinForms.Message msg, WinForms.Keys keyData)
        {
            if (keyData == WinForms.Keys.Enter)
            {
                CommitAndClose(WinForms.DialogResult.OK);
                return true;
            }
            if (keyData == WinForms.Keys.Escape)
            {
                CommitAndClose(WinForms.DialogResult.Cancel);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void CommitAndClose(WinForms.DialogResult result)
        {
            if (result == WinForms.DialogResult.OK)
                CollectMappings();

            DialogResult = result;
            Close();
        }

        private void CollectMappings()
        {
            Mappings.Clear();
            var usedSheets = new HashSet<DB.ElementId>();
            for (int i = 0; i < _grid.Rows.Count; ++i)
            {
                string viewName  = _grid.Rows[i].Cells[0].Value as string;
                string sheetDisp = _grid.Rows[i].Cells[1].Value as string;
                if (string.IsNullOrEmpty(viewName) || string.IsNullOrEmpty(sheetDisp))
                    continue;

                DB.View      view  = _views.FirstOrDefault(x => x.Name == viewName);
                DB.ViewSheet sheet = _sheets.FirstOrDefault(x => SheetDisplay(x) == sheetDisp);
                if (view == null || sheet == null)
                    continue;

                // If this sheet is already assigned earlier, drop the previous occurrence
                if (usedSheets.Contains(sheet.Id))
                {
                    for (int j = 0; j < Mappings.Count; ++j)
                    {
                        if (Mappings[j].Sheet.Id == sheet.Id)
                        {
                            Mappings.RemoveAt(j);
                            break;
                        }
                    }
                }

                usedSheets.Add(sheet.Id);
                Mappings.Add((view, sheet));
            }
        }

        private static string SheetDisplay(DB.ViewSheet s) => $"{s.SheetNumber} – {s.Name}";
    }
}
