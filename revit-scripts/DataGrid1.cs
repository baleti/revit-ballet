using Autodesk.Revit.UI;            // for TaskDialog
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

public partial class CustomGUIs
{
    public static List<T> DataGrid<T>(List<T> entries,
                                      List<string> propertyNames,
                                      List<int> initialSelectionIndices = null,
                                      string Title = null)
    {
        // ─────────────────────────────────────
        // 1. Early-out: nothing to display
        // ─────────────────────────────────────
        if (entries == null || entries.Count == 0)
        {
            TaskDialog.Show(Title ?? "Selection", "There were 0 elements.");
            return new List<T>();
        }

        var filteredEntries = new List<T>();

        // ─────────────────────────────────────
        // 2. Build form & grid
        // ─────────────────────────────────────
        var form = new Form { StartPosition = FormStartPosition.CenterScreen };
        if (!string.IsNullOrEmpty(Title)) form.Text = Title;

        var dataGridView = new DataGridView
        {
            Dock = DockStyle.Fill,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
            ReadOnly = true,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false
        };

        foreach (var propertyName in propertyNames)
        {
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText   = propertyName,
                DataPropertyName = propertyName,
                SortMode     = DataGridViewColumnSortMode.Programmatic
            });
        }

        // Sorting helpers
        string    lastSortedProperty = null;
        SortOrder lastSortOrder      = SortOrder.None;
        var _entries = new List<T>(entries);

        void ApplySort()
        {
            if (!string.IsNullOrEmpty(lastSortedProperty) && _entries.Count > 0)
            {
                _entries = (lastSortOrder == SortOrder.Ascending)
                           ? _entries.OrderBy(o => o.GetType().GetProperty(lastSortedProperty)?.GetValue(o, null)).ToList()
                           : _entries.OrderByDescending(o => o.GetType().GetProperty(lastSortedProperty)?.GetValue(o, null)).ToList();
            }
        }

        void RebindData()
        {
            dataGridView.DataSource = null;
            dataGridView.DataSource = _entries;
        }

        RebindData();
        dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        // ── Column-header click to sort ─────────────────────────
        dataGridView.ColumnHeaderMouseClick += (s, e) =>
        {
            var col  = dataGridView.Columns[e.ColumnIndex];
            var prop = col.DataPropertyName;
            if (string.IsNullOrEmpty(prop)) return;

            lastSortOrder      = (prop == lastSortedProperty && lastSortOrder == SortOrder.Ascending)
                                 ? SortOrder.Descending
                                 : SortOrder.Ascending;
            lastSortedProperty = prop;

            foreach (DataGridViewColumn c in dataGridView.Columns)
                c.HeaderCell.SortGlyphDirection = SortOrder.None;
            col.HeaderCell.SortGlyphDirection = lastSortOrder;

            ApplySort();
            RebindData();
        };

        // ── Maintain initial selection ─────────────────────────
        dataGridView.DataBindingComplete += (s, e) =>
        {
            dataGridView.ClearSelection();
            if (initialSelectionIndices != null)
            {
                foreach (var i in initialSelectionIndices.Where(i => i >= 0 && i < dataGridView.Rows.Count))
                {
                    dataGridView.Rows[i].Selected = true;
                    dataGridView.CurrentCell      = dataGridView.Rows[i].Cells[0];
                }
            }
        };

        // ─────────────────────────────────────
        // 3. Search box (qualified to avoid clash)
        // ─────────────────────────────────────
        var searchBox = new System.Windows.Forms.TextBox { Dock = DockStyle.Top };
        searchBox.TextChanged += (s, e) =>
        {
            string searchText = searchBox.Text.ToLower();
            var clauses = searchText.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(cl => cl.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

            _entries = entries.Where(entry =>
            {
                foreach (var orBlock in clauses)
                {
                    bool matchesOr = true;
                    foreach (var term in orBlock)
                    {
                        bool   neg    = term.StartsWith("!");
                        string actual = neg ? term.Substring(1) : term;   // ← C# 7-friendly

                        bool found = propertyNames.Any(p =>
                        {
                            var val = entry.GetType().GetProperty(p)?.GetValue(entry, null)?.ToString();
                            return val != null && val.IndexOf(actual, StringComparison.OrdinalIgnoreCase) >= 0;
                        });

                        if (neg ? found : !found)
                        {
                            matchesOr = false;
                            break;
                        }
                    }
                    if (matchesOr) return true;
                }
                return false;
            }).ToList();

            ApplySort();
            RebindData();
        };

        // ── Double-click / key shortcuts ───────────────────────
        dataGridView.CellDoubleClick += (s, e) =>
        {
            if (e.RowIndex < 0) return;
            foreach (DataGridViewRow row in dataGridView.SelectedRows)
                filteredEntries.Add((T)row.DataBoundItem);
            form.Close();
        };

        dataGridView.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Escape) form.Close();
            else if (e.KeyCode == Keys.Enter)
            {
                foreach (DataGridViewRow row in dataGridView.SelectedRows)
                    filteredEntries.Add((T)row.DataBoundItem);
                form.Close();
            }
            else if (e.KeyCode == Keys.Tab)
                searchBox.Focus();
        };

        searchBox.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Escape) form.Close();
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
            {
                dataGridView.Focus();
                if (dataGridView.CurrentRow == null && dataGridView.Rows.Count > 0)
                {
                    dataGridView.CurrentCell = dataGridView.Rows[0].Cells[0];
                    dataGridView.Rows[0].Selected = true;
                }
            }
            else if (e.KeyCode == Keys.Enter && dataGridView.Rows.Count > 0)
            {
                if (dataGridView.SelectedRows.Count == 0)
                {
                    dataGridView.Rows[0].Selected = true;
                    dataGridView.CurrentCell      = dataGridView.Rows[0].Cells[0];
                }
                foreach (DataGridViewRow row in dataGridView.SelectedRows)
                    filteredEntries.Add((T)row.DataBoundItem);
                form.Close();
            }
            else if (e.KeyCode == Keys.Space && string.IsNullOrWhiteSpace(searchBox.Text))
            {
                int count = dataGridView.Rows.Count;
                if (count == 0) return;

                int cur  = dataGridView.CurrentRow?.Index ?? -1;
                int next = e.Shift ? (cur - 1 + count) % count : (cur + 1) % count;

                dataGridView.ClearSelection();
                dataGridView.CurrentCell           = dataGridView.Rows[next].Cells[0];
                dataGridView.Rows[next].Selected   = true;

                filteredEntries.Add((T)dataGridView.Rows[next].DataBoundItem);
                form.Close();
            }
        };

        // ─────────────────────────────────────
        // 4. Defensive sizing – no Rows[0] assumption
        // ─────────────────────────────────────
        form.Load += (s, e) =>
        {
            int oneRowHeight = dataGridView.RowTemplate.Height;
            if (dataGridView.Rows.Count > 0)
                oneRowHeight = dataGridView.Rows[0].Height;

            int totalRows  = dataGridView.Rows.GetRowsHeight(DataGridViewElementStates.Visible);
            int heightNeed = totalRows
                             + dataGridView.ColumnHeadersHeight
                             + 2 * oneRowHeight
                             + SystemInformation.HorizontalScrollBarHeight;

            form.Height = heightNeed;

            if (heightNeed > Screen.PrimaryScreen.WorkingArea.Height)
            {
                form.StartPosition = FormStartPosition.Manual;
                form.Top           = 10;
                form.Height        = Screen.PrimaryScreen.WorkingArea.Height - 10;
            }
            else if (heightNeed > Screen.PrimaryScreen.WorkingArea.Height / 2)
            {
                form.StartPosition = FormStartPosition.Manual;
                form.Top           = 10;
            }

            int formMaxWidth = 0;
            foreach (DataGridViewColumn col in dataGridView.Columns)
            {
                int headerW = TextRenderer.MeasureText(col.HeaderText,
                                                        dataGridView.ColumnHeadersDefaultCellStyle.Font).Width;
                int cellW   = 0;
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    var v = row.Cells[col.Index].Value?.ToString();
                    if (v != null)
                        cellW = Math.Max(cellW,
                                         TextRenderer.MeasureText(v, dataGridView.DefaultCellStyle.Font).Width);
                }
                col.Width    = Math.Max(headerW, cellW) + 10;
                formMaxWidth += col.Width;
            }
            form.Width = formMaxWidth + SystemInformation.VerticalScrollBarWidth + 59;
        };

        // ─────────────────────────────────────
        // 5. Show the dialog
        // ─────────────────────────────────────
        form.Controls.Add(dataGridView);
        form.Controls.Add(searchBox);
        searchBox.Select();

        form.ShowDialog();
        return filteredEntries;
    }
}
