using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.ComponentModel;

public partial class CustomGUIs
{
    /// <summary>
    /// Shows a filterable / sortable read-only grid and returns the
    /// rows the user double-clicked or pressed Enter on.
    /// </summary>
    public static List<Dictionary<string, object>> DataGrid(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        bool spanAllScreens,
        List<int> initialSelectionIndices = null)
    {
        // ---------------- state ----------------
        List<Dictionary<string, object>> selectedEntries = new List<Dictionary<string, object>>();
        List<Dictionary<string, object>> workingSet = new List<Dictionary<string, object>>(entries);
        List<SortCriteria> sortCriteria = new List<SortCriteria>();

        // ---------------- UI shell -------------
        Form form = new Form();
        form.StartPosition = FormStartPosition.CenterScreen;
        form.Text = "Total Entries: " + entries.Count;
        form.BackColor = Color.White;

        DataGridView grid = new DataGridView();
        grid.Dock = DockStyle.Fill;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.MultiSelect = true;
        grid.ReadOnly = true;
        grid.AutoGenerateColumns = false;
        grid.RowHeadersVisible = false;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToResizeRows = false;
        grid.BackgroundColor = Color.White;
        grid.RowTemplate.Height = 25;
        
        // Disable built-in sorting to prevent the exception
        grid.SortCompare += (sender, e) =>
        {
            e.Handled = true; // Prevent default sorting
            e.SortResult = naturalComparer.Compare(e.CellValue1, e.CellValue2);
        };

        foreach (string col in propertyNames)
        {
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = col,
                DataPropertyName = col,
                SortMode = DataGridViewColumnSortMode.Programmatic // Use custom sorting
            };
            grid.Columns.Add(column);
        }

        TextBox searchBox = new TextBox { Dock = DockStyle.Top };

        // ---------------- helpers --------------
        void RefreshGrid(IEnumerable<Dictionary<string, object>> data)
        {
            grid.Rows.Clear();
            foreach (var rowSrc in data)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(grid);
                for (int c = 0; c < propertyNames.Count; c++)
                {
                    object val;
                    row.Cells[c].Value = rowSrc.TryGetValue(propertyNames[c], out val) ? val : null;
                }
                grid.Rows.Add(row);
            }
        }

        int GetFirstVisibleColumnIndex()
        {
            foreach (DataGridViewColumn c in grid.Columns)
                if (c.Visible) return c.Index;
            return -1;
        }

        // ---------------- filtering ------------
        void UpdateFilteredGrid()
        {
            var filteredData = ApplyFilters(entries, propertyNames, searchBox.Text, grid);
            
            // Apply sorting
            workingSet = filteredData;
            if (sortCriteria.Count > 0)
            {
                workingSet = ApplySorting(workingSet, sortCriteria);
            }

            RefreshGrid(workingSet);
            grid.AutoResizeColumns();
            int reqWidth = grid.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                          + SystemInformation.VerticalScrollBarWidth + 50;
            form.Width = Math.Min(reqWidth, Screen.PrimaryScreen.WorkingArea.Width - 20);
        }

        // first draw
        RefreshGrid(workingSet);

        // initial selection
        if (initialSelectionIndices != null && initialSelectionIndices.Count > 0)
        {
            int firstVisible = GetFirstVisibleColumnIndex();
            foreach (int idx in initialSelectionIndices)
                if (idx >= 0 && idx < grid.Rows.Count && firstVisible >= 0)
                {
                    grid.Rows[idx].Selected = true;
                    grid.CurrentCell = grid.Rows[idx].Cells[firstVisible];
                }
        }

        // ---------------- timers & events -------
        bool useDelay = entries.Count > 200;
        Timer delayTimer = new Timer { Interval = 200 };
        delayTimer.Tick += delegate { delayTimer.Stop(); UpdateFilteredGrid(); };
        form.FormClosed += delegate { delayTimer.Dispose(); };

        searchBox.TextChanged += delegate
        {
            if (useDelay)
            {
                delayTimer.Stop();
                delayTimer.Start();
            }
            else
            {
                UpdateFilteredGrid();
            }
        };

        void ApplySortFromHeader(int columnIndex)
        {
            string colName = grid.Columns[columnIndex].HeaderText;
            SortCriteria existing = sortCriteria.FirstOrDefault(s => s.ColumnName == colName);

            if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
            {
                if (existing != null) sortCriteria.Remove(existing);
            }
            else
            {
                if (existing != null)
                {
                    existing.Direction = existing.Direction == ListSortDirection.Ascending
                        ? ListSortDirection.Descending
                        : ListSortDirection.Ascending;
                    sortCriteria.Remove(existing);
                }
                else
                {
                    existing = new SortCriteria
                    {
                        ColumnName = colName,
                        Direction = ListSortDirection.Ascending
                    };
                }
                sortCriteria.Insert(0, existing);
                if (sortCriteria.Count > 3)
                    sortCriteria = sortCriteria.Take(3).ToList();
            }

            UpdateFilteredGrid();
        }

        grid.ColumnHeaderMouseClick += (s, e) => ApplySortFromHeader(e.ColumnIndex);

        void FinishSelection()
        {
            selectedEntries.Clear();
            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();
                for (int c = 0; c < propertyNames.Count; c++)
                    dict[propertyNames[c]] = row.Cells[c].Value;
                selectedEntries.Add(dict);
            }
            form.Close();
        }

        grid.CellDoubleClick += (s, e) => FinishSelection();

        // key handling shared by grid + searchBox
        void HandleKeyDown(KeyEventArgs e, Control sender)
        {
            if (e.KeyCode == Keys.Escape)
            {
                selectedEntries.Clear();   // nothing chosen
                form.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                FinishSelection();
            }
            else if (e.KeyCode == Keys.Tab && sender == grid && !e.Shift)
            {
                searchBox.Focus();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Tab && sender == searchBox && !e.Shift)
            {
                grid.Focus();
                e.Handled = true;
            }
            else if ((e.KeyCode == Keys.Down || e.KeyCode == Keys.Up) && sender == searchBox)
            {
                if (grid.Rows.Count > 0)
                {
                    grid.Focus();
                    int newIdx = 0;
                    if (grid.SelectedRows.Count > 0)
                    {
                        int curIdx = grid.SelectedRows[0].Index;
                        newIdx = e.KeyCode == Keys.Down
                            ? Math.Min(curIdx + 1, grid.Rows.Count - 1)
                            : Math.Max(curIdx - 1, 0);
                    }
                    int firstVisible = GetFirstVisibleColumnIndex();
                    if (firstVisible >= 0)
                    {
                        grid.ClearSelection();
                        grid.Rows[newIdx].Selected = true;
                        grid.CurrentCell = grid.Rows[newIdx].Cells[firstVisible];
                    }
                    e.Handled = true;
                }
            }
            else if ((e.KeyCode == Keys.Right || e.KeyCode == Keys.Left) &&
                     (e.Shift && sender == grid))
            {
                int offset = e.KeyCode == Keys.Right ? 1000 : -1000;
                grid.HorizontalScrollingOffset += offset;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Right && sender == grid)
            {
                grid.HorizontalScrollingOffset += 50;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Left && sender == grid)
            {
                grid.HorizontalScrollingOffset = Math.Max(grid.HorizontalScrollingOffset - 50, 0);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.D && e.Alt)
            {
                searchBox.Focus();
                e.Handled = true;
            }
        }

        grid.KeyDown += (s, e) => HandleKeyDown(e, grid);
        searchBox.KeyDown += (s, e) => HandleKeyDown(e, searchBox);

        // ---------------- initial layout -------
        form.Load += delegate
        {
            grid.AutoResizeColumns();

            int padding = 20;
            int rowsHeight = grid.Rows.GetRowsHeight(DataGridViewElementStates.Visible);
            int reqHeight = rowsHeight + grid.ColumnHeadersHeight +
                            2 * grid.RowTemplate.Height +
                            SystemInformation.HorizontalScrollBarHeight + 20;

            int availHeight = Screen.PrimaryScreen.WorkingArea.Height - padding * 2;
            form.Height = Math.Min(reqHeight, availHeight);

            if (spanAllScreens)
            {
                form.Width = Screen.AllScreens.Sum(s => s.WorkingArea.Width);
                form.Location = new Point(
                    Screen.AllScreens.Min(s => s.Bounds.X),
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2);
            }
            else
            {
                int reqWidth = grid.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                              + SystemInformation.VerticalScrollBarWidth + 43;
                form.Width = Math.Min(reqWidth, Screen.PrimaryScreen.WorkingArea.Width - padding * 2);
                form.Location = new Point(
                    (Screen.PrimaryScreen.WorkingArea.Width - form.Width) / 2,
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2);
            }
        };

        // ---------------- show -----------------
        form.Controls.Add(grid);
        form.Controls.Add(searchBox);
        searchBox.Select();
        form.ShowDialog();

        return selectedEntries;
    }
}
