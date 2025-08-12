using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Main DataGrid Method
    // ──────────────────────────────────────────────────────────────

    public static List<Dictionary<string, object>> DataGrid(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        bool spanAllScreens,
        List<int> initialSelectionIndices = null)
    {
        if (entries == null || entries.Count == 0 || propertyNames == null || propertyNames.Count == 0)
            return new List<Dictionary<string, object>>();

        // Clear any previous cached data
        _cachedOriginalData = entries;
        _cachedFilteredData = entries;
        _searchIndexByColumn = null;
        _searchIndexAllColumns = null;
        _lastVisibleColumns.Clear();
        _lastColumnVisibilityFilter = "";

        // Build search index upfront for performance
        BuildSearchIndex(entries, propertyNames);

        // State variables
        List<Dictionary<string, object>> selectedEntries = new List<Dictionary<string, object>>();
        List<Dictionary<string, object>> workingSet = new List<Dictionary<string, object>>(entries);
        List<SortCriteria> sortCriteria = new List<SortCriteria>();

        // Create form
        Form form = new Form
        {
            StartPosition = FormStartPosition.CenterScreen,
            Text = "Total Entries: " + entries.Count,
            BackColor = Color.White
        };

        // Create DataGridView with virtual mode
        DataGridView grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
            ReadOnly = true,
            AutoGenerateColumns = false,
            RowHeadersVisible = false,
            AllowUserToAddRows = false,
            AllowUserToResizeRows = false,
            BackgroundColor = Color.White,
            RowTemplate = { Height = 18 },
            VirtualMode = true,
            ScrollBars = ScrollBars.Both
        };

        // Disable built-in sorting
        grid.SortCompare += (sender, e) =>
        {
            e.Handled = true;
            e.SortResult = naturalComparer.Compare(e.CellValue1, e.CellValue2);
        };

        // Add columns
        foreach (string col in propertyNames)
        {
            var column = new DataGridViewTextBoxColumn
            {
                Name = col,
                HeaderText = col,
                DataPropertyName = col,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
            grid.Columns.Add(column);
        }

        // Search box - keep original appearance
        TextBox searchBox = new TextBox { Dock = DockStyle.Top };

        // Set up virtual mode cell value handler
        grid.CellValueNeeded += (s, e) =>
        {
            if (e.RowIndex >= 0 && e.RowIndex < _cachedFilteredData.Count && e.ColumnIndex >= 0)
            {
                var row = _cachedFilteredData[e.RowIndex];
                string columnName = grid.Columns[e.ColumnIndex].Name;
                object value;
                e.Value = row.TryGetValue(columnName, out value) ? value : null;
            }
        };

        // Initialize grid with data
        grid.RowCount = workingSet.Count;

        // Helper to get first visible column
        Func<int> GetFirstVisibleColumnIndex = () =>
        {
            foreach (DataGridViewColumn c in grid.Columns)
                if (c.Visible) return c.Index;
            return -1;
        };

        // Helper to update filtered grid
        Action UpdateFilteredGrid = () =>
        {
            // Use optimized filtering
            var filteredData = ApplyFilters(_cachedOriginalData, propertyNames, searchBox.Text, grid);

            // Apply sorting
            workingSet = filteredData;
            if (sortCriteria.Count > 0)
            {
                workingSet = ApplySorting(workingSet, sortCriteria);
            }

            // Update cached filtered data and virtual grid row count
            _cachedFilteredData = workingSet;
            grid.RowCount = 0; // Force refresh
            grid.RowCount = workingSet.Count;

            // Update form title
            form.Text = "Total Entries: " + workingSet.Count + " / " + entries.Count;

            // Auto-resize columns only on initial load or if column count is reasonable
            if (grid.Columns.Count < 20)
            {
                grid.AutoResizeColumns();
            }

            int reqWidth = grid.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                          + SystemInformation.VerticalScrollBarWidth + 50;
            form.Width = Math.Min(reqWidth, Screen.PrimaryScreen.WorkingArea.Width - 20);
        };

        // Initial draw
        _cachedFilteredData = workingSet;
        grid.RowCount = workingSet.Count;

        // Initial selection
        if (initialSelectionIndices != null && initialSelectionIndices.Count > 0)
        {
            int firstVisible = GetFirstVisibleColumnIndex();
            bool currentCellSet = false;
            
            foreach (int idx in initialSelectionIndices)
            {
                if (idx >= 0 && idx < grid.Rows.Count && firstVisible >= 0)
                {
                    grid.Rows[idx].Selected = true;
                    
                    // Only set CurrentCell once, for the first valid selection
                    if (!currentCellSet)
                    {
                        grid.CurrentCell = grid.Rows[idx].Cells[firstVisible];
                        currentCellSet = true;
                    }
                }
            }
        }

        // Setup delay timer for large datasets
        bool useDelay = entries.Count > 200;
        Timer delayTimer = new Timer { Interval = 200 };
        delayTimer.Tick += delegate { delayTimer.Stop(); UpdateFilteredGrid(); };
        form.FormClosed += delegate { delayTimer.Dispose(); };

        // Search box text changed
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

        // Column header click for sorting
        grid.ColumnHeaderMouseClick += (s, e) =>
        {
            string colName = grid.Columns[e.ColumnIndex].HeaderText;
            SortCriteria existing = sortCriteria.FirstOrDefault(sc => sc.ColumnName == colName);

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
        };

        // Finish selection helper
        Action FinishSelection = () =>
        {
            selectedEntries.Clear();
            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                if (row.Index < _cachedFilteredData.Count)
                {
                    selectedEntries.Add(_cachedFilteredData[row.Index]);
                }
            }
            form.Close();
        };

        // Double-click to select
        grid.CellDoubleClick += (s, e) => FinishSelection();

        // Key handling - restore original behavior
        Action<KeyEventArgs, Control> HandleKeyDown = (e, sender) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                selectedEntries.Clear();
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
            else if ((e.KeyCode == Keys.Right || e.KeyCode == Keys.Left) && (e.Shift && sender == grid))
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
            else if (e.KeyCode == Keys.E && e.Control)
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string storagePath = Path.Combine(appData, "revit-scripts", "DataGrid-last-export-location");
                string initialPath = "";
                if (File.Exists(storagePath))
                {
                    initialPath = File.ReadAllText(storagePath).Trim();
                }
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "CSV Files|*.csv";
                sfd.Title = "Export DataGrid to CSV";
                sfd.DefaultExt = "csv";
                if (!string.IsNullOrEmpty(initialPath))
                {
                    string dir = Path.GetDirectoryName(initialPath);
                    if (Directory.Exists(dir))
                    {
                        sfd.InitialDirectory = dir;
                        sfd.FileName = Path.GetFileName(initialPath);
                    }
                }
                else
                {
                    sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    sfd.FileName = "DataGridExport.csv";
                }
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExportToCsv(grid, _cachedFilteredData, sfd.FileName);
                        Directory.CreateDirectory(Path.GetDirectoryName(storagePath));
                        File.WriteAllText(storagePath, sfd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error exporting: " + ex.Message);
                    }
                }
                e.Handled = true;
            }
        };

        grid.KeyDown += (s, e) => HandleKeyDown(e, grid);
        searchBox.KeyDown += (s, e) => HandleKeyDown(e, searchBox);

        // Form load - restore original sizing logic
        form.Load += delegate
        {
            grid.AutoResizeColumns();

            int padding = 20;
            int rowsHeight = grid.Rows.GetRowsHeight(DataGridViewElementStates.Visible);
            int reqHeight = rowsHeight + grid.ColumnHeadersHeight +
                            2 * grid.RowTemplate.Height +
                            SystemInformation.HorizontalScrollBarHeight + 30;

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

        // Add controls and show
        form.Controls.Add(grid);
        form.Controls.Add(searchBox);
        searchBox.Select();
        form.ShowDialog();

        return selectedEntries;
    }

    private static void ExportToCsv(DataGridView grid, List<Dictionary<string, object>> data, string filePath)
    {
        var visibleColumns = grid.Columns.Cast<DataGridViewColumn>()
            .Where(c => c.Visible)
            .OrderBy(c => c.DisplayIndex)
            .ToList();

        using (var writer = new StreamWriter(filePath))
        {
            // Write header
            writer.WriteLine(string.Join(",", visibleColumns.Select(c => CsvQuote(c.HeaderText))));

            // Write rows
            foreach (var row in data)
            {
                var values = visibleColumns.Select(c => CsvQuote(row.ContainsKey(c.Name) ? row[c.Name]?.ToString() ?? "" : ""));
                writer.WriteLine(string.Join(",", values));
            }
        }
    }

    private static string CsvQuote(string value)
    {
        if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
        return value;
    }
}
