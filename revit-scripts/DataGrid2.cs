using System.Collections.Generic;
using System.Windows.Forms;
using System;
using System.Linq;
using System.Drawing;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

public partial class CustomGUIs
{
    // Helper class to store sort criteria.
    private class SortCriteria
    {
        public string ColumnName { get; set; }
        public ListSortDirection Direction { get; set; }
    }

    // Helper class to wrap each data entry along with a precomputed searchable text.
    private class DataEntry
    {
        public Dictionary<string, object> Entry { get; set; }
        public string SearchText { get; set; }
    }

    // Precompiled regex to tokenize search queries.
    private static readonly Regex SearchTokenRegex = new Regex(
       @"(\$""[^""]+"":""[^""]+""|\$[^ ]+?:""[^""]+""|\$""[^""]+"":[^ ]+|\$[^ ]+?:[^ ]+|\$""[^""]+""|\$[^ ]+|""[^""]+""|\S+)",
       RegexOptions.Compiled);

    public static List<Dictionary<string, object>> DataGrid(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        bool spanAllScreens,
        List<int> initialSelectionIndices = null)
    {
        // List to hold the final (user–selected) entries.
        List<Dictionary<string, object>> filteredEntries = new List<Dictionary<string, object>>();
        bool escapePressed = false;
        List<SortCriteria> sortCriteria = new List<SortCriteria>();

        // Precompute searchable text for each row.
        var allDataEntries = entries.Select(e => new DataEntry
        {
            Entry = e,
            SearchText = string.Join(" ", propertyNames.Select(p => e.ContainsKey(p) && e[p] != null ? e[p].ToString() : "")).ToLower()
        }).ToList();

        // Initially, the filtered (and sorted) data is the full set.
        List<DataEntry> sortedEntries = new List<DataEntry>(allDataEntries);

        // ---------------------------
        // Set up a light–themed form with standard window decoration.
        Form form = new Form();
        form.StartPosition = FormStartPosition.CenterScreen;
        form.Text = $"Total Entries: {entries.Count}";
        form.BackColor = Color.White;
        form.ForeColor = Color.Black;
        form.FormBorderStyle = FormBorderStyle.Sizable;

        // ---------------------------
        // Create a search box at the top.
        TextBox searchBox = new TextBox
        {
            Dock = DockStyle.Top,
            BackColor = Color.White,
            ForeColor = Color.Black,
            BorderStyle = BorderStyle.FixedSingle,
            Height = 25
        };

        // ---------------------------
        // Create the DataGridView.
        DataGridView dataGridView = new DataGridView();
        dataGridView.Dock = DockStyle.Fill;
        dataGridView.VirtualMode = true;
        dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dataGridView.MultiSelect = true;
        dataGridView.ReadOnly = true;
        dataGridView.AutoGenerateColumns = false;
        dataGridView.AllowUserToAddRows = false;
        // Prevent accidental row resizing.
        dataGridView.AllowUserToResizeRows = false;

        // Set light–theme appearance.
        dataGridView.BackgroundColor = Color.White;
        dataGridView.RowHeadersVisible = false;
        dataGridView.EnableHeadersVisualStyles = true;
        dataGridView.DefaultCellStyle.BackColor = Color.White;
        dataGridView.DefaultCellStyle.ForeColor = Color.Black;
        dataGridView.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
        dataGridView.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;

        // Create data columns.
        foreach (string propertyName in propertyNames)
        {
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = propertyName,
                DataPropertyName = propertyName
            });
        }

        // Provide cell values on demand.
        dataGridView.CellValueNeeded += (sender, e) =>
        {
            if (e.RowIndex >= 0 && e.RowIndex < sortedEntries.Count &&
                e.ColumnIndex >= 0 && e.ColumnIndex < propertyNames.Count)
            {
                var entry = sortedEntries[e.RowIndex].Entry;
                string propertyName = dataGridView.Columns[e.ColumnIndex].DataPropertyName;
                e.Value = entry.ContainsKey(propertyName) ? entry[propertyName] : null;
            }
        };

        // RefreshGrid: update filtered list and row count.
        void RefreshGrid(IEnumerable<DataEntry> entriesToShow)
        {
            sortedEntries = entriesToShow.ToList();
            dataGridView.RowCount = sortedEntries.Count;
            dataGridView.Invalidate();
        }
        RefreshGrid(sortedEntries);

        int GetFirstVisibleColumnIndex()
        {
            var visibleColumn = dataGridView.Columns.Cast<DataGridViewColumn>().FirstOrDefault(c => c.Visible);
            return visibleColumn?.Index ?? -1;
        }

        // Restore any initial selection.
        if (initialSelectionIndices != null && initialSelectionIndices.Any())
        {
            int firstVisibleCol = GetFirstVisibleColumnIndex();
            foreach (var index in initialSelectionIndices)
            {
                if (index >= 0 && index < dataGridView.RowCount && firstVisibleCol != -1)
                {
                    dataGridView.Rows[index].Selected = true;
                    dataGridView.CurrentCell = dataGridView.Rows[index].Cells[firstVisibleCol];
                }
            }
        }

        // -------------
        // Multi–column sorting.
        dataGridView.ColumnHeaderMouseClick += (sender, e) =>
        {
            string columnName = dataGridView.Columns[e.ColumnIndex].HeaderText;
            var existing = sortCriteria.FirstOrDefault(c => c.ColumnName == columnName);

            if (Control.ModifierKeys == Keys.Shift)
            {
                // If shift is held, remove this column from sort criteria.
                if (existing != null)
                    sortCriteria.Remove(existing);
            }
            else
            {
                // Toggle sort order.
                if (existing != null)
                {
                    existing.Direction = existing.Direction == ListSortDirection.Ascending
                        ? ListSortDirection.Descending
                        : ListSortDirection.Ascending;
                    sortCriteria.Remove(existing);
                }
                else
                {
                    existing = new SortCriteria { ColumnName = columnName, Direction = ListSortDirection.Ascending };
                }
                // Make it primary sort key.
                sortCriteria.Insert(0, existing);
                if (sortCriteria.Count > 3)
                    sortCriteria = sortCriteria.Take(3).ToList();
            }

            IOrderedEnumerable<DataEntry> ordered = null;
            foreach (var criteria in sortCriteria)
            {
                // For simplicity, sort on the string value of the property.
                if (ordered == null)
                {
                    ordered = criteria.Direction == ListSortDirection.Ascending
                        ? sortedEntries.OrderBy(x => x.Entry.ContainsKey(criteria.ColumnName) && x.Entry[criteria.ColumnName] != null ? x.Entry[criteria.ColumnName].ToString() : "")
                        : sortedEntries.OrderByDescending(x => x.Entry.ContainsKey(criteria.ColumnName) && x.Entry[criteria.ColumnName] != null ? x.Entry[criteria.ColumnName].ToString() : "");
                }
                else
                {
                    ordered = criteria.Direction == ListSortDirection.Ascending
                        ? ordered.ThenBy(x => x.Entry.ContainsKey(criteria.ColumnName) && x.Entry[criteria.ColumnName] != null ? x.Entry[criteria.ColumnName].ToString() : "")
                        : ordered.ThenByDescending(x => x.Entry.ContainsKey(criteria.ColumnName) && x.Entry[criteria.ColumnName] != null ? x.Entry[criteria.ColumnName].ToString() : "");
                }
            }
            sortedEntries = ordered?.ToList() ?? sortedEntries;
            RefreshGrid(sortedEntries);
        };

        // -------------
        // Filtering logic: precomputed searchable text is used.
        List<DataEntry> FilterEntries(string query)
        {
            var tokens = SearchTokenRegex.Matches(query).Cast<Match>()
                        .Select(m => m.Value.Trim())
                        .Where(t => !string.IsNullOrEmpty(t))
                        .ToList();

            List<(List<string> colParts, string valFilter)> columnValueFilters = new List<(List<string>, string)>();
            List<List<string>> columnVisibilityFilters = new List<List<string>>();
            List<string> generalFilters = new List<string>();

            foreach (var token in tokens)
            {
                if (token.StartsWith("$"))
                {
                    int colonIndex = token.IndexOf(':');
                    if (colonIndex > 0)
                    {
                        string colPart = token.Substring(1, colonIndex - 1);
                        string valPart = token.Substring(colonIndex + 1);
                        bool wasQuoted = colPart.StartsWith("\"") && colPart.EndsWith("\"");
                        string cleanCol = StripQuotes(colPart).ToLower();
                        List<string> colParts = wasQuoted
                            ? cleanCol.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList()
                            : new List<string> { cleanCol };
                        string valFilter = StripQuotes(valPart).ToLower();
                        if (colParts.Count > 0 && !string.IsNullOrWhiteSpace(valFilter))
                            columnValueFilters.Add((colParts, valFilter));
                    }
                    else
                    {
                        string content = token.Substring(1);
                        bool wasQuoted = content.StartsWith("\"") && content.EndsWith("\"");
                        string cleanContent = StripQuotes(content).ToLower();
                        List<string> parts = wasQuoted
                            ? cleanContent.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList()
                            : new List<string> { cleanContent };
                        if (parts.Count > 0)
                            columnVisibilityFilters.Add(parts);
                    }
                }
                else
                {
                    string filter = StripQuotes(token).ToLower();
                    if (!string.IsNullOrWhiteSpace(filter))
                        generalFilters.Add(filter);
                }
            }

            // Filter using precomputed search text.
            return allDataEntries.Where(de =>
                   // For each column-specific filter, require at least one property contains the text.
                   columnValueFilters.All(cvf =>
                       propertyNames.Any(p =>
                           de.Entry.ContainsKey(p) && de.Entry[p] != null &&
                           de.Entry[p].ToString().ToLower().Contains(cvf.valFilter)))
                   // And for general tokens, the precomputed search text must contain the token.
                   && generalFilters.All(gf => de.SearchText.Contains(gf))
            ).ToList();
        }

        // -------------
        // Asynchronous filtering: use a 200ms timer and background task.
        Timer searchTimer = new Timer();
        searchTimer.Interval = 200;
        searchTimer.Tick += (s, args) =>
        {
            searchTimer.Stop();
            Task.Run(() =>
            {
                var filtered = FilterEntries(searchBox.Text);
                form.Invoke(new Action(() =>
                {
                    RefreshGrid(filtered);
                    dataGridView.AutoResizeColumns();
                    // Recenter the form.
                    form.Location = new Point(
                        (Screen.PrimaryScreen.WorkingArea.Width - form.Width) / 2,
                        (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2);
                }));
            });
        };

        searchBox.TextChanged += (sender, e) =>
        {
            searchTimer.Stop();
            searchTimer.Start();
        };

        // -------------
        // --- Multi-range selection hack ---
        // We want to allow the user to select an additional range without clearing previous selection.
        int lastClickedRow = -1;
        dataGridView.CellMouseDown += (sender, e) =>
        {
            if (e.RowIndex < 0) return;
            // Check if both Ctrl and Shift are pressed.
            if ((Control.ModifierKeys & (Keys.Control | Keys.Shift)) == (Keys.Control | Keys.Shift))
            {
                // Get current selected row indices.
                var currentSelection = dataGridView.SelectedRows.Cast<DataGridViewRow>()
                                         .Select(r => r.Index).ToList();
                // If no previous row stored, use the clicked row.
                if (lastClickedRow == -1) lastClickedRow = e.RowIndex;
                int start = Math.Min(lastClickedRow, e.RowIndex);
                int end = Math.Max(lastClickedRow, e.RowIndex);
                for (int i = start; i <= end; i++)
                {
                    if (!currentSelection.Contains(i))
                        currentSelection.Add(i);
                }
                // Apply the union of the previous selection and new range.
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    row.Selected = currentSelection.Contains(row.Index);
                }
            }
            else
            {
                // Otherwise, update the lastClickedRow normally.
                lastClickedRow = e.RowIndex;
            }
        };

        // -------------
        // Key handling for navigation and selection.
        dataGridView.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                escapePressed = true;
                form.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                filteredEntries.Clear();
                foreach (DataGridViewRow row in dataGridView.SelectedRows)
                {
                    int index = row.Index;
                    if (index >= 0 && index < sortedEntries.Count)
                    {
                        var underlying = sortedEntries[index].Entry;
                        Dictionary<string, object> entry = new Dictionary<string, object>();
                        foreach (string prop in propertyNames)
                        {
                            entry[prop] = underlying.ContainsKey(prop) ? underlying[prop] : null;
                        }
                        filteredEntries.Add(entry);
                    }
                }
                form.Close();
            }
            else if (e.KeyCode == Keys.Tab && !e.Shift)
            {
                searchBox.Focus();
                e.Handled = true;
            }
            else if ((e.KeyCode == Keys.Right || e.KeyCode == Keys.Left) && e.Shift)
            {
                int scrollOffset = 1000;
                dataGridView.HorizontalScrollingOffset += e.KeyCode == Keys.Right ? scrollOffset : -scrollOffset;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Right)
            {
                dataGridView.HorizontalScrollingOffset += 50;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Left)
            {
                dataGridView.HorizontalScrollingOffset = Math.Max(dataGridView.HorizontalScrollingOffset - 50, 0);
                e.Handled = true;
            }
        };

        searchBox.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                escapePressed = true;
                form.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                filteredEntries.Clear();
                foreach (DataGridViewRow row in dataGridView.SelectedRows)
                {
                    int index = row.Index;
                    if (index >= 0 && index < sortedEntries.Count)
                    {
                        var underlying = sortedEntries[index].Entry;
                        Dictionary<string, object> entry = new Dictionary<string, object>();
                        foreach (string prop in propertyNames)
                        {
                            entry[prop] = underlying.ContainsKey(prop) ? underlying[prop] : null;
                        }
                        filteredEntries.Add(entry);
                    }
                }
                form.Close();
            }
            else if (e.KeyCode == Keys.Tab && !e.Shift)
            {
                dataGridView.Focus();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
            {
                if (dataGridView.RowCount > 0)
                {
                    dataGridView.Focus();
                    int newIndex = 0;
                    int firstVisibleCol = GetFirstVisibleColumnIndex();
                    if (dataGridView.SelectedRows.Count > 0)
                    {
                        int currentIndex = dataGridView.SelectedRows[0].Index;
                        newIndex = e.KeyCode == Keys.Down ? Math.Min(currentIndex + 1, dataGridView.RowCount - 1)
                                                         : Math.Max(currentIndex - 1, 0);
                    }
                    if (firstVisibleCol != -1)
                    {
                        dataGridView.ClearSelection();
                        dataGridView.Rows[newIndex].Selected = true;
                        dataGridView.CurrentCell = dataGridView.Rows[newIndex].Cells[firstVisibleCol];
                    }
                }
                e.Handled = true;
            }
        };

        // -------------
        // On form load, adjust window size and center it.
        form.Load += (sender, e) =>
        {
            if (sortedEntries.Count > 0)
            {
                // Default sort on first column.
                sortedEntries = sortedEntries.OrderBy(x => x.Entry.ContainsKey(propertyNames[0]) && x.Entry[propertyNames[0]] != null ? x.Entry[propertyNames[0]].ToString() : "").ToList();
                RefreshGrid(sortedEntries);
            }
            dataGridView.AutoResizeColumns();
            int padding = 20;
            int headerHeight = dataGridView.ColumnHeadersHeight;
            int rowHeight = sortedEntries.Count > 0 ? dataGridView.Rows[0].Height : dataGridView.RowTemplate.Height;
            int maxRowsToShow = 25;
            int rowCountToShow = Math.Min(sortedEntries.Count, maxRowsToShow);
            int requiredHeight = headerHeight + (rowCountToShow * rowHeight) + SystemInformation.HorizontalScrollBarHeight;
            int availableHeight = Screen.PrimaryScreen.WorkingArea.Height - padding * 2;
            form.Height = Math.Min(requiredHeight, availableHeight);

            if (spanAllScreens)
            {
                form.Width = Screen.AllScreens.Sum(s => s.WorkingArea.Width);
                form.Location = new Point(
                    Screen.AllScreens.Min(s => s.Bounds.X),
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2);
            }
            else
            {
                int requiredWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                    + SystemInformation.VerticalScrollBarWidth + 20; // extra right padding of 20
                form.Width = Math.Min(requiredWidth, Screen.PrimaryScreen.WorkingArea.Width - padding * 2);
                form.Location = new Point(
                    (Screen.PrimaryScreen.WorkingArea.Width - form.Width) / 2,
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2);
            }
        };

        // -------------
        // Add controls (search box at top, then grid).
        form.Controls.Add(dataGridView);
        form.Controls.Add(searchBox);

        searchBox.Select();
        form.ShowDialog();

        return filteredEntries;
    }

    // Helper method to remove enclosing quotes.
    private static string StripQuotes(string input)
    {
        return input.StartsWith("\"") && input.EndsWith("\"") && input.Length > 1
            ? input.Substring(1, input.Length - 2)
            : input;
    }
}
