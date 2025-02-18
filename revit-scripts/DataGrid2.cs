using System.Collections.Generic;
using System.Windows.Forms;
using System;
using System.Linq;
using System.Drawing;
using System.ComponentModel;
using System.Text.RegularExpressions;

public partial class CustomGUIs
{
    private class SortCriteria
    {
        public string ColumnName { get; set; }
        public ListSortDirection Direction { get; set; }
    }

    // NaturalComparer: compares two objects using "natural sort" rules.
    // It attempts to compare as numbers when possible.
    private class NaturalComparer : IComparer<object>
    {
        public int Compare(object x, object y)
        {
            // Handle null values.
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            string s1 = x.ToString();
            string s2 = y.ToString();

            // If both strings represent numbers, compare numerically.
            bool aIsNumber = double.TryParse(s1, out double numA);
            bool bIsNumber = double.TryParse(s2, out double numB);
            if (aIsNumber && bIsNumber)
            {
                return numA.CompareTo(numB);
            }
            return CompareNatural(s1, s2);
        }

        // A simple natural string comparison: digits are compared numerically.
        private int CompareNatural(string a, string b)
        {
            if (a == null && b == null) return 0;
            if (a == null) return -1;
            if (b == null) return 1;

            int i = 0, j = 0;
            while (i < a.Length && j < b.Length)
            {
                if (char.IsDigit(a[i]) && char.IsDigit(b[j]))
                {
                    int startI = i;
                    while (i < a.Length && char.IsDigit(a[i])) i++;
                    int startJ = j;
                    while (j < b.Length && char.IsDigit(b[j])) j++;

                    string numA = a.Substring(startI, i - startI).TrimStart('0');
                    string numB = b.Substring(startJ, j - startJ).TrimStart('0');
                    if (numA == "") numA = "0";
                    if (numB == "") numB = "0";

                    // First compare based on number of digits.
                    int cmp = numA.Length.CompareTo(numB.Length);
                    if (cmp != 0)
                        return cmp;
                    // Then lexicographically.
                    cmp = string.Compare(numA, numB, StringComparison.Ordinal);
                    if (cmp != 0)
                        return cmp;
                }
                else
                {
                    int cmp = a[i].CompareTo(b[j]);
                    if (cmp != 0)
                        return cmp;
                    i++;
                    j++;
                }
            }
            return a.Length.CompareTo(b.Length);
        }
    }

    // A static instance of NaturalComparer for use in sorting.
    private static readonly NaturalComparer naturalComparer = new NaturalComparer();

    public static List<Dictionary<string, object>> DataGrid(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        bool spanAllScreens,
        List<int> initialSelectionIndices = null)
    {
        List<Dictionary<string, object>> filteredEntries = new List<Dictionary<string, object>>();
        bool escapePressed = false;
        List<SortCriteria> sortCriteria = new List<SortCriteria>();
        List<Dictionary<string, object>> sortedEntries = new List<Dictionary<string, object>>(entries);

        Form form = new Form();
        form.StartPosition = FormStartPosition.CenterScreen;
        form.Text = $"Total Entries: {entries.Count}";
        form.BackColor = Color.White;

        DataGridView dataGridView = new DataGridView();
        dataGridView.Dock = DockStyle.Fill;
        dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dataGridView.MultiSelect = true;
        dataGridView.ReadOnly = true;
        dataGridView.AutoGenerateColumns = false;
        dataGridView.BackgroundColor = Color.White;
        dataGridView.AllowUserToAddRows = false;
        dataGridView.RowHeadersVisible = false;
        // Prevent row height modification.
        dataGridView.AllowUserToResizeRows = false;
        dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
        dataGridView.RowTemplate.Height = 25; // Set a fixed height for all rows.

        foreach (string propertyName in propertyNames)
        {
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = propertyName,
                DataPropertyName = propertyName
            });
        }

        void RefreshGrid(IEnumerable<Dictionary<string, object>> entriesToShow)
        {
            dataGridView.Rows.Clear();
            foreach (var entry in entriesToShow)
            {
                var row = new DataGridViewRow();
                row.CreateCells(dataGridView);
                for (int i = 0; i < propertyNames.Count; i++)
                {
                    row.Cells[i].Value = entry.ContainsKey(propertyNames[i]) ? entry[propertyNames[i]] : null;
                }
                dataGridView.Rows.Add(row);
            }
        }

        RefreshGrid(sortedEntries);

        int GetFirstVisibleColumnIndex()
        {
            var visibleColumn = dataGridView.Columns
                .Cast<DataGridViewColumn>()
                .FirstOrDefault(c => c.Visible);
            return visibleColumn?.Index ?? -1;
        }

        if (initialSelectionIndices != null && initialSelectionIndices.Any())
        {
            int firstVisibleCol = GetFirstVisibleColumnIndex();
            foreach (var index in initialSelectionIndices)
            {
                if (index >= 0 && index < dataGridView.Rows.Count && firstVisibleCol != -1)
                {
                    dataGridView.Rows[index].Selected = true;
                    dataGridView.CurrentCell = dataGridView.Rows[index].Cells[firstVisibleCol];
                }
            }
        }

        TextBox searchBox = new TextBox { Dock = DockStyle.Top };
        Timer delayTimer = new Timer { Interval = 200 };
        bool useDelay = entries.Count > 200;

        delayTimer.Tick += (s, e) =>
        {
            delayTimer.Stop();
            UpdateFilteredGrid();
        };

        form.FormClosed += (s, e) =>
        {
            delayTimer.Stop();
            delayTimer.Dispose();
        };

        void UpdateFilteredGrid()
        {
            string searchText = searchBox.Text;

            var tokens = Regex.Matches(searchText,
                @"(\$""[^""]+"":""[^""]+""|\$[^ ]+?:""[^""]+""|\$""[^""]+"":[^ ]+|\$[^ ]+?:[^ ]+|\$""[^""]+""|\$[^ ]+|""[^""]+""|\S+)")
                .Cast<Match>()
                .Select(m => m.Value.Trim())
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList();

            List<List<string>> columnVisibilityFilters = new List<List<string>>();
            List<(List<string> colParts, string valFilter)> columnValueFilters = new List<(List<string>, string)>();
            List<string> generalFilters = new List<string>();

            foreach (string token in tokens)
            {
                if (token.StartsWith("$"))
                {
                    // Handle tokens with double colon "::"
                    if (token.Contains("::"))
                    {
                        int doubleColonIndex = token.IndexOf("::");
                        string colPart = token.Substring(1, doubleColonIndex - 1);
                        string valPart = token.Substring(doubleColonIndex + 2);
                        bool wasQuoted = colPart.StartsWith("\"") && colPart.EndsWith("\"");
                        string cleanCol = StripQuotes(colPart).ToLower();
                        List<string> colParts = wasQuoted
                            ? cleanCol.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList()
                            : new List<string> { cleanCol };

                        string valFilter = StripQuotes(valPart).ToLower();

                        if (colParts.Count > 0)
                        {
                            // Always add column filter.
                            columnVisibilityFilters.Add(colParts);
                            // Only add row filtering if a value is provided.
                            if (!string.IsNullOrWhiteSpace(valFilter))
                            {
                                columnValueFilters.Add((colParts, valFilter));
                            }
                        }
                    }
                    else
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
                            {
                                // With a single colon, only row filtering is applied.
                                columnValueFilters.Add((colParts, valFilter));
                            }
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
                            {
                                // Column filtering only.
                                columnVisibilityFilters.Add(parts);
                            }
                        }
                    }
                }
                else
                {
                    string filter = StripQuotes(token).ToLower();
                    if (!string.IsNullOrWhiteSpace(filter))
                    {
                        generalFilters.Add(filter);
                    }
                }
            }

            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                string columnName = column.HeaderText.ToLower();
                bool shouldShow = columnVisibilityFilters.Count == 0 ||
                    columnVisibilityFilters.Any(filterParts =>
                        filterParts.All(part => columnName.Contains(part)));
                column.Visible = shouldShow;
            }

            List<Dictionary<string, object>> filtered = entries.Where(entry =>
            {
                foreach (var (colParts, valFilter) in columnValueFilters)
                {
                    List<string> matchingColumns = propertyNames
                        .Where(p => colParts.All(part => p.ToLower().Contains(part)))
                        .ToList();

                    if (matchingColumns.Count == 0) return false;

                    bool valueFound = matchingColumns.Any(col =>
                        entry.ContainsKey(col) &&
                        entry[col]?.ToString().ToLower().Contains(valFilter) == true);

                    if (!valueFound) return false;
                }

                if (generalFilters.Count > 0)
                {
                    string joinedValues = string.Join(" ", entry.Values
                        .Where(v => v != null)
                        .Select(v => v.ToString().ToLower()));

                    bool hasExclusions = generalFilters.Any(f => f.StartsWith("!"));
                    bool hasInclusions = generalFilters.Any(f => !f.StartsWith("!"));

                    if (hasInclusions && !generalFilters
                        .Where(f => !f.StartsWith("!"))
                        .All(f => joinedValues.Contains(f)))
                    {
                        return false;
                    }

                    if (hasExclusions && generalFilters
                        .Where(f => f.StartsWith("!"))
                        .Select(f => f.Substring(1))
                        .Any(excl => joinedValues.Contains(excl)))
                    {
                        return false;
                    }
                }

                return true;
            }).ToList();

            sortedEntries = filtered;
            if (sortCriteria.Count > 0)
            {
                IOrderedEnumerable<Dictionary<string, object>> ordered = null;
                foreach (var criteria in sortCriteria)
                {
                    if (ordered == null)
                    {
                        ordered = criteria.Direction == ListSortDirection.Ascending
                            ? sortedEntries.OrderBy(x => x[criteria.ColumnName], naturalComparer)
                            : sortedEntries.OrderByDescending(x => x[criteria.ColumnName], naturalComparer);
                    }
                    else
                    {
                        ordered = criteria.Direction == ListSortDirection.Ascending
                            ? ordered.ThenBy(x => x[criteria.ColumnName], naturalComparer)
                            : ordered.ThenByDescending(x => x[criteria.ColumnName], naturalComparer);
                    }
                }
                sortedEntries = ordered?.ToList() ?? sortedEntries;
            }

            RefreshGrid(sortedEntries);

            dataGridView.AutoResizeColumns();
            int requiredWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                + SystemInformation.VerticalScrollBarWidth + 50;
            form.Width = Math.Min(requiredWidth, Screen.PrimaryScreen.WorkingArea.Width - 20);
        }

        searchBox.TextChanged += (sender, e) =>
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

        dataGridView.ColumnHeaderMouseClick += (sender, e) =>
        {
            string columnName = dataGridView.Columns[e.ColumnIndex].HeaderText;
            var existing = sortCriteria.FirstOrDefault(c => c.ColumnName == columnName);

            if (Control.ModifierKeys == Keys.Shift)
            {
                if (existing != null)
                {
                    sortCriteria.Remove(existing);
                }
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
                        ColumnName = columnName,
                        Direction = ListSortDirection.Ascending
                    };
                }

                sortCriteria.Insert(0, existing);
                if (sortCriteria.Count > 3) sortCriteria = sortCriteria.Take(3).ToList();
            }

            IOrderedEnumerable<Dictionary<string, object>> ordered = null;
            foreach (var criteria in sortCriteria)
            {
                if (ordered == null)
                {
                    ordered = criteria.Direction == ListSortDirection.Ascending
                        ? sortedEntries.OrderBy(x => x[criteria.ColumnName], naturalComparer)
                        : sortedEntries.OrderByDescending(x => x[criteria.ColumnName], naturalComparer);
                }
                else
                {
                    ordered = criteria.Direction == ListSortDirection.Ascending
                        ? ordered.ThenBy(x => x[criteria.ColumnName], naturalComparer)
                        : ordered.ThenByDescending(x => x[criteria.ColumnName], naturalComparer);
                }
            }

            sortedEntries = ordered?.ToList() ?? sortedEntries;
            RefreshGrid(sortedEntries);
        };

        void HandleSelection()
        {
            filteredEntries.Clear();
            foreach (DataGridViewRow row in dataGridView.SelectedRows)
            {
                Dictionary<string, object> entry = new Dictionary<string, object>();
                for (int i = 0; i < propertyNames.Count; i++)
                {
                    entry[propertyNames[i]] = row.Cells[i].Value;
                }
                filteredEntries.Add(entry);
            }
            escapePressed = true;
            form.Close();
        }

        dataGridView.CellDoubleClick += (sender, e) => HandleSelection();

        Action<KeyEventArgs> handleAltD = e =>
        {
            if ((e.KeyCode == Keys.D) && (e.Alt))
            {
                searchBox.Focus();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        };

        dataGridView.KeyDown += (sender, e) =>
        {
            handleAltD(e);

            if (e.KeyCode == Keys.Escape)
            {
                escapePressed = true;
                form.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                HandleSelection();
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
            handleAltD(e);

            if (e.KeyCode == Keys.Escape)
            {
                escapePressed = true;
                form.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                HandleSelection();
            }
            else if (e.KeyCode == Keys.Tab && !e.Shift)
            {
                dataGridView.Focus();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
            {
                if (dataGridView.Rows.Count > 0)
                {
                    dataGridView.Focus();

                    int newIndex = 0;
                    int firstVisibleCol = GetFirstVisibleColumnIndex();

                    if (dataGridView.SelectedRows.Count > 0)
                    {
                        int currentIndex = dataGridView.SelectedRows[0].Index;
                        newIndex = e.KeyCode == Keys.Down
                            ? Math.Min(currentIndex + 1, dataGridView.Rows.Count - 1)
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

        form.Load += (sender, e) =>
        {
            dataGridView.Sort(dataGridView.Columns[0], ListSortDirection.Ascending);
            dataGridView.AutoResizeColumns();

            int padding = 20;
            int requiredHeight = dataGridView.Rows.GetRowsHeight(DataGridViewElementStates.Visible)
                + dataGridView.ColumnHeadersHeight + 2 * dataGridView.Rows[0].Height
                + SystemInformation.HorizontalScrollBarHeight
                + 20;
            int availableHeight = Screen.PrimaryScreen.WorkingArea.Height - padding * 2;

            form.Height = Math.Min(requiredHeight, availableHeight);

            if (spanAllScreens)
            {
                form.Width = Screen.AllScreens.Sum(s => s.WorkingArea.Width);
                form.Location = new Point(
                    Screen.AllScreens.Min(s => s.Bounds.X),
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2
                );
            }
            else
            {
                int requiredWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                    + SystemInformation.VerticalScrollBarWidth + 43;
                form.Width = Math.Min(requiredWidth, Screen.PrimaryScreen.WorkingArea.Width - padding * 2);
                form.Location = new Point(
                    (Screen.PrimaryScreen.WorkingArea.Width - form.Width) / 2,
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2
                );
            }
        };

        form.Controls.Add(dataGridView);
        form.Controls.Add(searchBox);
        searchBox.Select();
        form.ShowDialog();

        return filteredEntries;
    }

    private static string StripQuotes(string input)
    {
        return input.StartsWith("\"") && input.EndsWith("\"") && input.Length > 1
            ? input.Substring(1, input.Length - 2)
            : input;
    }
}
