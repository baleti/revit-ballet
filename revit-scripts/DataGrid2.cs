// DataGrid2.cs   –   works on C# 7.3
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Helper types
    // ──────────────────────────────────────────────────────────────
    private struct ColumnValueFilter
    {
        public List<string> ColumnParts;   // column-header fragments to match
        public string       Value;         // value to look for in the cell
        public bool         IsExclusion;   // true ⇒ "must NOT contain"
    }

    private class SortCriteria
    {
        public string ColumnName { get; set; }
        public ListSortDirection Direction { get; set; }
    }

    /// <summary>A string comparer that sorts "A2" before "A10" and handles mixed numeric/text data.</summary>
    private sealed class NaturalComparer : IComparer<object>
    {
        public int Compare(object x, object y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            string s1 = x.ToString();
            string s2 = y.ToString();

            // Handle special non-numeric values that should be treated as text
            bool s1IsNonNumeric = IsNonNumericValue(s1);
            bool s2IsNonNumeric = IsNonNumericValue(s2);
            
            // If both are non-numeric, compare as strings
            if (s1IsNonNumeric && s2IsNonNumeric)
            {
                return string.Compare(s1, s2, StringComparison.OrdinalIgnoreCase);
            }
            
            // If one is non-numeric and one is numeric, non-numeric comes last
            if (s1IsNonNumeric && !s2IsNonNumeric) return 1;
            if (!s1IsNonNumeric && s2IsNonNumeric) return -1;

            // Try to parse as numbers
            double numA, numB;
            bool aIsNum = double.TryParse(s1, out numA);
            bool bIsNum = double.TryParse(s2, out numB);
            
            if (aIsNum && bIsNum) return numA.CompareTo(numB);

            // Fall back to natural string comparison
            return CompareNatural(s1, s2);
        }

        /// <summary>Checks if a value should be treated as non-numeric text (like "-", "N/A", etc.)</summary>
        private static bool IsNonNumericValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return true;
            
            // Single dash or common placeholder values
            if (value == "-" || 
                value.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("NULL", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("NONE", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("--", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            
            // If it can't be parsed as a number, treat as non-numeric
            double dummy;
            return !double.TryParse(value, out dummy);
        }

        private static int CompareNatural(string a, string b)
        {
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
                    if (numA.Length == 0) numA = "0";
                    if (numB.Length == 0) numB = "0";

                    int cmp = numA.Length.CompareTo(numB.Length);
                    if (cmp != 0) return cmp;

                    cmp = string.Compare(numA, numB, StringComparison.Ordinal);
                    if (cmp != 0) return cmp;
                }
                else
                {
                    int cmp = a[i].CompareTo(b[j]);
                    if (cmp != 0) return cmp;
                    i++;
                    j++;
                }
            }
            return a.Length.CompareTo(b.Length);
        }
    }

    private static readonly NaturalComparer naturalComparer = new NaturalComparer();

    // ──────────────────────────────────────────────────────────────
    //  Public entry point
    // ──────────────────────────────────────────────────────────────
    /// <summary>
    ///   Shows a filterable / sortable read-only grid and returns the
    ///   rows the user double-clicked or pressed Enter on.
    /// </summary>
    public static List<Dictionary<string, object>> DataGrid(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        bool spanAllScreens,
        List<int> initialSelectionIndices = null)
    {
        // ---------------- state ----------------
        List<Dictionary<string, object>> selectedEntries = new List<Dictionary<string, object>>();
        List<Dictionary<string, object>> workingSet      = new List<Dictionary<string, object>>(entries);
        List<SortCriteria>               sortCriteria    = new List<SortCriteria>();

        // ---------------- UI shell -------------
        Form form = new Form();
        form.StartPosition = FormStartPosition.CenterScreen;
        form.Text          = "Total Entries: " + entries.Count;
        form.BackColor     = Color.White;

        DataGridView grid = new DataGridView();
        grid.Dock               = DockStyle.Fill;
        grid.SelectionMode      = DataGridViewSelectionMode.FullRowSelect;
        grid.MultiSelect        = true;
        grid.ReadOnly           = true;
        grid.AutoGenerateColumns = false;
        grid.RowHeadersVisible   = false;
        grid.AllowUserToAddRows  = false;
        grid.AllowUserToResizeRows = false;
        grid.BackgroundColor     = Color.White;
        grid.RowTemplate.Height  = 25;
        
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
                HeaderText      = col,
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
            string searchText = searchBox.Text;

            // (1) split into tokens
            List<string> tokens = Regex.Matches(
                    searchText,
                    @"(\$""[^""]+?""::""[^""]+?""|\$""[^""]+?""\:\:[^ ]+|\$[^ ]+?::""[^""]+?""|\$[^ ]+?::[^ ]+|\$""[^""]+?""\:[^ ]+|\$[^ ]+?:[^ ]+|\$""[^""]+?""|\$[^ ]+|""[^""]+""|\S+)")
                .Cast<Match>()
                .Select(m => m.Value.Trim())
                .Where(t => t.Length > 0)
                .ToList();

            // buckets for the parsed filters
            List<List<string>>      colVisibilityFilters = new List<List<string>>();
            List<ColumnValueFilter> colValueFilters      = new List<ColumnValueFilter>();
            List<string>            generalFilters       = new List<string>();

            // (2) parse
            foreach (string rawToken in tokens)
            {
                bool isExcl = rawToken.StartsWith("!");
                string token = isExcl ? rawToken.Substring(1) : rawToken;

                // plain (general) token
                if (!token.StartsWith("$"))
                {
                    generalFilters.Add(isExcl ? "!" + StripQuotes(token)
                                              :       StripQuotes(token));
                    continue;
                }

                // token begins with '$'  → column-qualified
                string body        = token.Substring(1); // drop '$'
                int    dblColonPos = body.IndexOf("::", StringComparison.Ordinal);
                int    colonPos    = dblColonPos >= 0 ? dblColonPos
                                                      : body.IndexOf(':');

                string colPart = colonPos > 0 ? body.Substring(0, colonPos) : body;
                string valPart = "";
                if (colonPos > 0)
                {
                    int start = colonPos + (dblColonPos >= 0 ? 2 : 1);
                    valPart = body.Substring(start);
                }

                bool  quotedCol = colPart.StartsWith("\"") && colPart.EndsWith("\"");
                string cleanCol = StripQuotes(colPart).ToLowerInvariant();
                List<string> colPieces = quotedCol
                    ? cleanCol.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList()
                    : new List<string> { cleanCol };

                if (colPieces.Count == 0) continue;

                // (2a) column visibility (same logic as before)
                colVisibilityFilters.Add(colPieces);

                // (2b) row include / exclude
                if (!string.IsNullOrWhiteSpace(valPart))
                {
                    ColumnValueFilter f = new ColumnValueFilter
                    {
                        ColumnParts = colPieces,
                        Value       = StripQuotes(valPart).ToLowerInvariant(),
                        IsExclusion = isExcl
                    };
                    colValueFilters.Add(f);
                }
            }

            // (3) hide / show columns
            foreach (DataGridViewColumn col in grid.Columns)
            {
                string colName = col.HeaderText.ToLowerInvariant();
                bool show = colVisibilityFilters.Count == 0 ||
                            colVisibilityFilters.Any(parts =>
                                parts.All(p => colName.Contains(p)));
                col.Visible = show;
            }

            // (4) filter the rows
            List<Dictionary<string, object>> filtered = entries.Where(entry =>
            {
                // 4-a  column-qualified rules
                foreach (ColumnValueFilter f in colValueFilters)
                {
                    List<string> matchCols = propertyNames
                        .Where(p => f.ColumnParts.All(part =>
                                    p.ToLowerInvariant().Contains(part)))
                        .ToList();

                    if (matchCols.Count == 0) continue;   // column absent

                    bool valuePresent = matchCols.Any(c =>
                    {
                        object v;
                        return entry.TryGetValue(c, out v) &&
                               v != null &&
                               v.ToString().ToLowerInvariant().Contains(f.Value);
                    });

                    if (!f.IsExclusion && !valuePresent) return false; // required not found
                    if ( f.IsExclusion &&  valuePresent) return false; // forbidden found
                }

                // 4-b  general include / exclude (unchanged)
                if (generalFilters.Count > 0)
                {
                    string allValues = string.Join(" ",
                        entry.Values.Where(v => v != null)
                                    .Select(v => v.ToString().ToLowerInvariant()));

                    bool anyInc = generalFilters.Any(g => !g.StartsWith("!"));
                    bool anyExc = generalFilters.Any(g =>  g.StartsWith("!"));

                    if (anyInc &&
                        !generalFilters.Where(g => !g.StartsWith("!"))
                                       .All(inc => allValues.Contains(inc)))
                        return false;

                    if (anyExc &&
                        generalFilters.Where(g => g.StartsWith("!"))
                                       .Select(ex => ex.Substring(1))
                                       .Any(ex => allValues.Contains(ex)))
                        return false;
                }

                return true;
            }).ToList();

            // (5) sort (honours multi-column sortCriteria list)
            workingSet = filtered;
            if (sortCriteria.Count > 0)
            {
                IOrderedEnumerable<Dictionary<string, object>> ordered = null;
                foreach (SortCriteria sc in sortCriteria)
                {
                    Func<Dictionary<string, object>, object> key =
                        x => x.ContainsKey(sc.ColumnName) ? x[sc.ColumnName] : null;

                    if (ordered == null)
                    {
                        ordered = (sc.Direction == ListSortDirection.Ascending)
                            ? workingSet.OrderBy   (key, naturalComparer)
                            : workingSet.OrderByDescending(key, naturalComparer);
                    }
                    else
                    {
                        ordered = (sc.Direction == ListSortDirection.Ascending)
                            ? ordered.ThenBy   (key, naturalComparer)
                            : ordered.ThenByDescending(key, naturalComparer);
                    }
                }
                workingSet = ordered.ToList();
            }

            // (6) redraw grid
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
        bool useDelay     = entries.Count > 200;
        Timer delayTimer  = new Timer { Interval = 200 };
        delayTimer.Tick  += delegate { delayTimer.Stop(); UpdateFilteredGrid(); };
        form.FormClosed  += delegate { delayTimer.Dispose(); };

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
                        Direction  = ListSortDirection.Ascending
                    };
                }
                sortCriteria.Insert(0, existing);
                if (sortCriteria.Count > 3)
                    sortCriteria = sortCriteria.Take(3).ToList();
            }

            // resort using our custom logic
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

        grid.KeyDown      += (s, e) => HandleKeyDown(e, grid);
        searchBox.KeyDown += (s, e) => HandleKeyDown(e, searchBox);

        // ---------------- initial layout -------
        form.Load += delegate
        {
            // Don't use built-in sort - we'll handle it ourselves
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
                    (Screen.PrimaryScreen.WorkingArea.Width  - form.Width)  / 2,
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

    // ──────────────────────────────────────────────────────────────
    //  Utility
    // ──────────────────────────────────────────────────────────────
    private static string StripQuotes(string s)
    {
        return s.StartsWith("\"") && s.EndsWith("\"") && s.Length > 1
            ? s.Substring(1, s.Length - 2)
            : s;
    }
}
