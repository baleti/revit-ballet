using System.Collections.Generic;
using System.Windows.Forms;
using System;
using System.Linq;
using System.Drawing;
using System.ComponentModel;
using System.Text.RegularExpressions;

public partial class CustomGUIs
{
    public static List<Dictionary<string, object>> DataGrid(List<Dictionary<string, object>> entries, List<string> propertyNames, bool spanAllScreens, List<int> initialSelectionIndices = null)
    {
        List<Dictionary<string, object>> filteredEntries = new List<Dictionary<string, object>>();
        bool escapePressed = false;

        var form = new Form();
        form.StartPosition = FormStartPosition.CenterScreen;
        form.Text = $"Total Entries: {entries.Count}";
        form.BackColor = Color.White;

        var dataGridView = new DataGridView();
        dataGridView.Dock = DockStyle.Fill;
        dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dataGridView.MultiSelect = true;
        dataGridView.ReadOnly = true;
        dataGridView.AutoGenerateColumns = false;
        dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        dataGridView.BackgroundColor = Color.White;
        dataGridView.AllowUserToAddRows = false;

        foreach (var propertyName in propertyNames)
        {
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = propertyName,
                DataPropertyName = propertyName
            });
        }

        foreach (var entry in entries)
        {
            var row = new DataGridViewRow();
            row.CreateCells(dataGridView);

            for (int i = 0; i < propertyNames.Count; i++)
            {
                row.Cells[i].Value = entry.ContainsKey(propertyNames[i]) ? entry[propertyNames[i]] : null;
            }

            dataGridView.Rows.Add(row);
        }

        bool hasInitialSelection = initialSelectionIndices != null && initialSelectionIndices.Any();
        if (hasInitialSelection)
        {
            foreach (var index in initialSelectionIndices)
            {
                if (index >= 0 && index < dataGridView.Rows.Count)
                {
                    dataGridView.Rows[index].Selected = true;
                    if (dataGridView.CurrentCell == null)
                    {
                        dataGridView.CurrentCell = dataGridView.Rows[index].Cells[0];
                    }
                }
            }
        }

        var searchBox = new TextBox();
        searchBox.Dock = DockStyle.Top;

        searchBox.TextChanged += (sender, e) =>
        {
            string searchText = searchBox.Text;

            // Extract column filters using regex
            var columnFilterMatches = Regex.Matches(searchText, @"\$(\S+)");
            var columnFilters = columnFilterMatches.Cast<Match>()
                .Select(m => m.Groups[1].Value.ToLower())
                .ToList();

            // Remove column filters from search text to get row filter text
            string rowFilterText = Regex.Replace(searchText, @"\$\S+", "").Trim();
            string rowFilterTextLower = rowFilterText.ToLower();

            // Apply column visibility
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                string columnNameLower = column.HeaderText.ToLower();
                bool shouldShow = columnFilters.Count == 0 || columnFilters.Any(cf => columnNameLower.Contains(cf));
                column.Visible = shouldShow;
            }

            // Existing row filtering logic with exclusion terms and OR groups
            var exclusionTerms = rowFilterTextLower.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(term => term.StartsWith("!"))
                .Select(term => term.Substring(1))
                .ToList();

            var searchTerms = rowFilterTextLower.Split(new[] { "||" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(query => query.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                .ToList();

            var filtered = entries.Where(entry =>
            {
                if (exclusionTerms.Any(exclTerm =>
                    entry.Values.Any(value => value != null && value.ToString().ToLower().Contains(exclTerm))))
                {
                    return false;
                }

                foreach (var orQuery in searchTerms)
                {
                    bool orQueryMatched = true;
                    foreach (var term in orQuery)
                    {
                        bool isNegation = term.StartsWith("!");
                        string actualTerm = isNegation ? term.Substring(1) : term;
                        bool termFound = entry.Values.Any(value =>
                            value != null && value.ToString().ToLower().Contains(actualTerm));

                        if (isNegation ? termFound : !termFound)
                        {
                            orQueryMatched = false;
                            break;
                        }
                    }

                    if (orQueryMatched)
                    {
                        return true;
                    }
                }

                return false;
            }).ToList();

            dataGridView.Rows.Clear();
            foreach (var item in filtered)
            {
                var newRow = new DataGridViewRow();
                newRow.CreateCells(dataGridView);
                for (int i = 0; i < propertyNames.Count; i++)
                {
                    newRow.Cells[i].Value = item.ContainsKey(propertyNames[i]) ? item[propertyNames[i]] : null;
                }
                dataGridView.Rows.Add(newRow);
            }

            // Preserve original sizing and positioning logic
            if (dataGridView.Columns.Count > 0)
            {
                dataGridView.Sort(dataGridView.Columns[0], ListSortDirection.Ascending);
            }

            dataGridView.AutoResizeColumns();
            int requiredWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible) 
                + SystemInformation.VerticalScrollBarWidth + 50;
            int availableWidth = Screen.PrimaryScreen.WorkingArea.Width - 20;
            form.Width = Math.Min(requiredWidth, availableWidth);

            if (form.Width != requiredWidth)
            {
                form.StartPosition = FormStartPosition.Manual;
                form.Location = new Point(
                    (Screen.PrimaryScreen.WorkingArea.Width - form.Width) / 2,
                    form.Location.Y
                );
            }
        };

        // Keep all original event handlers
        dataGridView.CellDoubleClick += (sender, e) =>
        {
            if (e.RowIndex >= 0)
            {
                filteredEntries.Clear();
                foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                {
                    var selectedEntry = new Dictionary<string, object>();
                    for (int i = 0; i < propertyNames.Count; i++)
                    {
                        selectedEntry[propertyNames[i]] = selectedRow.Cells[i].Value;
                    }
                    filteredEntries.Add(selectedEntry);
                }
                escapePressed = true;
                form.Close();
            }
        };

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
                foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                {
                    var selectedEntry = new Dictionary<string, object>();
                    for (int i = 0; i < propertyNames.Count; i++)
                    {
                        selectedEntry[propertyNames[i]] = selectedRow.Cells[i].Value;
                    }
                    filteredEntries.Add(selectedEntry);
                }
                form.Close();
            }
            // Keep original scrolling and navigation logic
            else if ((e.KeyCode == Keys.Right || e.KeyCode == Keys.Left) && e.Shift)
            {
                int scrollOffset = 1000;
                if (e.KeyCode == Keys.Right)
                {
                    dataGridView.HorizontalScrollingOffset += scrollOffset;
                }
                else
                {
                    dataGridView.HorizontalScrollingOffset = Math.Max(
                        dataGridView.HorizontalScrollingOffset - scrollOffset, 0);
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Right)
            {
                dataGridView.HorizontalScrollingOffset += 50;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Left)
            {
                dataGridView.HorizontalScrollingOffset = Math.Max(
                    dataGridView.HorizontalScrollingOffset - 50, 0);
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
            else if (e.KeyCode == Keys.Down)
            {
                if (dataGridView.Rows.Count > 0)
                {
                    dataGridView.Focus();
                    if (dataGridView.SelectedRows.Count > 0)
                    {
                        int currentIndex = dataGridView.SelectedRows[0].Index;
                        int nextIndex = Math.Min(currentIndex + 1, dataGridView.Rows.Count - 1);
                        dataGridView.ClearSelection();
                        dataGridView.Rows[nextIndex].Selected = true;
                    }
                    else
                    {
                        dataGridView.Rows[0].Selected = true;
                    }
                    dataGridView.CurrentCell = dataGridView.SelectedRows[0].Cells[0];
                }
            }
            else if (e.KeyCode == Keys.Up)
            {
                if (dataGridView.Rows.Count > 0)
                {
                    dataGridView.Focus();
                    if (dataGridView.SelectedRows.Count > 0)
                    {
                        int currentIndex = dataGridView.SelectedRows[0].Index;
                        int previousIndex = Math.Max(currentIndex - 1, 0);
                        dataGridView.ClearSelection();
                        dataGridView.Rows[previousIndex].Selected = true;
                    }
                    else
                    {
                        dataGridView.Rows[0].Selected = true;
                    }
                    dataGridView.CurrentCell = dataGridView.SelectedRows[0].Cells[0];
                }
            }
            else if (e.KeyCode == Keys.Enter)
            {
                dataGridView.Focus();
                foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                {
                    var selectedEntry = new Dictionary<string, object>();
                    for (int i = 0; i < propertyNames.Count; i++)
                    {
                        selectedEntry[propertyNames[i]] = selectedRow.Cells[i].Value;
                    }
                    filteredEntries.Add(selectedEntry);
                }
                form.Close();
            }
        };

        form.Load += (sender, e) =>
        {
            dataGridView.Sort(dataGridView.Columns[0], ListSortDirection.Ascending);
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            dataGridView.AutoResizeColumns();

            int padding = 20;
            int requiredHeight = dataGridView.Rows.GetRowsHeight(DataGridViewElementStates.Visible) 
                + dataGridView.ColumnHeadersHeight + 2 * dataGridView.Rows[0].Height 
                + SystemInformation.HorizontalScrollBarHeight;
            int availableHeight = Screen.PrimaryScreen.WorkingArea.Height - padding * 2;

            form.Height = Math.Min(requiredHeight, availableHeight);

            if (spanAllScreens)
            {
                int totalScreenWidth = Screen.AllScreens.Sum(screen => screen.WorkingArea.Width);
                form.Width = totalScreenWidth;
                form.StartPosition = FormStartPosition.Manual;
                form.Location = new Point(
                    Screen.AllScreens.OrderBy(screen => screen.Bounds.X).First().Bounds.X,
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2
                );
            }
            else
            {
                int requiredWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible) 
                    + SystemInformation.VerticalScrollBarWidth + 43;
                int availableWidth = Screen.PrimaryScreen.WorkingArea.Width - padding * 2;

                form.Height = Math.Min(requiredHeight, availableHeight);
                form.Width = Math.Min(requiredWidth, availableWidth);

                if (form.Height != requiredHeight || form.Width != requiredWidth)
                {
                    form.StartPosition = FormStartPosition.Manual;
                    form.Location = new Point(
                        (Screen.PrimaryScreen.WorkingArea.Width - form.Width) / 2,
                        (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2
                    );
                }
            }
        };

        form.Controls.Add(dataGridView);
        form.Controls.Add(searchBox);
        searchBox.Select();
        form.ShowDialog();

        if (!escapePressed)
        {
            foreach (DataGridViewRow row in dataGridView.SelectedRows)
            {
                var selectedEntry = new Dictionary<string, object>();
                for (int i = 0; i < propertyNames.Count; i++)
                {
                    selectedEntry[propertyNames[i]] = row.Cells[i].Value;
                }
                filteredEntries.Add(selectedEntry);
            }
        }

        return filteredEntries;
    }
}
