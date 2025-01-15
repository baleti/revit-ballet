using System.Collections.Generic;
using System.Windows.Forms;
using System;
using System.Linq;
using System.Drawing;
using System.ComponentModel;

public partial class CustomGUIs
{
    public static List<Dictionary<string, object>> DataGrid(List<Dictionary<string, object>> entries, List<string> propertyNames, bool spanAllScreens, List<int> initialSelectionIndices = null)
    {
        List<Dictionary<string, object>> filteredEntries = new List<Dictionary<string, object>>();
        bool escapePressed = false;  // Flag to track if Escape was pressed

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

        // Manually populate DataGridView rows
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

        // Pre-select rows based on initialSelectionIndices
        bool hasInitialSelection = initialSelectionIndices != null && initialSelectionIndices.Any();

        if (initialSelectionIndices != null && initialSelectionIndices.Any())
        {
            foreach (var index in initialSelectionIndices)
            {
                if (index >= 0 && index < dataGridView.Rows.Count)
                {
                    dataGridView.Rows[index].Selected = true;
                    if (dataGridView.CurrentCell == null)
                    {
                        dataGridView.CurrentCell = dataGridView.Rows[index].Cells[0]; // Set focus to the first cell of the first selected row
                    }
                }
            }
        }

        var searchBox = new TextBox();
        searchBox.Dock = DockStyle.Top;

        searchBox.TextChanged += (sender, e) =>
        {
            string searchText = searchBox.Text.ToLower();

            // Step 1: Extract the exclusion terms starting with '!'
            var exclusionTerms = searchText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Where(term => term.StartsWith("!"))
                                           .Select(term => term.Substring(1))
                                           .ToList();

            // Step 2: Extract inclusion terms (terms not starting with '!')
            var searchTerms = searchText.Split(new[] { "||" }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(query => query.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

            // Step 3: Filter entries based on AND/OR/NOT logic
            var filtered = entries.Where(entry =>
            {
                // Exclude entries containing any of the exclusion terms
                bool exclude = exclusionTerms.Any(exclTerm =>
                    entry.Values.Any(value => value != null && value.ToString().IndexOf(exclTerm, StringComparison.OrdinalIgnoreCase) >= 0));

                if (exclude)
                {
                    return false; // Skip this entry if any exclusion terms match
                }

                // Check if the entry satisfies all AND conditions (terms in OR groups)
                foreach (var orQuery in searchTerms)
                {
                    bool orQueryMatched = true;
                    foreach (var term in orQuery)
                    {
                        bool isNegation = term.StartsWith("!");
                        string actualTerm = isNegation ? term.Substring(1) : term;
                        bool termFound = false;

                        foreach (var propertyName in propertyNames)
                        {
                            var propertyValue = entry.ContainsKey(propertyName) ? entry[propertyName]?.ToString() : null;
                            if (propertyValue != null && propertyValue.IndexOf(actualTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                termFound = true;
                                break; // Term found in any property, break inner loop
                            }
                        }

                        if (isNegation ? termFound : !termFound)
                        {
                            orQueryMatched = false;
                            break; // If term does not match in AND condition, break
                        }
                    }

                    if (orQueryMatched)
                    {
                        return true; // At least one OR group matches, include the entry
                    }
                }

                return false; // If no OR group matches, exclude the entry
            }).ToList();

            // Update DataGridView with filtered results
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

            // Reapply sorting by the first column
            if (dataGridView.Columns.Count > 0)
            {
                dataGridView.Sort(dataGridView.Columns[0], ListSortDirection.Ascending);
            }

            // Adjust form size
            dataGridView.AutoResizeColumns();
            int requiredWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible) + SystemInformation.VerticalScrollBarWidth + 50;
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

        dataGridView.CellDoubleClick += (sender, e) =>
        {
            if (e.RowIndex >= 0)
            {
                // Clear previously filtered entries if any
                filteredEntries.Clear();

                foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                {
                    Dictionary<string, object> selectedEntry = new Dictionary<string, object>();
                    for (int i = 0; i < propertyNames.Count; i++)
                    {
                        selectedEntry[propertyNames[i]] = selectedRow.Cells[i].Value;
                    }
                    filteredEntries.Add(selectedEntry);
                }

                // Indicate we've accepted a selection
                escapePressed = true;
                form.Close();
            }
        };


        // KeyDown event handling for DataGridView
        dataGridView.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                escapePressed = true;  // Set flag when Escape is pressed
                form.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                // Clear previously filtered entries if any
                filteredEntries.Clear();

                foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                {
                    Dictionary<string, object> selectedEntry = new Dictionary<string, object>();
                    for (int i = 0; i < propertyNames.Count; i++)
                    {
                        selectedEntry[propertyNames[i]] = selectedRow.Cells[i].Value;
                    }
                    filteredEntries.Add(selectedEntry);
                }
                escapePressed = true;  // Set flag when Escape is pressed
                form.Close();
            }
            else if (e.KeyCode == Keys.Tab)
            {
                searchBox.Focus();
            }
            // Check for Shift + Left/Right Arrow keys
            else if ((e.KeyCode == Keys.Right || e.KeyCode == Keys.Left) && e.Shift)
            {
                int totalColumnWidth = dataGridView.Columns.Cast<DataGridViewColumn>().Sum(c => c.Width);
                int visibleWidth = dataGridView.ClientRectangle.Width;
                int currentOffset = dataGridView.HorizontalScrollingOffset;
                int scrollOffset = 1000; // You can adjust this value to change scroll speed

                if (e.KeyCode == Keys.Right)
                {
                    // Attempt to scroll right
                    if (currentOffset + visibleWidth < totalColumnWidth)
                    {
                        dataGridView.HorizontalScrollingOffset += scrollOffset;
                    }
                }
                else if (e.KeyCode == Keys.Left)
                {
                    // Attempt to scroll left
                    dataGridView.HorizontalScrollingOffset = Math.Max(currentOffset - scrollOffset, 0);
                }
                e.Handled = true;  // Prevent further processing of the key in other controls
            }
            else if (e.KeyCode == Keys.Right)
            {
                int totalColumnWidth = dataGridView.Columns.Cast<DataGridViewColumn>().Sum(c => c.Width);
                int visibleWidth = dataGridView.ClientRectangle.Width;
                int currentOffset = dataGridView.HorizontalScrollingOffset;
                int scrollOffset = 50; // You can adjust this value to change scroll speed

                if (currentOffset + visibleWidth < totalColumnWidth)
                {
                    dataGridView.HorizontalScrollingOffset += scrollOffset;
                }
                e.Handled = true;  // Prevent further processing of the key in other controls
            }
            else if (e.KeyCode == Keys.Left)
            {
                int totalColumnWidth = dataGridView.Columns.Cast<DataGridViewColumn>().Sum(c => c.Width);
                int visibleWidth = dataGridView.ClientRectangle.Width;
                int currentOffset = dataGridView.HorizontalScrollingOffset;
                int scrollOffset = 50; // You can adjust this value to change scroll speed

                dataGridView.HorizontalScrollingOffset = Math.Max(currentOffset - scrollOffset, 0);
                e.Handled = true;  // Prevent further processing of the key in other controls
            }
        };

        // KeyDown event handling for SearchBox
        searchBox.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                escapePressed = true;  // Set flag when Escape is pressed
                form.Close();
            }
            else if (e.KeyCode == Keys.Down)
            {
                if (dataGridView.Rows.Count > 0)
                {
                    dataGridView.Focus();
                    // If a row is already selected, move the selection down by one row
                    if (dataGridView.SelectedRows.Count > 0)
                    {
                        int currentIndex = dataGridView.SelectedRows[0].Index;
                        int nextIndex = Math.Min(currentIndex + 1, dataGridView.Rows.Count - 1);

                        dataGridView.ClearSelection();
                        dataGridView.Rows[nextIndex].Selected = true;
                        dataGridView.CurrentCell = dataGridView.Rows[nextIndex].Cells[0];
                    }
                    else
                    {
                        // Select the first row if no row is currently selected
                        dataGridView.Rows[0].Selected = true;
                        dataGridView.CurrentCell = dataGridView.Rows[0].Cells[0];
                    }
                }
            }
            else if (e.KeyCode == Keys.Up)
            {
                if (dataGridView.Rows.Count > 0)
                {
                    dataGridView.Focus();

                    // If a row is already selected, move the selection up by one row
                    if (dataGridView.SelectedRows.Count > 0)
                    {
                        int currentIndex = dataGridView.SelectedRows[0].Index;
                        int previousIndex = Math.Max(currentIndex - 1, 0);

                        dataGridView.ClearSelection();
                        dataGridView.Rows[previousIndex].Selected = true;
                        dataGridView.CurrentCell = dataGridView.Rows[previousIndex].Cells[0];
                    }
                    else
                    {
                        // Select the first row if no row is currently selected
                        dataGridView.Rows[0].Selected = true;
                        dataGridView.CurrentCell = dataGridView.Rows[0].Cells[0];
                    }
                }
            }
            else if (e.KeyCode == Keys.Enter)
            {
                dataGridView.Focus();
                foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                {
                    Dictionary<string, object> selectedEntry = new Dictionary<string, object>();
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
            //sort the entries based on the first column
            dataGridView.Sort(dataGridView.Columns[0], ListSortDirection.Ascending);

            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            dataGridView.AutoResizeColumns();

            int padding = 20;
            int requiredHeight = dataGridView.Rows.GetRowsHeight(DataGridViewElementStates.Visible) + dataGridView.ColumnHeadersHeight + 2 * dataGridView.Rows[0].Height + SystemInformation.HorizontalScrollBarHeight;
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
                int requiredWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible) + SystemInformation.VerticalScrollBarWidth + 43;

                // Calculate the available size minus padding
                int availableWidth = Screen.PrimaryScreen.WorkingArea.Width - padding * 2;

                // Ensure the form fits within the padded screen working area
                form.Height = Math.Min(requiredHeight, availableHeight);
                form.Width = Math.Min(requiredWidth, availableWidth);

                // Center the form manually if any dimension is adjusted
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
            // Extract filtered entries based on selected rows
            foreach (DataGridViewRow row in dataGridView.SelectedRows)
            {
                Dictionary<string, object> selectedEntry = new Dictionary<string, object>();
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
