using System.Collections.Generic;
using System.Windows.Forms;
using System;
using System.Linq;

public partial class CustomGUIs
{
    public static List<T> DataGrid<T>(List<T> entries, List<string> propertyNames, List<int> initialSelectionIndices = null, string Title = null)
    {
        List<T> filteredEntries = new List<T>();
        var form = new Form();
        form.StartPosition = FormStartPosition.CenterScreen;

        // Add title to form if not null
        if (Title != null)
            form.Text = Title;

        // DataGridView
        var dataGridView = new DataGridView();
        dataGridView.Dock = DockStyle.Fill;
        dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dataGridView.MultiSelect = true;
        dataGridView.ReadOnly = true;
        dataGridView.AutoGenerateColumns = false;
        dataGridView.AllowUserToAddRows = false;

        // Set up columns
        // Use "Programmatic" sort mode so we can manually control the sorting.
        foreach (var propertyName in propertyNames)
        {
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = propertyName,
                DataPropertyName = propertyName,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
            dataGridView.Columns.Add(column);
        }

        // Keep track of the last sorted column & direction to reapply after a search
        string lastSortedProperty = null;
        SortOrder lastSortOrder = SortOrder.None;

        // This list will store the currently displayed rows (filtered + sorted).
        // Start by showing all entries (i.e., no filter yet).
        var _entries = new List<T>(entries);

        // Function to apply the active sort, if any, to _entries
        void ApplySort()
        {
            if (!string.IsNullOrEmpty(lastSortedProperty) && _entries.Count > 0)
            {
                if (lastSortOrder == SortOrder.Ascending)
                {
                    _entries = _entries
                        .OrderBy(x => x.GetType().GetProperty(lastSortedProperty)?.GetValue(x, null))
                        .ToList();
                }
                else if (lastSortOrder == SortOrder.Descending)
                {
                    _entries = _entries
                        .OrderByDescending(x => x.GetType().GetProperty(lastSortedProperty)?.GetValue(x, null))
                        .ToList();
                }
            }
        }

        // Bind _entries to the grid
        void RebindData()
        {
            dataGridView.DataSource = null;
            dataGridView.DataSource = _entries;
        }

        // Initially bind all entries
        RebindData();
        dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        // ColumnHeaderMouseClick event to sort the grid
        dataGridView.ColumnHeaderMouseClick += (sender, e) =>
        {
            // Which column did the user click?
            var column = dataGridView.Columns[e.ColumnIndex];
            string propertyName = column.DataPropertyName;
            if (string.IsNullOrEmpty(propertyName)) return;

            // Toggle sort order for this column
            if (lastSortedProperty == propertyName)
            {
                // If clicking the same column, flip Ascending <-> Descending
                if (lastSortOrder == SortOrder.Ascending)
                {
                    lastSortOrder = SortOrder.Descending;
                }
                else
                {
                    lastSortOrder = SortOrder.Ascending;
                }
            }
            else
            {
                // If clicking a new column, default to ascending
                lastSortedProperty = propertyName;
                lastSortOrder = SortOrder.Ascending;
            }

            // Clear sort glyphs on all columns first
            foreach (DataGridViewColumn col in dataGridView.Columns)
            {
                col.HeaderCell.SortGlyphDirection = SortOrder.None;
            }
            // Show the sort glyph on the clicked column
            column.HeaderCell.SortGlyphDirection = lastSortOrder;

            // Re-sort _entries
            ApplySort();
            RebindData();
        };

        dataGridView.DataBindingComplete += (sender, e) =>
        {
            // Clear existing selections to avoid any unwanted selected rows.
            dataGridView.ClearSelection();

            // Check if initialSelectionIndices is not null and has elements
            if (initialSelectionIndices != null && initialSelectionIndices.Any())
            {
                foreach (int index in initialSelectionIndices)
                {
                    if (index >= 0 && index < dataGridView.Rows.Count)
                    {
                        dataGridView.Rows[index].Selected = true;
                        dataGridView.CurrentCell = dataGridView.Rows[index].Cells[0]; // Optionally set focus
                    }
                }
            }
        };

        // Search box
        var searchBox = new TextBox();
        searchBox.Dock = DockStyle.Top;
        searchBox.TextChanged += (sender, e) =>
        {
            // Split on "|" for OR blocks, then each block splits on spaces for AND logic
            string searchText = searchBox.Text.ToLower();
            var searchTerms = searchText.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(query => query.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

            // Filter from the *original* entries each time, then re-sort if needed
            _entries = entries.Where(entry =>
            {
                // For each "OR block"
                foreach (var orQuery in searchTerms)
                {
                    bool orQueryMatched = true;
                    // For each term in that block (AND)
                    foreach (var term in orQuery)
                    {
                        bool isNegation = term.StartsWith("!");
                        string actualTerm = isNegation ? term.Substring(1) : term;
                        bool termFound = false;

                        foreach (var propertyName in propertyNames)
                        {
                            var propertyValue = entry.GetType()
                                                     .GetProperty(propertyName)
                                                     ?.GetValue(entry, null)
                                                     ?.ToString();
                            if (propertyValue != null &&
                                propertyValue.IndexOf(actualTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                termFound = true;
                                break; // Found this term in at least one property
                            }
                        }

                        // If it's a negation term and we found it, then it's not a match
                        // If it's a normal term and we didn't find it, also not a match
                        if (isNegation ? termFound : !termFound)
                        {
                            orQueryMatched = false;
                            break;
                        }
                    }

                    // If the OR block matched, no need to check others
                    if (orQueryMatched)
                        return true;
                }

                // If no OR block matched, exclude
                return false;
            }).ToList();

            // Re-apply sort
            ApplySort();
            RebindData();
        };

        // CellDoubleClick event handling
        dataGridView.CellDoubleClick += (sender, e) =>
        {
            if (e.RowIndex >= 0) // Ensure the double-click is on a valid row
            {
                foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                {
                    T selectedItem = (T)selectedRow.DataBoundItem;
                    filteredEntries.Add(selectedItem);
                }
                form.Close(); // Close the form after action
            }
        };

        // KeyDown event handling for dataGridView
        dataGridView.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                form.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                {
                    T selectedItem = (T)selectedRow.DataBoundItem;
                    filteredEntries.Add(selectedItem);
                }
                form.Close();
            }
            else if (e.KeyCode == Keys.Tab)
            {
                searchBox.Focus();
            }
        };

        // KeyDown event handling for searchBox
        searchBox.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                form.Close();
            }
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
            {
                // Set focus to dataGridView
                dataGridView.Focus();
                if (dataGridView.CurrentRow != null)
                {
                    int rowCount = dataGridView.Rows.Count;
                    int currentIndex = dataGridView.CurrentRow.Index;
                    int nextIndex = (e.KeyCode == Keys.Down)
                        ? (currentIndex + 1) % rowCount
                        : (currentIndex - 1 + rowCount) % rowCount;

                    dataGridView.CurrentCell = dataGridView.Rows[nextIndex].Cells[0];
                    dataGridView.Rows[nextIndex].Selected = true;
                }
                else
                {
                    // If no row is selected, select the first row
                    if (dataGridView.Rows.Count > 0)
                    {
                        dataGridView.CurrentCell = dataGridView.Rows[0].Cells[0];
                        dataGridView.Rows[0].Selected = true;
                    }
                }
            }
            else if (e.KeyCode == Keys.Enter)
            {
                if (dataGridView.Rows.Count > 0)
                {
                    if (dataGridView.SelectedRows.Count == 0)
                    {
                        dataGridView.Rows[0].Selected = true;
                        dataGridView.CurrentCell = dataGridView.Rows[0].Cells[0];
                    }
                    foreach (DataGridViewRow selectedRow in dataGridView.SelectedRows)
                    {
                        T selectedItem = (T)selectedRow.DataBoundItem;
                        filteredEntries.Add(selectedItem);
                    }
                    form.Close();
                }
            }
            // Press space immediately after window is opened to return the next row
            else if (e.KeyCode == Keys.Space && string.IsNullOrWhiteSpace(searchBox.Text))
            {
                int rowCount = dataGridView.Rows.Count;
                if (rowCount > 0)
                {
                    int currentIndex = dataGridView.CurrentRow?.Index ?? -1;
                    int nextIndex;
                    if (e.Shift)
                        nextIndex = (currentIndex - 1 + rowCount) % rowCount; // previous row
                    else
                        nextIndex = (currentIndex + 1) % rowCount; // next row

                    dataGridView.ClearSelection();
                    dataGridView.CurrentCell = dataGridView.Rows[nextIndex].Cells[0];
                    dataGridView.Rows[nextIndex].Selected = true;

                    T selectedItem = (T)dataGridView.Rows[nextIndex].DataBoundItem;
                    filteredEntries.Add(selectedItem);
                    form.Close();
                }
            }
        };

        // Form load: Adjust form size
        form.Load += (sender, e) =>
        {
            int totalRowHeight = dataGridView.Rows.GetRowsHeight(DataGridViewElementStates.Visible);
            int requiredHeight = totalRowHeight
                + dataGridView.ColumnHeadersHeight
                + 2 * dataGridView.Rows[0].Height
                + SystemInformation.HorizontalScrollBarHeight;
            form.Height = requiredHeight;

            // Fit form inside working area
            if (requiredHeight > Screen.PrimaryScreen.WorkingArea.Height)
            {
                form.StartPosition = FormStartPosition.Manual;
                form.Top = 10;
                form.Height = Screen.PrimaryScreen.WorkingArea.Height - 10;
            }
            else if (requiredHeight > Screen.PrimaryScreen.WorkingArea.Height / 2)
            {
                form.StartPosition = FormStartPosition.Manual;
                form.Top = 10;
            }
            else
            {
                form.StartPosition = FormStartPosition.CenterScreen;
            }

            // Calculate and set column widths
            int formMaxWidth = 0;
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                int maxWidth = 0;
                // Calculate maximum width of header text
                int headerWidth = TextRenderer.MeasureText(column.HeaderText, dataGridView.ColumnHeadersDefaultCellStyle.Font).Width;

                // Calculate maximum width of cell content
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    if (row.Cells[column.Index].Value != null)
                    {
                        int cellWidth = TextRenderer.MeasureText(
                            row.Cells[column.Index].Value.ToString(),
                            dataGridView.DefaultCellStyle.Font).Width;
                        maxWidth = Math.Max(maxWidth, cellWidth);
                    }
                }

                // Set column width to the maximum of header text width and cell content width
                column.Width = Math.Max(headerWidth, maxWidth) + 10;
                formMaxWidth += column.Width;
            }
            // Add scrollbar + window frame offsets
            form.Width = formMaxWidth + SystemInformation.VerticalScrollBarWidth + 59;
        };

        // Add controls to form and show it
        form.Controls.Add(dataGridView);
        form.Controls.Add(searchBox);
        searchBox.Select();
        form.ShowDialog();

        return filteredEntries;
    }
}
