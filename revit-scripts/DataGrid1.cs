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

        // Add static columns from propertyNames
        foreach (var propertyName in propertyNames)
        {
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = propertyName,
                DataPropertyName = propertyName
            });
        }

        dataGridView.DataSource = entries;
        dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

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
                        dataGridView.CurrentCell = dataGridView.Rows[index].Cells[0]; // Optionally set focus to the first cell of the row
                    }
                }
            }
        };

        // Search box
        var searchBox = new TextBox();
        searchBox.Dock = DockStyle.Top;
        var _entries = new List<T>();
        searchBox.TextChanged += (sender, e) =>
        {
            string searchText = searchBox.Text.ToLower();
            var searchTerms = searchText.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(query => query.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

            _entries = entries.Where(entry =>
            {
                // Check if any of the search terms match any property value
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
                            var propertyValue = entry.GetType().GetProperty(propertyName)?.GetValue(entry, null)?.ToString();
                            if (propertyValue != null && propertyValue.IndexOf(actualTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                termFound = true;
                                break; // Exit inner loop if term is found in any property
                            }
                        }
                        if (isNegation ? termFound : !termFound)
                        {
                            orQueryMatched = false;
                            break; // Exit if any term in OR query is not found in any property or found if it's a negation
                        }
                    }
                    if (orQueryMatched)
                        return true; // Return true if any term in OR query is found in any property
                }
                return false; // Return false if none of the OR queries are matched
            }).ToList();

            // Set sorted entries as the DataSource for the DataGridView
            dataGridView.DataSource = null;
            dataGridView.DataSource = _entries;
        };

        // CellDoubleClick event handling
        dataGridView.CellDoubleClick += (sender, e) =>
        {
            if (e.RowIndex >= 0) // Ensure the double-click is on a valid row
            {
                // Perform the action associated with pressing Enter
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

                // Check if there is a currently selected row
                if (dataGridView.CurrentRow != null)
                {
                    int rowCount = dataGridView.Rows.Count;
                    int currentIndex = dataGridView.CurrentRow.Index;

                    // Calculate the next index based on the key pressed
                    int nextIndex = e.KeyCode == Keys.Down ? (currentIndex + 1) % rowCount : (currentIndex - 1 + rowCount) % rowCount; // Wraps around

                    // Set the current cell which will also select the row
                    dataGridView.CurrentCell = dataGridView.Rows[nextIndex].Cells[0];
                    dataGridView.Rows[nextIndex].Selected = true; // Ensure the row is selected
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
                    {
                        // Calculate the previous index, wrapping around to the last row if at the beginning
                        nextIndex = (currentIndex - 1 + rowCount) % rowCount;
                    }
                    else
                    {
                        // Calculate the next index, wrapping around to the first row if at the end
                        nextIndex = (currentIndex + 1) % rowCount;
                    }

                    dataGridView.ClearSelection();
                    dataGridView.CurrentCell = dataGridView.Rows[nextIndex].Cells[0];
                    dataGridView.Rows[nextIndex].Selected = true;

                    T selectedItem = (T)dataGridView.Rows[nextIndex].DataBoundItem;
                    filteredEntries.Add(selectedItem);
                    form.Close();
                }
            }
        };

        form.Load += (sender, e) =>
        {
            int totalRowHeight = dataGridView.Rows.GetRowsHeight(DataGridViewElementStates.Visible);
            int requiredHeight = totalRowHeight + dataGridView.ColumnHeadersHeight + 2 * dataGridView.Rows[0].Height + SystemInformation.HorizontalScrollBarHeight;
            form.Height = requiredHeight;

            // Check if required height exceeds the screen height
            // Make it fit the windows "working area" (excludes taskbar)
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
                // Center the form vertically if the required height fits within the screen height
                form.StartPosition = FormStartPosition.CenterScreen;
            }

            // Calculate and set column widths based on cell content and header text
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
                        int cellWidth = TextRenderer.MeasureText(row.Cells[column.Index].Value.ToString(), dataGridView.DefaultCellStyle.Font).Width;
                        maxWidth = Math.Max(maxWidth, cellWidth);
                    }
                }

                // Set column width to the maximum of header text width and cell content width
                column.Width = Math.Max(headerWidth, maxWidth) + 10;
                formMaxWidth += Math.Max(headerWidth, maxWidth) + 10;
            }
            form.Width = formMaxWidth + SystemInformation.VerticalScrollBarWidth + 43;
        };

        // Add controls to form and show it
        form.Controls.Add(dataGridView);
        form.Controls.Add(searchBox);
        searchBox.Select();
        form.ShowDialog();

        return filteredEntries;
    }
}
