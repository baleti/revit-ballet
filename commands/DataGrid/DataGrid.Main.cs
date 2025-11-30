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
    //  Font size tracking for Ctrl+/- zoom functionality
    // ──────────────────────────────────────────────────────────────
    private static float _currentFontSize = 9f; // Default font size
    private const float MIN_FONT_SIZE = 6f;
    private const float MAX_FONT_SIZE = 24f;

    // ──────────────────────────────────────────────────────────────
    //  Screen expansion state tracking
    // ──────────────────────────────────────────────────────────────
    private static int _currentScreenState = 0; // 0 = original, 1+ = expanded to N screens
    private static Rectangle _originalFormBounds;
    private static Screen _originalScreen;

    // ──────────────────────────────────────────────────────────────
    //  DataGrid sizing state tracking
    // ──────────────────────────────────────────────────────────────
    private static bool _initialSizingDone = false;

    // ──────────────────────────────────────────────────────────────
    //  Command name inference from call stack
    // ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Infers the Revit command name from the call stack by looking for methods
    /// with [CommandMethod] attributes or recognizable command patterns.
    /// </summary>
    private static string InferCommandNameFromCallStack()
    {
        try
        {
            var stackTrace = new System.Diagnostics.StackTrace();
            var frames = stackTrace.GetFrames();

            if (frames == null) return null;

            // Look through the call stack for methods that look like Revit commands
            for (int i = 0; i < frames.Length; i++)
            {
                var method = frames[i].GetMethod();
                if (method == null) continue;

                var declaringType = method.DeclaringType;
                if (declaringType == null) continue;

                // Skip our own DataGrid-related methods
                if (declaringType.Name == "CustomGUIs" ||
                    declaringType.Name.Contains("DataGrid"))
                    continue;

                // Look for methods that might be Revit commands:
                // 1. Class names ending with "Command"
                // 2. Method names matching common command patterns
                // 3. Namespace RevitBallet

                if (declaringType.Namespace == "RevitBallet" ||
                    declaringType.Name.EndsWith("Command") ||
                    method.Name.Contains("Execute"))
                {
                    // Try to extract command name from class name first
                    string className = declaringType.Name;

                    // Remove "Command" suffix if present
                    if (className.EndsWith("Command"))
                    {
                        className = className.Substring(0, className.Length - "Command".Length);
                    }

                    // Convert PascalCase to kebab-case
                    string commandName = ConvertToKebabCase(className);

                    // If we got a reasonable command name, return it
                    if (!string.IsNullOrWhiteSpace(commandName) &&
                        commandName.Length > 2 &&
                        commandName != "custom-gu-is")
                    {
                        return commandName;
                    }
                }
            }
        }
        catch
        {
            // If stack trace inspection fails, silently return null
        }

        return null;
    }

    /// <summary>
    /// Converts PascalCase to kebab-case (e.g., "SwitchView" -> "switch-view")
    /// </summary>
    private static string ConvertToKebabCase(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return input;

        // Insert hyphen before uppercase letters (except first character)
        var kebabCase = System.Text.RegularExpressions.Regex.Replace(input, "(?<!^)([A-Z])", "-$1");

        // Convert to lowercase
        return kebabCase.ToLower();
    }

    /// <summary>
    /// Formats column headers: replaces underscores with spaces, converts PascalCase to lowercase with spaces
    /// </summary>
    private static string FormatColumnHeader(string columnName)
    {
        if (string.IsNullOrEmpty(columnName))
            return columnName;

        var result = new System.Text.StringBuilder();

        for (int i = 0; i < columnName.Length; i++)
        {
            char c = columnName[i];

            // Replace underscores with spaces
            if (c == '_')
            {
                result.Append(' ');
            }
            // Add space before uppercase letters (except at start)
            else if (i > 0 && char.IsUpper(c) && !char.IsUpper(columnName[i - 1]))
            {
                result.Append(' ');
                result.Append(char.ToLower(c));
            }
            // Convert to lowercase
            else
            {
                result.Append(char.ToLower(c));
            }
        }

        return result.ToString();
    }

    // ──────────────────────────────────────────────────────────────
    //  Main DataGrid Method
    // ──────────────────────────────────────────────────────────────

    public static List<Dictionary<string, object>> DataGrid(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        bool spanAllScreens,
        List<int> initialSelectionIndices = null,
        System.Func<List<Dictionary<string, object>>, bool> onDeleteEntries = null,
        bool allowCreateFromSearch = false,
        string commandName = null)
    {
        if (entries == null || propertyNames == null || propertyNames.Count == 0)
            return new List<Dictionary<string, object>>();

        // Allow empty entries when allowCreateFromSearch is enabled (user can type new values)
        if (entries.Count == 0 && !allowCreateFromSearch)
            return new List<Dictionary<string, object>>();

        // Auto-infer command name from call stack if not provided
        if (string.IsNullOrWhiteSpace(commandName))
        {
            commandName = InferCommandNameFromCallStack();
        }

        // CRITICAL: Assign unique internal IDs to all entries for stable edit tracking
        // This ensures edits are correctly applied even when filters change row order
        AssignInternalIdsToEntries(entries);

        // Clear any previous cached data
        _cachedOriginalData = entries;
        _cachedFilteredData = entries;
        _searchIndexByColumn = null;
        _searchIndexAllColumns = null;
        _lastVisibleColumns.Clear();
        _lastColumnVisibilityFilter = "";

        // Reset sizing and edit mode state
        _initialSizingDone = false;
        _currentFontSize = 9f; // Reset to default font size
        ResetEditMode(); // Reset edit mode state from previous sessions
        ResetEditsAppliedFlag(); // Reset flag tracking if edits were applied

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
            BackColor = Color.White,
            KeyPreview = true // Enable form-level key handling for Shift+Arrow interception
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
            ScrollBars = ScrollBars.Both,
            ShowCellToolTips = false  // Disable tooltips on hover
        };

        // Disable built-in sorting
        grid.SortCompare += (sender, e) =>
        {
            e.Handled = true;
            e.SortResult = naturalComparer.Compare(e.CellValue1, e.CellValue2);
        };

        // Add columns (skip internal ID column)
        foreach (string col in propertyNames)
        {
            // Skip the internal tracking ID - never display it to users
            if (col == INTERNAL_ID_KEY)
                continue;

            var column = new DataGridViewTextBoxColumn
            {
                Name = col,
                HeaderText = FormatColumnHeader(col),
                DataPropertyName = col,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
            grid.Columns.Add(column);
        }

        // Search box with dropdown button container
        Panel searchPanel = new Panel { Dock = DockStyle.Top, Height = 20, Padding = new Padding(0) };

        TextBox searchBox = new TextBox
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Arial", 9, FontStyle.Regular)
        };

        // Dropdown button for search history
        Button dropdownButton = new Button
        {
            Dock = DockStyle.Right,
            Width = 20,
            Height = 20,
            Text = "▼",
            FlatStyle = FlatStyle.Flat,
            TabStop = false,
            Cursor = Cursors.Hand,
            ForeColor = Color.Black,
            BackColor = Color.White,
            Font = new Font("Arial", 8, FontStyle.Regular),
            Margin = new Padding(0),
            Padding = new Padding(0, 0, 0, 1)
        };

        dropdownButton.FlatAppearance.BorderSize = 1;
        dropdownButton.FlatAppearance.BorderColor = Color.FromArgb(127, 127, 127);
        dropdownButton.FlatAppearance.MouseOverBackColor = Color.White;
        dropdownButton.FlatAppearance.MouseDownBackColor = Color.White;

        // Add tooltip to dropdown button
        ToolTip dropdownTooltip = new ToolTip();
        dropdownTooltip.SetToolTip(dropdownButton, "Show search history (F4)");

        // Only show dropdown button if commandName is provided
        if (!string.IsNullOrWhiteSpace(commandName))
        {
            searchPanel.Controls.Add(searchBox);
            searchPanel.Controls.Add(dropdownButton);
        }
        else
        {
            // No command name, just use searchBox without button
            searchBox.Dock = DockStyle.Top;
        }

        // Track last search text for history recording
        string lastSearchTextForHistory = "";

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

        // Helper to update form title with edit mode indicator and entry counts
        Action UpdateFormTitle = () =>
        {
            string editModeIndicator = _isEditMode ? "[EDIT MODE] " : "";
            string pendingEditsInfo = _pendingCellEdits.Count > 0 ? $" ({_pendingCellEdits.Count} pending edits)" : "";
            form.Text = $"{editModeIndicator}Total Entries: {workingSet.Count} / {entries.Count}{pendingEditsInfo}";
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

            // Update form title (uses helper that includes edit mode indicator)
            UpdateFormTitle();

            // Auto-resize columns and form width only on initial load
            if (!_initialSizingDone)
            {
                if (grid.Columns.Count < 20)
                {
                    grid.AutoResizeColumns();
                }

                int reqWidth = grid.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                              + SystemInformation.VerticalScrollBarWidth + 50;
                form.Width = Math.Min(reqWidth, Screen.PrimaryScreen.WorkingArea.Width - 20);

                _initialSizingDone = true;
            }
        };

        // Initial draw
        _cachedFilteredData = workingSet;
        grid.RowCount = workingSet.Count;

        // Initial selection
        if (initialSelectionIndices != null && initialSelectionIndices.Count > 0)
        {
            int firstVisible = GetFirstVisibleColumnIndex();
            bool currentCellSet = false;
            int firstSelectedIndex = -1;

            foreach (int idx in initialSelectionIndices)
            {
                if (idx >= 0 && idx < grid.Rows.Count && firstVisible >= 0)
                {
                    grid.Rows[idx].Selected = true;

                    // Only set CurrentCell once, for the first valid selection
                    if (!currentCellSet)
                    {
                        grid.CurrentCell = grid.Rows[idx].Cells[firstVisible];
                        firstSelectedIndex = idx;
                        currentCellSet = true;
                    }
                }
            }

            // Scroll to make the first selected row visible
            if (firstSelectedIndex >= 0)
            {
                grid.FirstDisplayedScrollingRowIndex = firstSelectedIndex;
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

        // Record search query when focus leaves searchBox (user is done typing)
        searchBox.LostFocus += delegate
        {
            if (!string.IsNullOrWhiteSpace(commandName) &&
                !string.IsNullOrWhiteSpace(searchBox.Text) &&
                searchBox.Text.Trim() != lastSearchTextForHistory)
            {
                SearchQueryHistory.RecordQuery(commandName, searchBox.Text.Trim());
                lastSearchTextForHistory = searchBox.Text.Trim();
            }
        };

        // Dropdown button click - show search history
        if (!string.IsNullOrWhiteSpace(commandName))
        {
            dropdownButton.Click += delegate
            {
                ShowSearchHistoryDropdown(searchBox, commandName);
            };
        }

        // Column header click for sorting
        grid.ColumnHeaderMouseClick += (s, e) =>
        {
            string colName = grid.Columns[e.ColumnIndex].Name;
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

        // Track selected cells in edit mode
        grid.SelectionChanged += delegate
        {
            if (_isEditMode)
            {
                _selectedEditCells.Clear();
                foreach (DataGridViewCell cell in grid.SelectedCells)
                {
                    if (IsColumnEditable(grid.Columns[cell.ColumnIndex].Name))
                    {
                        _selectedEditCells.Add(cell);
                    }
                }
            }
        };

        // Set selection anchor on mouse down in edit mode
        grid.CellMouseDown += (s, e) =>
        {
            if (_isEditMode && e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string columnName = grid.Columns[e.ColumnIndex].Name;
                if (IsColumnEditable(columnName))
                {
                    // Only set anchor if not using Shift or Ctrl modifiers
                    if ((Control.ModifierKeys & (Keys.Shift | Keys.Control)) == Keys.None)
                    {
                        SetSelectionAnchor(grid.Rows[e.RowIndex].Cells[e.ColumnIndex]);
                    }
                }
            }
        };

        // Key handling - restore original behavior
        Action<KeyEventArgs, Control> HandleKeyDown = (e, sender) =>
        {
            // F2: Toggle edit mode / Show cell editor
            if (e.KeyCode == Keys.F2)
            {
                if (_isEditMode && _selectedEditCells.Count > 0)
                {
                    // Already in edit mode, show advanced edit dialog
                    ShowCellEditPrompt(grid);
                    UpdateFormTitle();
                }
                else
                {
                    // Toggle edit mode
                    ToggleEditMode(grid);
                    UpdateFormTitle();
                }
                e.Handled = true;
            }
            // Escape: Exit edit mode (if active) or cancel selection
            else if (e.KeyCode == Keys.Escape)
            {
                if (_isEditMode)
                {
                    // First press: Exit edit mode
                    ToggleEditMode(grid);
                    UpdateFormTitle();
                    e.Handled = true;
                }
                else
                {
                    // Not in edit mode: Cancel and close
                    selectedEntries.Clear();
                    form.Close();
                }
            }
            // Arrow keys in edit mode: Navigate cells
            else if (_isEditMode && sender == grid && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.Left || e.KeyCode == Keys.Right))
            {
                if (e.Shift)
                {
                    // Shift+Arrow: Extend selection (handled by form.KeyDown for proper interception)
                    // This block won't be reached because form intercepts Shift+Arrow first
                }
                else if (e.Control)
                {
                    // Ctrl+Arrow: Multi-select behavior
                    ExtendSelectionWithArrows(grid, e.KeyCode, false);
                    e.Handled = true;
                }
                else
                {
                    // Plain Arrow: Move cell
                    MoveCellWithArrows(grid, e.KeyCode);
                    e.Handled = true;
                }
            }
            // Ctrl+V: Paste clipboard
            else if (e.KeyCode == Keys.V && e.Control && _isEditMode)
            {
                HandleClipboardPaste(grid);
                UpdateFormTitle();
                e.Handled = true;
            }
            // Shift+Space: Select rows of selected cells (Excel-like)
            else if (e.KeyCode == Keys.Space && e.Shift && _isEditMode && sender == grid)
            {
                SelectRowsOfSelectedCells(grid);
                e.Handled = true;
            }
            // Ctrl+Space: Select columns of selected cells (Excel-like)
            else if (e.KeyCode == Keys.Space && e.Control && _isEditMode && sender == grid)
            {
                SelectColumnsOfSelectedCells(grid);
                e.Handled = true;
            }
            // Delete: Call onDeleteEntries callback
            else if (e.KeyCode == Keys.Delete && onDeleteEntries != null && sender == grid)
            {
                var entriesToDelete = new List<Dictionary<string, object>>();
                if (_isEditMode)
                {
                    // In edit mode: Delete entries corresponding to selected cells' rows
                    var rowIndices = _selectedEditCells.Select(cell => cell.RowIndex).Distinct();
                    foreach (int rowIndex in rowIndices)
                    {
                        if (rowIndex >= 0 && rowIndex < _cachedFilteredData.Count)
                            entriesToDelete.Add(_cachedFilteredData[rowIndex]);
                    }
                }
                else
                {
                    // Not in edit mode: Delete selected rows
                    foreach (DataGridViewRow row in grid.SelectedRows)
                    {
                        if (row.Index < _cachedFilteredData.Count)
                            entriesToDelete.Add(_cachedFilteredData[row.Index]);
                    }
                }

                if (entriesToDelete.Count > 0)
                {
                    bool deleted = onDeleteEntries(entriesToDelete);
                    if (deleted)
                    {
                        // Remove from working set and refresh
                        foreach (var entry in entriesToDelete)
                        {
                            workingSet.Remove(entry);
                            _cachedOriginalData.Remove(entry);
                        }
                        UpdateFilteredGrid();
                    }
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                FinishSelection();
            }
            // Spacebar: Select next/previous entry and close (when search box is empty)
            else if (e.KeyCode == Keys.Space && string.IsNullOrWhiteSpace(searchBox.Text) && !_isEditMode)
            {
                int count = grid.Rows.Count;
                if (count == 0) return;

                int cur = grid.CurrentRow?.Index ?? -1;
                int next = e.Shift ? (cur - 1 + count) % count : (cur + 1) % count;

                int firstVisible = GetFirstVisibleColumnIndex();
                if (firstVisible >= 0)
                {
                    grid.ClearSelection();
                    grid.CurrentCell = grid.Rows[next].Cells[firstVisible];
                    grid.Rows[next].Selected = true;
                }

                FinishSelection();
                e.Handled = true;
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
            else if (e.KeyCode == Keys.F4 && sender == searchBox && !string.IsNullOrWhiteSpace(commandName))
            {
                // Show search history dropdown
                ShowSearchHistoryDropdown(searchBox, commandName);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.S && e.Alt)
            {
                // Alt+S: Cycle through screen expansion states
                CycleScreenExpansion(form);
                e.Handled = true;
            }
            else if ((e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Add) && e.Control)
            {
                // Ctrl+ or Ctrl+Numpad+: Increase font size
                ResizeDataGrid(grid, searchBox, form, 1f);
                e.Handled = true;
            }
            else if ((e.KeyCode == Keys.OemMinus || e.KeyCode == Keys.Subtract) && e.Control)
            {
                // Ctrl- or Ctrl+Numpad-: Decrease font size
                ResizeDataGrid(grid, searchBox, form, -1f);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.E && e.Control)
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string storagePath = Path.Combine(appData, "revit-ballet", "runtime", "DataGrid-last-export-location");
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

        // Form-level KeyDown handler to intercept Shift+Arrow before grid processes it
        form.KeyDown += (s, e) =>
        {
            // Intercept Shift+Arrow in edit mode for Excel-like rectangular selection
            if (_isEditMode && e.Shift && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.Left || e.KeyCode == Keys.Right))
            {
                ExtendSelectionWithArrows(grid, e.KeyCode, true);
                e.Handled = true;
                e.SuppressKeyPress = true; // Prevent grid from processing this key
            }
        };

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
                _currentScreenState = Screen.AllScreens.Length;
                form.Width = Screen.AllScreens.Sum(s => s.WorkingArea.Width);
                form.Location = new Point(
                    Screen.AllScreens.Min(s => s.Bounds.X),
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2);
            }
            else
            {
                _currentScreenState = 0;
                int reqWidth = grid.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                              + SystemInformation.VerticalScrollBarWidth + 43;
                form.Width = Math.Min(reqWidth, Screen.PrimaryScreen.WorkingArea.Width - padding * 2);
                form.Location = new Point(
                    (Screen.PrimaryScreen.WorkingArea.Width - form.Width) / 2,
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2);
            }

            // Store original bounds and screen for cycling
            _originalFormBounds = form.Bounds;
            _originalScreen = Screen.FromControl(form);
        };

        // Add controls and show
        form.Controls.Add(grid);
        if (!string.IsNullOrWhiteSpace(commandName))
        {
            form.Controls.Add(searchPanel);
        }
        else
        {
            form.Controls.Add(searchBox);
        }
        searchBox.Select();
        form.ShowDialog();

        // Return modified entries if edits were made, otherwise return selected entries
        if (_modifiedEntries.Count > 0)
        {
            return _modifiedEntries.ToList();
        }

        return selectedEntries;
    }

    private static void ShowSearchHistoryDropdown(TextBox searchBox, string commandName)
    {
        // Get search history for this command
        var history = SearchQueryHistory.GetQueryHistory(commandName);

        if (history.Count == 0)
        {
            // No history to show
            return;
        }

        // Reverse to show most recent first
        history.Reverse();

        // Create a ListBox to show history
        ListBox historyList = new ListBox
        {
            Width = searchBox.Width,
            Height = Math.Min(history.Count * 20 + 4, 200), // Limit height
            BorderStyle = BorderStyle.FixedSingle,
            Font = searchBox.Font
        };

        // Add items to list
        foreach (string query in history)
        {
            historyList.Items.Add(query);
        }

        // Create a form to host the listbox (acts as dropdown)
        Form dropdown = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            StartPosition = FormStartPosition.Manual,
            ShowInTaskbar = false,
            TopMost = true,
            Width = historyList.Width,
            Height = historyList.Height
        };

        dropdown.Controls.Add(historyList);
        historyList.Dock = DockStyle.Fill;

        // Position below search box
        var searchBoxLocation = searchBox.PointToScreen(new Point(0, searchBox.Height));
        dropdown.Location = searchBoxLocation;

        // Handle selection
        historyList.Click += (s, e) =>
        {
            if (historyList.SelectedIndex >= 0)
            {
                searchBox.Text = historyList.SelectedItem.ToString();
                searchBox.SelectionStart = searchBox.Text.Length;
                dropdown.Close();
                searchBox.Focus();
            }
        };

        // Handle keyboard navigation
        historyList.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter && historyList.SelectedIndex >= 0)
            {
                searchBox.Text = historyList.SelectedItem.ToString();
                searchBox.SelectionStart = searchBox.Text.Length;
                dropdown.Close();
                searchBox.Focus();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                dropdown.Close();
                searchBox.Focus();
                e.Handled = true;
            }
        };

        // Close when focus is lost
        dropdown.Deactivate += (s, e) => dropdown.Close();

        // Select first item by default
        if (historyList.Items.Count > 0)
        {
            historyList.SelectedIndex = 0;
        }

        // Show dropdown
        dropdown.Show();
        historyList.Focus();
    }

    private static void CycleScreenExpansion(Form form)
    {
        Screen[] allScreens = Screen.AllScreens.OrderBy(s => s.Bounds.X).ToArray();
        Screen currentScreen = Screen.FromControl(form);

        // Set window state to normal first (in case it was maximized)
        form.WindowState = FormWindowState.Normal;

        if (allScreens.Length <= 1)
        {
            // Only one screen available, toggle between original size and maximized
            if (_currentScreenState == 0)
            {
                form.WindowState = FormWindowState.Maximized;
                _currentScreenState = 1;
            }
            else
            {
                form.Bounds = _originalFormBounds;
                _currentScreenState = 0;
            }
            return;
        }

        // Multi-screen cycling logic
        if (_currentScreenState == 0)
        {
            // State 0: Original -> State 1: Expand to screen on right
            int currentIndex = Array.IndexOf(allScreens, currentScreen);
            int rightIndex = (currentIndex + 1) % allScreens.Length;
            Screen rightScreen = allScreens[rightIndex];

            // Calculate combined bounds of current and right screen
            int leftX = Math.Min(currentScreen.WorkingArea.X, rightScreen.WorkingArea.X);
            int rightX = Math.Max(currentScreen.WorkingArea.Right, rightScreen.WorkingArea.Right);
            int topY = Math.Min(currentScreen.WorkingArea.Y, rightScreen.WorkingArea.Y);
            int maxHeight = Math.Max(currentScreen.WorkingArea.Height, rightScreen.WorkingArea.Height);

            form.Location = new Point(leftX, topY);
            form.Width = rightX - leftX;
            form.Height = Math.Min(form.Height, maxHeight);
            _currentScreenState = 1;
        }
        else if (_currentScreenState == 1)
        {
            // State 1: Right expansion -> State 2: Expand to screen on left
            int currentIndex = Array.IndexOf(allScreens, _originalScreen);
            int leftIndex = currentIndex == 0 ? allScreens.Length - 1 : currentIndex - 1;
            Screen leftScreen = allScreens[leftIndex];

            // Calculate combined bounds of original and left screen
            int leftX = Math.Min(_originalScreen.WorkingArea.X, leftScreen.WorkingArea.X);
            int rightX = Math.Max(_originalScreen.WorkingArea.Right, leftScreen.WorkingArea.Right);
            int topY = Math.Min(_originalScreen.WorkingArea.Y, leftScreen.WorkingArea.Y);
            int maxHeight = Math.Max(_originalScreen.WorkingArea.Height, leftScreen.WorkingArea.Height);

            form.Location = new Point(leftX, topY);
            form.Width = rightX - leftX;
            form.Height = Math.Min(form.Height, maxHeight);
            _currentScreenState = 2;
        }
        else if (_currentScreenState < allScreens.Length)
        {
            // Continue expanding to include more screens until all are covered
            _currentScreenState++;

            // Calculate bounds for first N screens
            var screensToInclude = allScreens.Take(_currentScreenState).ToArray();
            int leftX = screensToInclude.Min(s => s.WorkingArea.X);
            int rightX = screensToInclude.Max(s => s.WorkingArea.Right);
            int topY = screensToInclude.Min(s => s.WorkingArea.Y);
            int maxHeight = screensToInclude.Max(s => s.WorkingArea.Height);

            form.Location = new Point(leftX, topY);
            form.Width = rightX - leftX;
            form.Height = Math.Min(form.Height, maxHeight);
        }
        else
        {
            // All screens covered -> Reset to original
            form.Bounds = _originalFormBounds;
            _currentScreenState = 0;
        }

        // Bring form to front
        form.BringToFront();
        form.Focus();
    }

    /// <summary>
    /// Resizes the DataGrid font and row height based on Ctrl+/- zoom commands
    /// </summary>
    private static void ResizeDataGrid(DataGridView grid, TextBox searchBox, Form form, float fontSizeDelta)
    {
        // Calculate new font size with bounds checking
        float newFontSize = Math.Max(MIN_FONT_SIZE, Math.Min(MAX_FONT_SIZE, _currentFontSize + fontSizeDelta));

        if (newFontSize == _currentFontSize)
        {
            // Already at min or max, no change needed
            return;
        }

        _currentFontSize = newFontSize;

        // Update grid font
        grid.Font = new Font(grid.Font.FontFamily, _currentFontSize);

        // Update search box font to match
        searchBox.Font = new Font(searchBox.Font.FontFamily, _currentFontSize);

        // Update search box container (panel or direct parent) height to accommodate new font
        Control searchParent = searchBox.Parent;
        if (searchParent != null)
        {
            int newSearchHeight = (int)Math.Ceiling(_currentFontSize * 2.0f) + 6;
            searchParent.Height = newSearchHeight;

            // If there's a dropdown button, update its height as well
            foreach (Control child in searchParent.Controls)
            {
                if (child is Button)
                {
                    child.Height = newSearchHeight;
                }
            }
        }

        // Calculate row height based on font size
        // Use a multiplier to ensure adequate spacing (approx 2x font size)
        int newRowHeight = (int)Math.Ceiling(_currentFontSize * 2.0f);
        grid.RowTemplate.Height = newRowHeight;

        // Update existing rows to new height
        foreach (DataGridViewRow row in grid.Rows)
        {
            row.Height = newRowHeight;
        }

        // Update column header height based on font size
        // Use a multiplier to ensure adequate spacing (approx 2.2x font size + padding)
        int newHeaderHeight = (int)Math.Ceiling(_currentFontSize * 2.2f) + 4;
        grid.ColumnHeadersHeight = newHeaderHeight;

        // Auto-resize columns to fit new font size
        grid.AutoResizeColumns();

        // Adjust form height to accommodate new row heights if there's room
        int padding = 20;
        int rowsHeight = grid.Rows.GetRowsHeight(DataGridViewElementStates.Visible);
        int requiredHeight = rowsHeight + grid.ColumnHeadersHeight +
                            2 * grid.RowTemplate.Height +
                            SystemInformation.HorizontalScrollBarHeight + 30;

        int availableHeight = Screen.FromControl(form).WorkingArea.Height - padding * 2;
        int newFormHeight = Math.Min(requiredHeight, availableHeight);

        // Only adjust height if it's different from current height
        if (newFormHeight != form.Height)
        {
            form.Height = newFormHeight;

            // Re-center vertically after height change
            Screen currentScreen = Screen.FromControl(form);
            form.Location = new Point(
                form.Location.X,
                (currentScreen.WorkingArea.Height - form.Height) / 2 + currentScreen.WorkingArea.Top
            );
        }

        // Force grid to refresh
        grid.Invalidate();
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
