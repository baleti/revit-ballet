// EditSelectionSets.cs  –  Revit 2024 / C# 7.3 compatible
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using WinForms = System.Windows.Forms;

namespace RevitAddin
{
    // ─────────────────────────────────────────────────────────────
    //  Selection Set Info Class
    // ─────────────────────────────────────────────────────────────
    public class SelectionSetInfo : ICloneable
    {
        private static int _nextId = 1;
        
        public int UniqueId { get; set; }
        public SelectionFilterElement FilterElement { get; set; }
        public string Name { get; set; }
        public string OriginalName { get; set; }
        public int ElementCount { get; set; }
        public bool NameChanged => Name != OriginalName;
        public bool MarkedForDeletion { get; set; }
        public bool IsNew { get; set; } // Indicates this is a duplicated item

        public SelectionSetInfo()
        {
            UniqueId = _nextId++;
        }

        public object Clone()
        {
            var clone = new SelectionSetInfo
            {
                FilterElement = this.FilterElement,
                Name = this.Name,
                OriginalName = this.OriginalName,
                ElementCount = this.ElementCount,
                MarkedForDeletion = this.MarkedForDeletion,
                IsNew = this.IsNew
            };
            // Manually set the UniqueId to preserve it
            clone.UniqueId = this.UniqueId;
            return clone;
        }
    }

    // ─────────────────────────────────────────────────────────────
    //  Undo/Redo System
    // ─────────────────────────────────────────────────────────────
    internal class UndoRedoManager
    {
        private readonly Stack<List<SelectionSetInfo>> _undoStack = new Stack<List<SelectionSetInfo>>();
        private readonly Stack<List<SelectionSetInfo>> _redoStack = new Stack<List<SelectionSetInfo>>();

        public bool CanUndo => _undoStack.Count > 0;
        public bool CanRedo => _redoStack.Count > 0;

        public void SaveState(List<SelectionSetInfo> state)
        {
            var stateCopy = state.Select(s => (SelectionSetInfo)s.Clone()).ToList();
            _undoStack.Push(stateCopy);
            _redoStack.Clear(); // Clear redo stack when new action is performed
        }

        public List<SelectionSetInfo> Undo(List<SelectionSetInfo> currentState)
        {
            if (!CanUndo) return null;

            // Save current state to redo stack before returning previous state
            var stateCopy = currentState.Select(s => (SelectionSetInfo)s.Clone()).ToList();
            _redoStack.Push(stateCopy);

            return _undoStack.Pop().Select(s => (SelectionSetInfo)s.Clone()).ToList();
        }

        public List<SelectionSetInfo> Redo()
        {
            if (!CanRedo) return null;

            var state = _redoStack.Pop();
            _undoStack.Push(state.Select(s => (SelectionSetInfo)s.Clone()).ToList());
            return state.Select(s => (SelectionSetInfo)s.Clone()).ToList();
        }
    }
    
    // ─────────────────────────────────────────────────────────────
    //  WinForm
    // ─────────────────────────────────────────────────────────────
    internal class SelectionSetEditForm : WinForms.Form
    {
        private readonly WinForms.DataGridView _dataGrid;
        private readonly WinForms.Button _duplicateButton;
        private readonly List<SelectionSetInfo> _selectionSets;
        private readonly UndoRedoManager _undoRedoManager = new UndoRedoManager();
        private bool _isEditingCell = false;
        private bool _cancellingEdit = false;
        private bool _isUndoRedoOperation = false;

        public List<SelectionSetInfo> NewSelectionSets { get; private set; } = 
            new List<SelectionSetInfo>();
        public List<(SelectionFilterElement element, string newName)> RenamedSelectionSets { get; private set; } = 
            new List<(SelectionFilterElement, string)>();
        public List<SelectionFilterElement> DeletedSelectionSets { get; private set; } = 
            new List<SelectionFilterElement>();

        public SelectionSetEditForm(List<SelectionSetInfo> selectionSets)
        {
            _selectionSets = selectionSets;
            
            Text            = "Edit Selection Sets";
            FormBorderStyle = WinForms.FormBorderStyle.Sizable;
            StartPosition   = WinForms.FormStartPosition.CenterScreen;
            MinimumSize = new System.Drawing.Size(500, 300);

            // Calculate optimal window height
            var screenHeight = WinForms.Screen.PrimaryScreen.WorkingArea.Height;
            var maxWindowHeight = screenHeight - 100; // Leave 100px padding
            var desiredHeight = Math.Max(400, Math.Min(700, selectionSets.Count * 22 + 120));
            var windowHeight = Math.Min(maxWindowHeight, desiredHeight);

            ClientSize = new System.Drawing.Size(600, windowHeight);

            // Create DataGridView
            _dataGrid = new WinForms.DataGridView
            {
                Location = new System.Drawing.Point(12, 12),
                Size = new System.Drawing.Size(576, windowHeight - 60),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | 
                        WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                BackgroundColor = SystemColors.Window,
                BorderStyle = WinForms.BorderStyle.Fixed3D,
                RowTemplate = { Height = 20 },
                EditMode = WinForms.DataGridViewEditMode.EditProgrammatically,
                StandardTab = true, // Tab moves focus out of grid
                TabStop = true
            };

            // Handle keyboard events
            _dataGrid.KeyDown += DataGrid_KeyDown;
            _dataGrid.CellBeginEdit += DataGrid_CellBeginEdit;
            _dataGrid.CellEndEdit += DataGrid_CellEndEdit;
            _dataGrid.CellValidating += DataGrid_CellValidating;
            _dataGrid.RowPrePaint += DataGrid_RowPrePaint;

            // Add columns
            var nameColumn = new WinForms.DataGridViewTextBoxColumn
            {
                HeaderText = "Selection Set Name",
                Name = "Name",
                DataPropertyName = "Name",
                Width = 400,
                ReadOnly = false,
                AutoSizeMode = WinForms.DataGridViewAutoSizeColumnMode.Fill
            };
            
            var countColumn = new WinForms.DataGridViewTextBoxColumn
            {
                HeaderText = "Elements",
                Name = "ElementCount",
                DataPropertyName = "ElementCount",
                Width = 80,
                ReadOnly = true,
                AutoSizeMode = WinForms.DataGridViewAutoSizeColumnMode.None
            };

            _dataGrid.Columns.AddRange(new WinForms.DataGridViewColumn[] 
                { nameColumn, countColumn });

            // Bind data
            _dataGrid.DataSource = _selectionSets;

            // Duplicate Selected button
            _duplicateButton = new WinForms.Button
            {
                Text = "Duplicate Selected",
                Location = new System.Drawing.Point(12, windowHeight - 38),
                Size = new System.Drawing.Size(120, 25),
                Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left
            };
            _duplicateButton.Click += DuplicateButton_Click;

            // OK and Cancel buttons
            var ok = new WinForms.Button 
            { 
                Text = "OK", 
                Location = new System.Drawing.Point(396, windowHeight - 38), 
                Size = new System.Drawing.Size(80, 25), 
                DialogResult = WinForms.DialogResult.OK,
                Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right
            };
            
            var cancel = new WinForms.Button 
            { 
                Text = "Cancel", 
                Location = new System.Drawing.Point(486, windowHeight - 38), 
                Size = new Size(80, 25), 
                DialogResult = WinForms.DialogResult.Cancel,
                Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right
            };

            AcceptButton = ok;
            CancelButton = cancel;

            // Set tab order
            _dataGrid.TabIndex = 0;
            _duplicateButton.TabIndex = 1;
            ok.TabIndex = 2;
            cancel.TabIndex = 3;

            Controls.AddRange(new WinForms.Control[]
            {
                _dataGrid, _duplicateButton, ok, cancel
            });

            // Save initial state for undo
            _undoRedoManager.SaveState(_selectionSets);
        }

        private void SaveUndoState()
        {
            if (!_isUndoRedoOperation)
            {
                _undoRedoManager.SaveState(_selectionSets);
            }
        }

        private void PerformUndo()
        {
            if (!_undoRedoManager.CanUndo) return;
            
            _isUndoRedoOperation = true;
            var previousState = _undoRedoManager.Undo(_selectionSets);
            if (previousState != null)
            {
                _selectionSets.Clear();
                _selectionSets.AddRange(previousState.Select(s => (SelectionSetInfo)s.Clone()));
                RefreshDataGrid();
            }
            _isUndoRedoOperation = false;
        }

        private void PerformRedo()
        {
            _isUndoRedoOperation = true;
            var nextState = _undoRedoManager.Redo();
            if (nextState != null)
            {
                _selectionSets.Clear();
                _selectionSets.AddRange(nextState.Select(s => (SelectionSetInfo)s.Clone()));
                RefreshDataGrid();
            }
            _isUndoRedoOperation = false;
        }

        private void RefreshDataGrid()
        {
            var selectedIndices = _dataGrid.SelectedRows.Cast<WinForms.DataGridViewRow>()
                .Select(r => r.Index).ToList();

            _dataGrid.DataSource = null;
            _dataGrid.DataSource = _selectionSets;

            // Re-apply columns
            _dataGrid.Columns["Name"].AutoSizeMode = WinForms.DataGridViewAutoSizeColumnMode.Fill;
            _dataGrid.Columns["ElementCount"].Width = 80;

            // Restore selection
            foreach (var index in selectedIndices)
            {
                if (index < _dataGrid.Rows.Count)
                {
                    _dataGrid.Rows[index].Selected = true;
                }
            }
        }

        private void DuplicateButton_Click(object sender, EventArgs e)
        {
            DuplicateSelectedRows();
        }

        private void DuplicateSelectedRows()
        {
            var selectedRows = _dataGrid.SelectedRows.Cast<WinForms.DataGridViewRow>()
                .Where(r => !(r.DataBoundItem as SelectionSetInfo)?.MarkedForDeletion ?? false)
                .OrderBy(r => r.Index)
                .ToList();

            if (selectedRows.Count == 0) return;

            SaveUndoState();

            // Get all existing names (excluding items marked for deletion)
            var existingNames = _selectionSets
                .Where(s => !s.MarkedForDeletion)
                .Select(s => s.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var itemsToInsert = new List<(int index, SelectionSetInfo item)>();

            foreach (var row in selectedRows)
            {
                if (row.DataBoundItem is SelectionSetInfo originalInfo)
                {
                    // Generate unique name
                    string newName = GetNextAvailableName(originalInfo.Name, existingNames);
                    existingNames.Add(newName);

                    // Create new info
                    var newInfo = new SelectionSetInfo
                    {
                        FilterElement = originalInfo.FilterElement,
                        Name = newName,
                        OriginalName = newName,
                        ElementCount = originalInfo.ElementCount,
                        MarkedForDeletion = false,
                        IsNew = true
                    };

                    // Store the item with its insertion index (right after the original)
                    itemsToInsert.Add((row.Index + 1, newInfo));
                }
            }

            // Insert items in reverse order to maintain correct positions
            foreach (var (insertIndex, item) in itemsToInsert.OrderByDescending(x => x.index))
            {
                var adjustedIndex = Math.Min(insertIndex + itemsToInsert.Count(x => x.index < insertIndex), _selectionSets.Count);
                _selectionSets.Insert(adjustedIndex, item);
            }

            RefreshDataGrid();

            // Select the new rows
            _dataGrid.ClearSelection();
            foreach (var (_, item) in itemsToInsert)
            {
                var rowIndex = _selectionSets.IndexOf(item);
                if (rowIndex >= 0)
                {
                    _dataGrid.Rows[rowIndex].Selected = true;
                }
            }
        }

        private string GetNextAvailableName(string baseName, HashSet<string> usedNames)
        {
            var numberSuffix = new Regex(@"\s+(\d+)$", RegexOptions.Compiled);
            var match = numberSuffix.Match(baseName);
            
            string nameWithoutNumber;
            int startNumber;

            if (match.Success)
            {
                nameWithoutNumber = baseName.Substring(0, match.Index);
                startNumber = int.Parse(match.Groups[1].Value) + 1;
            }
            else
            {
                nameWithoutNumber = baseName;
                startNumber = 2;
            }

            int number = startNumber;
            string candidate = $"{nameWithoutNumber} {number}";
            
            while (usedNames.Contains(candidate, StringComparer.OrdinalIgnoreCase))
            {
                number++;
                candidate = $"{nameWithoutNumber} {number}";
            }

            return candidate;
        }

        private void DataGrid_RowPrePaint(object sender, WinForms.DataGridViewRowPrePaintEventArgs e)
        {
            // Style rows based on their state
            if (e.RowIndex >= 0 && e.RowIndex < _selectionSets.Count)
            {
                var info = _selectionSets[e.RowIndex];
                var row = _dataGrid.Rows[e.RowIndex];
                
                if (info.MarkedForDeletion)
                {
                    // Light red background with strikethrough
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 220, 220);
                    row.DefaultCellStyle.ForeColor = System.Drawing.Color.DarkGray;
                    row.DefaultCellStyle.Font = new Font(_dataGrid.Font, FontStyle.Strikeout);
                }
                else if (info.IsNew)
                {
                    // Light green background for new/duplicated items
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(220, 255, 220);
                    row.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
                    row.DefaultCellStyle.Font = _dataGrid.Font;
                }
                else if (info.NameChanged)
                {
                    // Light blue background for renamed items
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(220, 230, 255);
                    row.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
                    row.DefaultCellStyle.Font = _dataGrid.Font;
                }
                else
                {
                    // Default style
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.White;
                    row.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
                    row.DefaultCellStyle.Font = _dataGrid.Font;
                }
            }
        }

        private void DataGrid_KeyDown(object sender, WinForms.KeyEventArgs e)
        {
            // Handle ESC key for cell editing first
            if (e.KeyCode == WinForms.Keys.Escape && _isEditingCell)
            {
                // Cancel edit only
                _cancellingEdit = true;
                _dataGrid.CancelEdit();
                _dataGrid.EndEdit();
                _isEditingCell = false;
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }
            
            if (e.Control && e.KeyCode == WinForms.Keys.Z && !e.Shift && !_isEditingCell)
            {
                // Ctrl+Z - Undo
                PerformUndo();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if ((e.Control && e.KeyCode == WinForms.Keys.Y) || 
                     (e.Control && e.Shift && e.KeyCode == WinForms.Keys.Z) && !_isEditingCell)
            {
                // Ctrl+Y or Ctrl+Shift+Z - Redo
                PerformRedo();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == WinForms.Keys.F2 && !_isEditingCell)
            {
                HandleF2Edit();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == WinForms.Keys.Delete && !_isEditingCell)
            {
                ToggleDeleteOnSelectedRows();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == WinForms.Keys.D && !_isEditingCell)
            {
                // Ctrl+D to duplicate
                DuplicateSelectedRows();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == WinForms.Keys.Enter && !_isEditingCell)
            {
                // Act as OK button (only if not editing)
                e.Handled = true;
                e.SuppressKeyPress = true;
                DialogResult = WinForms.DialogResult.OK;
                Close();
            }
            else if (e.Control && e.KeyCode == WinForms.Keys.A && !_isEditingCell)
            {
                // Select all rows (only if not editing)
                e.Handled = true;
                e.SuppressKeyPress = true;
                _dataGrid.SelectAll();
            }
            else if (e.KeyCode == WinForms.Keys.Escape && !_isEditingCell)
            {
                // Close dialog
                DialogResult = WinForms.DialogResult.Cancel;
                Close();
            }
        }

        private string ShowInputDialog(string prompt, string title, string defaultValue)
        {
            var form = new WinForms.Form
            {
                Text = title,
                FormBorderStyle = WinForms.FormBorderStyle.FixedDialog,
                StartPosition = WinForms.FormStartPosition.CenterParent,
                ClientSize = new System.Drawing.Size(400, 150),
                MaximizeBox = false,
                MinimizeBox = false
            };

            var label = new WinForms.Label
            {
                Text = prompt,
                Location = new System.Drawing.Point(12, 20),
                Size = new System.Drawing.Size(376, 40),
                AutoSize = false
            };

            var textBox = new WinForms.TextBox
            {
                Text = defaultValue,
                Location = new System.Drawing.Point(12, 65),
                Size = new System.Drawing.Size(376, 20)
            };

            var okButton = new WinForms.Button
            {
                Text = "OK",
                DialogResult = WinForms.DialogResult.OK,
                Location = new System.Drawing.Point(232, 95),
                Size = new System.Drawing.Size(75, 23)
            };

            var cancelButton = new WinForms.Button
            {
                Text = "Cancel",
                DialogResult = WinForms.DialogResult.Cancel,
                Location = new System.Drawing.Point(313, 95),
                Size = new Size(75, 23)
            };

            form.Controls.AddRange(new WinForms.Control[] { label, textBox, okButton, cancelButton });
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            textBox.SelectAll();
            textBox.Focus();

            return form.ShowDialog() == WinForms.DialogResult.OK ? textBox.Text : string.Empty;
        }

        private void HandleF2Edit()
        {
            var selectedRows = _dataGrid.SelectedRows.Cast<WinForms.DataGridViewRow>()
                .Where(r => !(r.DataBoundItem as SelectionSetInfo)?.MarkedForDeletion ?? false)
                .OrderBy(r => r.Index)
                .ToList();

            if (selectedRows.Count == 0) return;

            if (selectedRows.Count == 1)
            {
                // Single row - just edit it
                _dataGrid.CurrentCell = selectedRows[0].Cells[0]; // Name column (now at index 0)
                _dataGrid.BeginEdit(true);
            }
            else
            {
                // Multiple rows - show dialog for batch rename
                var firstInfo = selectedRows[0].DataBoundItem as SelectionSetInfo;
                if (firstInfo == null) return;

                string baseName = ShowInputDialog(
                    $"Enter new base name for {selectedRows.Count} selected items.\n" +
                    "Numbers will be appended automatically (2, 3, 4...).",
                    "Batch Rename",
                    firstInfo.Name);

                if (!string.IsNullOrEmpty(baseName))
                {
                    SaveUndoState();
                    baseName = baseName.Trim();
                    
                    // Get all existing names for duplicate checking (excluding deleted items)
                    var existingNames = _selectionSets
                        .Where((s, idx) => !s.MarkedForDeletion && !selectedRows.Any(r => r.Index == idx))
                        .Select(s => s.Name)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);

                    // Apply names with numbering
                    bool first = true;
                    int number = 2;
                    
                    foreach (var row in selectedRows)
                    {
                        if (row.DataBoundItem is SelectionSetInfo info)
                        {
                            string newName;
                            if (first)
                            {
                                newName = baseName;
                                first = false;
                            }
                            else
                            {
                                // Find next available number
                                do
                                {
                                    newName = $"{baseName} {number}";
                                    number++;
                                } while (existingNames.Contains(newName) || 
                                         selectedRows.Take(selectedRows.IndexOf(row))
                                             .Any(r => (r.DataBoundItem as SelectionSetInfo)?.Name == newName));
                            }
                            
                            // Check if this name already exists in the list
                            // Only mark for deletion if it's a different item (by UniqueId) and not being renamed in this batch
                            var existingItem = _selectionSets.FirstOrDefault(s => 
                                s.UniqueId != info.UniqueId && 
                                !s.MarkedForDeletion && 
                                string.Equals(s.Name, newName, StringComparison.OrdinalIgnoreCase) &&
                                !selectedRows.Any(r => r.DataBoundItem is SelectionSetInfo rowInfo && rowInfo.UniqueId == s.UniqueId));
                            
                            if (existingItem != null)
                            {
                                // Mark the existing item for deletion (overwrite)
                                existingItem.MarkedForDeletion = true;
                            }
                        }
                    }
                    
                    _dataGrid.Refresh();
                }
            }
        }

        private void ToggleDeleteOnSelectedRows()
        {
            var selectedRows = _dataGrid.SelectedRows;
            if (selectedRows.Count == 0) return;

            SaveUndoState();

            // Create a list to track items to remove
            var itemsToRemove = new List<SelectionSetInfo>();

            foreach (WinForms.DataGridViewRow row in selectedRows)
            {
                if (row.DataBoundItem is SelectionSetInfo info)
                {
                    // Don't allow deleting new items that haven't been saved yet
                    if (info.IsNew)
                    {
                        // Instead, remove them from the list entirely
                        itemsToRemove.Add(info);
                    }
                    else
                    {
                        info.MarkedForDeletion = !info.MarkedForDeletion;
                    }
                }
            }

            // Remove items after iteration to avoid modifying collection during enumeration
            foreach (var item in itemsToRemove)
            {
                _selectionSets.Remove(item);
            }
            
            RefreshDataGrid();
        }

        private void DataGrid_CellValidating(object sender, WinForms.DataGridViewCellValidatingEventArgs e)
        {
            // Validate name column
            if (e.ColumnIndex == 0) // Name column
            {
                string newName = e.FormattedValue.ToString().Trim();
                
                if (string.IsNullOrEmpty(newName))
                {
                    e.Cancel = true;
                    WinForms.MessageBox.Show("Selection set name cannot be empty.", 
                        "Invalid Name", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                }
                else
                {
                    var currentInfo = _selectionSets[e.RowIndex];
                    
                    // Check for duplicate names (excluding current row and deleted items)
                    SelectionSetInfo duplicateItem = null;
                    for (int i = 0; i < _selectionSets.Count; i++)
                    {
                        if (_selectionSets[i].UniqueId != currentInfo.UniqueId && 
                            !_selectionSets[i].MarkedForDeletion &&
                            string.Equals(_selectionSets[i].Name, newName, StringComparison.OrdinalIgnoreCase))
                        {
                            duplicateItem = _selectionSets[i];
                            break;
                        }
                    }
                    
                    if (duplicateItem != null)
                    {
                        // Mark the duplicate for deletion
                        duplicateItem.MarkedForDeletion = true;
                        // Allow the rename to proceed
                    }
                }
            }
        }

        private void DataGrid_CellBeginEdit(object sender, WinForms.DataGridViewCellCancelEventArgs e)
        {
            // Prevent editing of deleted items
            if (e.RowIndex >= 0 && e.RowIndex < _selectionSets.Count)
            {
                var info = _selectionSets[e.RowIndex];
                if (info.MarkedForDeletion)
                {
                    e.Cancel = true;
                    return;
                }
            }
            
            _isEditingCell = true;
        }
        
        private void DataGrid_CellEndEdit(object sender, WinForms.DataGridViewCellEventArgs e)
        {
            _isEditingCell = false;
            
            if (_cancellingEdit)
            {
                _cancellingEdit = false;
                return;
            }
            
            // Save undo state after successful edit
            if (e.ColumnIndex == 0 && e.RowIndex >= 0) // Name column
            {
                var info = _selectionSets[e.RowIndex];
                if (info != null && info.Name != info.OriginalName)
                {
                    SaveUndoState();
                }
                _dataGrid.Refresh();
            }
        }

        protected override bool ProcessCmdKey(ref WinForms.Message msg, WinForms.Keys keyData)
        {
            // Prevent ESC from closing form when editing
            if (keyData == WinForms.Keys.Escape && _isEditingCell)
            {
                _cancellingEdit = true;
                _dataGrid.CancelEdit();
                _dataGrid.EndEdit();
                _isEditingCell = false;
                return true;
            }
            
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);
            if (DialogResult != WinForms.DialogResult.OK) return;

            // Collect new selection sets to create
            NewSelectionSets = _selectionSets
                .Where(s => s.IsNew && !s.MarkedForDeletion)
                .ToList();

            // Collect selection sets to rename (excluding deleted and new ones)
            RenamedSelectionSets = _selectionSets
                .Where(s => !s.IsNew && s.NameChanged && !s.MarkedForDeletion)
                .Select(s => (s.FilterElement, s.Name))
                .ToList();

            // Collect selection sets to delete
            DeletedSelectionSets = _selectionSets
                .Where(s => !s.IsNew && s.MarkedForDeletion)
                .Select(s => s.FilterElement)
                .ToList();

            if (NewSelectionSets.Count == 0 && RenamedSelectionSets.Count == 0 && DeletedSelectionSets.Count == 0)
            {
                WinForms.MessageBox.Show("No changes were made.",
                    "No Changes", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                e.Cancel = true;
                return;
            }
        }
    }

    // ─────────────────────────────────────────────────────────────
    //  Core editing helper
    // ─────────────────────────────────────────────────────────────
    internal static class SelectionSetEditor
    {
        public static List<SelectionFilterElement> CreateNewSelectionSets(
            Document doc,
            IEnumerable<SelectionSetInfo> newSetInfos)
        {
            var createdSelectionSets = new List<SelectionFilterElement>();

            foreach (var info in newSetInfos)
            {
                // Get the element IDs from the original selection set
                ICollection<ElementId> elementIds = info.FilterElement.GetElementIds();

                // Create new selection set
                SelectionFilterElement newSet = SelectionFilterElement.Create(doc, info.Name);
                
                // Add the same elements to the new selection set
                newSet.SetElementIds(elementIds);

                createdSelectionSets.Add(newSet);
            }

            return createdSelectionSets;
        }

        public static void RenameSelectionSets(
            Document doc,
            IEnumerable<(SelectionFilterElement element, string newName)> renamePairs)
        {
            foreach (var (element, newName) in renamePairs)
            {
                element.Name = newName;
            }
        }

        public static void DeleteSelectionSets(
            Document doc,
            IEnumerable<SelectionFilterElement> selectionSets)
        {
            foreach (var set in selectionSets)
            {
                doc.Delete(set.Id);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────
    //  Revit external command
    // ─────────────────────────────────────────────────────────────
    [Transaction(TransactionMode.Manual)]
    public class EditSelectionSets : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uiDoc.Document;

            // Get all selection sets
            var selectionSets = new FilteredElementCollector(doc)
                .OfClass(typeof(SelectionFilterElement))
                .Cast<SelectionFilterElement>()
                .OrderBy(s => s.Name)
                .ToList();

            if (selectionSets.Count == 0)
            {
                TaskDialog.Show("Edit Selection Sets", 
                    "No selection sets found in the current document.");
                return Result.Cancelled;
            }

            // Prepare selection set info
            var selectionSetInfos = new List<SelectionSetInfo>();
            foreach (var set in selectionSets)
            {
                selectionSetInfos.Add(new SelectionSetInfo
                {
                    FilterElement = set,
                    Name = set.Name,
                    OriginalName = set.Name,
                    ElementCount = set.GetElementIds().Count,
                    MarkedForDeletion = false,
                    IsNew = false
                });
            }

            // Show dialog
            List<SelectionSetInfo> newSets;
            List<(SelectionFilterElement, string)> renamedSets;
            List<SelectionFilterElement> deletedSets;
            
            using (var dlg = new SelectionSetEditForm(selectionSetInfos))
            {
                if (dlg.ShowDialog() != WinForms.DialogResult.OK) 
                    return Result.Cancelled;
                    
                newSets = dlg.NewSelectionSets;
                renamedSets = dlg.RenamedSelectionSets;
                deletedSets = dlg.DeletedSelectionSets;
            }

            // Perform operations
            using (var t = new Transaction(doc, "Edit Selection Sets"))
            {
                t.Start();
                
                try
                {
                    // Delete first
                    if (deletedSets.Count > 0)
                    {
                        SelectionSetEditor.DeleteSelectionSets(doc, deletedSets);
                    }
                    
                    // Apply renames
                    if (renamedSets.Count > 0)
                    {
                        SelectionSetEditor.RenameSelectionSets(doc, renamedSets);
                    }
                    
                    // Create new selection sets
                    List<SelectionFilterElement> createdSets = null;
                    if (newSets.Count > 0)
                    {
                        createdSets = SelectionSetEditor.CreateNewSelectionSets(doc, newSets);
                    }
                    
                    t.Commit();
                    
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    message = ex.Message;
                    return Result.Failed;
                }
            }
            
            return Result.Succeeded;
        }
    }
}
