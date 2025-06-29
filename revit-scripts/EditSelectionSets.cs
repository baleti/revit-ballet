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
    public class SelectionSetInfo
    {
        public SelectionFilterElement FilterElement { get; set; }
        public bool Selected { get; set; }
        public string Name { get; set; }
        public string OriginalName { get; set; }
        public int ElementCount { get; set; }
        public bool NameChanged => Name != OriginalName;
        public bool MarkedForDeletion { get; set; }
    }
    
    // ─────────────────────────────────────────────────────────────
    //  WinForm
    // ─────────────────────────────────────────────────────────────
    internal class SelectionSetEditForm : WinForms.Form
    {
        private readonly WinForms.DataGridView _dataGrid;
        private readonly WinForms.NumericUpDown _numCopies;
        private readonly List<SelectionSetInfo> _selectionSets;
        private bool _isEditingCell = false;
        private bool _cancellingEdit = false;

        public List<SelectionFilterElement> SelectedSelectionSets { get; private set; } = 
            new List<SelectionFilterElement>();
        public List<(SelectionFilterElement element, string newName)> RenamedSelectionSets { get; private set; } = 
            new List<(SelectionFilterElement, string)>();
        public List<SelectionFilterElement> DeletedSelectionSets { get; private set; } = 
            new List<SelectionFilterElement>();
        public int CopyCount { get; private set; } = 1;

        public SelectionSetEditForm(List<SelectionSetInfo> selectionSets)
        {
            _selectionSets = selectionSets;
            
            Text            = "Edit Selection Sets";
            FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            StartPosition   = WinForms.FormStartPosition.CenterScreen;
            MaximizeBox = MinimizeBox = false;

            // Calculate optimal window height
            var screenHeight = WinForms.Screen.PrimaryScreen.WorkingArea.Height;
            var maxWindowHeight = screenHeight - 100; // Leave 100px padding
            var desiredDataGridHeight = Math.Max(300, Math.Min(600, selectionSets.Count * 18 + 50));
            var windowHeight = Math.Min(maxWindowHeight, desiredDataGridHeight + 60);
            var dataGridHeight = windowHeight - 60;

            ClientSize = new System.Drawing.Size(480, windowHeight);

            // Create DataGridView
            _dataGrid = new WinForms.DataGridView
            {
                Location = new System.Drawing.Point(12, 12),
                Size = new System.Drawing.Size(456, dataGridHeight),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                BackgroundColor = SystemColors.Window,
                BorderStyle = WinForms.BorderStyle.Fixed3D,
                RowTemplate = { Height = 18 },
                EditMode = WinForms.DataGridViewEditMode.EditProgrammatically
            };

            // Handle keyboard events
            _dataGrid.KeyDown += DataGrid_KeyDown;
            _dataGrid.CellBeginEdit += DataGrid_CellBeginEdit;
            _dataGrid.CellEndEdit += DataGrid_CellEndEdit;
            _dataGrid.CellValidating += DataGrid_CellValidating;
            _dataGrid.RowPrePaint += DataGrid_RowPrePaint;

            // Add columns
            var checkColumn = new WinForms.DataGridViewCheckBoxColumn
            {
                HeaderText = "Duplicate",
                Name = "Selected",
                DataPropertyName = "Selected",
                Width = 70,
                ReadOnly = false
            };
            
            var nameColumn = new WinForms.DataGridViewTextBoxColumn
            {
                HeaderText = "Selection Set Name",
                Name = "Name",
                DataPropertyName = "Name",
                Width = 290,
                ReadOnly = false
            };
            
            var countColumn = new WinForms.DataGridViewTextBoxColumn
            {
                HeaderText = "Elements",
                Name = "ElementCount",
                DataPropertyName = "ElementCount",
                Width = 80,
                ReadOnly = true
            };

            _dataGrid.Columns.AddRange(new WinForms.DataGridViewColumn[] 
                { checkColumn, nameColumn, countColumn });

            // Bind data
            _dataGrid.DataSource = _selectionSets;

            // Number of copies controls
            var lblCopies = new WinForms.Label
            {
                Text = "Number of copies:",
                Location = new System.Drawing.Point(12, dataGridHeight + 25),
                Size = new System.Drawing.Size(110, 20)
            };
            
            _numCopies = new WinForms.NumericUpDown
            {
                Location = new System.Drawing.Point(130, dataGridHeight + 22),
                Size = new System.Drawing.Size(60, 20),
                Minimum = 1,
                Maximum = 50,
                Value = 1
            };

            // OK and Cancel buttons
            var ok = new WinForms.Button 
            { 
                Text = "OK", 
                Location = new System.Drawing.Point(270, dataGridHeight + 20), 
                Size = new System.Drawing.Size(80, 25), 
                DialogResult = WinForms.DialogResult.OK 
            };
            
            var cancel = new WinForms.Button 
            { 
                Text = "Cancel", 
                Location = new System.Drawing.Point(360, dataGridHeight + 20), 
                Size = new Size(80, 25), 
                DialogResult = WinForms.DialogResult.Cancel 
            };

            AcceptButton = ok;
            CancelButton = cancel;

            Controls.AddRange(new WinForms.Control[]
            {
                _dataGrid, lblCopies, _numCopies, ok, cancel
            });
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
            
            if (e.KeyCode == WinForms.Keys.F2 && !_isEditingCell)
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
            else if (e.KeyCode == WinForms.Keys.Tab && !e.Shift && !_isEditingCell)
            {
                // Move focus to Number of Copies
                e.Handled = true;
                e.SuppressKeyPress = true;
                _numCopies.Focus();
            }
            else if (e.KeyCode == WinForms.Keys.Enter && !_isEditingCell)
            {
                // Act as OK button (only if not editing)
                e.Handled = true;
                e.SuppressKeyPress = true;
                DialogResult = WinForms.DialogResult.OK;
                Close();
            }
            else if (e.KeyCode == WinForms.Keys.Space && !_isEditingCell)
            {
                // Toggle checkboxes for selected rows (only if not editing)
                e.Handled = true;
                e.SuppressKeyPress = true;
                ToggleSelectedRows();
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
                _dataGrid.CurrentCell = selectedRows[0].Cells[1]; // Name column
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
                            
                            // Check if this name already exists in the list (excluding deleted items)
                            var existingItem = _selectionSets.FirstOrDefault(s => 
                                s != info && !s.MarkedForDeletion && 
                                string.Equals(s.Name, newName, StringComparison.OrdinalIgnoreCase));
                            
                            if (existingItem != null)
                            {
                                // Mark the existing item for deletion (overwrite)
                                existingItem.MarkedForDeletion = true;
                            }
                            
                            info.Name = newName;
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

            foreach (WinForms.DataGridViewRow row in selectedRows)
            {
                if (row.DataBoundItem is SelectionSetInfo info)
                {
                    info.MarkedForDeletion = !info.MarkedForDeletion;
                }
            }
            
            _dataGrid.Refresh();
        }

        private void DataGrid_CellValidating(object sender, WinForms.DataGridViewCellValidatingEventArgs e)
        {
            // Validate name column
            if (e.ColumnIndex == 1) // Name column
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
                    // Check for duplicate names (excluding current row and deleted items)
                    SelectionSetInfo duplicateItem = null;
                    for (int i = 0; i < _selectionSets.Count; i++)
                    {
                        if (i != e.RowIndex && 
                            !_selectionSets[i].MarkedForDeletion &&
                            string.Equals(_selectionSets[i].Name, newName, StringComparison.OrdinalIgnoreCase))
                        {
                            duplicateItem = _selectionSets[i];
                            break;
                        }
                    }
                    
                    if (duplicateItem != null)
                    {
                        // Mark the duplicate for deletion (overwrite behavior)
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
            
            // Update the underlying data when editing ends
            if (e.ColumnIndex == 1) // Name column
            {
                _dataGrid.Refresh();
            }
        }

        private void ToggleSelectedRows()
        {
            // Get all selected rows
            var selectedRows = _dataGrid.SelectedRows;
            if (selectedRows.Count == 0) return;

            // Determine new state based on first non-deleted selected row
            bool newState = true;
            bool foundNonDeleted = false;
            
            foreach (WinForms.DataGridViewRow row in selectedRows)
            {
                if (row.DataBoundItem is SelectionSetInfo info && !info.MarkedForDeletion)
                {
                    newState = !info.Selected;
                    foundNonDeleted = true;
                    break;
                }
            }
            
            if (!foundNonDeleted) return; // All selected items are marked for deletion

            // Apply new state to all selected rows (excluding deleted ones)
            foreach (WinForms.DataGridViewRow row in selectedRows)
            {
                if (row.DataBoundItem is SelectionSetInfo info && !info.MarkedForDeletion)
                {
                    info.Selected = newState;
                }
            }
            
            _dataGrid.Refresh();
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

            // Collect selection sets to duplicate (excluding deleted ones)
            SelectedSelectionSets = _selectionSets
                .Where(s => s.Selected && !s.MarkedForDeletion)
                .Select(s => s.FilterElement)
                .ToList();

            // Collect selection sets to rename (excluding deleted ones)
            RenamedSelectionSets = _selectionSets
                .Where(s => s.NameChanged && !s.MarkedForDeletion)
                .Select(s => (s.FilterElement, s.Name))
                .ToList();

            // Collect selection sets to delete
            DeletedSelectionSets = _selectionSets
                .Where(s => s.MarkedForDeletion)
                .Select(s => s.FilterElement)
                .ToList();

            if (SelectedSelectionSets.Count == 0 && RenamedSelectionSets.Count == 0 && DeletedSelectionSets.Count == 0)
            {
                WinForms.MessageBox.Show("No changes were made.",
                    "No Changes", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                e.Cancel = true;
                return;
            }

            CopyCount = (int)_numCopies.Value;
        }
    }

    // ─────────────────────────────────────────────────────────────
    //  Core editing helper
    // ─────────────────────────────────────────────────────────────
    internal static class SelectionSetEditor
    {
        private static readonly Regex _numberSuffix =
            new Regex(@"\s+(\d+)$", RegexOptions.Compiled);

        public static List<SelectionFilterElement> DuplicateSelectionSets(
            Document doc,
            IEnumerable<SelectionFilterElement> selectionSets,
            int copies,
            HashSet<string> usedNames)
        {
            var createdSelectionSets = new List<SelectionFilterElement>();

            foreach (SelectionFilterElement originalSet in selectionSets)
            {
                // Get the element IDs from the original selection set
                ICollection<ElementId> elementIds = originalSet.GetElementIds();

                for (int n = 0; n < copies; ++n)
                {
                    // Generate new name
                    string newName = GetNextAvailableName(originalSet.Name, usedNames);
                    usedNames.Add(newName);

                    // Create new selection set
                    SelectionFilterElement newSet = SelectionFilterElement.Create(doc, newName);
                    
                    // Add the same elements to the new selection set
                    newSet.SetElementIds(elementIds);

                    createdSelectionSets.Add(newSet);
                }
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

        private static string GetNextAvailableName(string baseName, HashSet<string> usedNames)
        {
            // Check if name already ends with a number
            Match match = _numberSuffix.Match(baseName);
            
            string nameWithoutNumber;
            int startNumber;

            if (match.Success)
            {
                // Extract base name and current number
                nameWithoutNumber = baseName.Substring(0, match.Index);
                startNumber = int.Parse(match.Groups[1].Value) + 1;
            }
            else
            {
                // No number suffix, start with 2
                nameWithoutNumber = baseName;
                startNumber = 2;
            }

            // Find next available number
            int number = startNumber;
            string candidate = $"{nameWithoutNumber} {number}";
            
            while (usedNames.Contains(candidate, StringComparer.OrdinalIgnoreCase))
            {
                number++;
                candidate = $"{nameWithoutNumber} {number}";
            }

            return candidate;
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
                    Selected = false,
                    Name = set.Name,
                    OriginalName = set.Name,
                    ElementCount = set.GetElementIds().Count,
                    MarkedForDeletion = false
                });
            }

            // Show dialog
            List<SelectionFilterElement> selectedSets;
            List<(SelectionFilterElement, string)> renamedSets;
            List<SelectionFilterElement> deletedSets;
            int copies;
            
            using (var dlg = new SelectionSetEditForm(selectionSetInfos))
            {
                if (dlg.ShowDialog() != WinForms.DialogResult.OK) 
                    return Result.Cancelled;
                    
                selectedSets = dlg.SelectedSelectionSets;
                renamedSets = dlg.RenamedSelectionSets;
                deletedSets = dlg.DeletedSelectionSets;
                copies = dlg.CopyCount;
            }

            // Perform operations
            using (var t = new Transaction(doc, "Edit Selection Sets"))
            {
                t.Start();
                
                try
                {
                    // Delete first - this includes items marked for deletion and those being overwritten
                    if (deletedSets.Count > 0)
                    {
                        SelectionSetEditor.DeleteSelectionSets(doc, deletedSets);
                    }
                    
                    // Get all current names for duplicate checking (after deletion)
                    var usedNames = new FilteredElementCollector(doc)
                        .OfClass(typeof(SelectionFilterElement))
                        .Cast<SelectionFilterElement>()
                        .Select(s => s.Name)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                    
                    // Apply renames (only for items that weren't deleted)
                    if (renamedSets.Count > 0)
                    {
                        SelectionSetEditor.RenameSelectionSets(doc, renamedSets);
                        
                        // Update used names with new names
                        foreach (var (element, newName) in renamedSets)
                        {
                            usedNames.Add(newName);
                        }
                    }
                    
                    // Then duplicate (only non-deleted items)
                    List<SelectionFilterElement> createdSets = null;
                    if (selectedSets.Count > 0)
                    {
                        createdSets = SelectionSetEditor.DuplicateSelectionSets(
                            doc, selectedSets, copies, usedNames);
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
