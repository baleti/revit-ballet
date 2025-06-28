// DuplicateSelectionSets.cs  –  Revit 2024 / C# 7.3 compatible
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
        public string Name => FilterElement?.Name ?? string.Empty;
        public int ElementCount { get; set; }
    }
    
    // ─────────────────────────────────────────────────────────────
    //  WinForm
    // ─────────────────────────────────────────────────────────────
    internal class SelectionSetDuplicationForm : WinForms.Form
    {
        private readonly WinForms.DataGridView _dataGrid;
        private readonly WinForms.NumericUpDown _numCopies;
        private readonly List<SelectionSetInfo> _selectionSets;

        public List<SelectionFilterElement> SelectedSelectionSets { get; private set; } = 
            new List<SelectionFilterElement>();
        public int CopyCount { get; private set; } = 1;

        public SelectionSetDuplicationForm(List<SelectionSetInfo> selectionSets)
        {
            _selectionSets = selectionSets;
            
            Text            = "Duplicate Selection Sets";
            FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            StartPosition   = WinForms.FormStartPosition.CenterScreen;
            MaximizeBox = MinimizeBox = false;

            // Calculate optimal window height
            var screenHeight = WinForms.Screen.PrimaryScreen.WorkingArea.Height;
            var maxWindowHeight = screenHeight - 100; // Leave 100px padding
            var desiredDataGridHeight = Math.Max(300, Math.Min(600, selectionSets.Count * 18 + 50));
            var windowHeight = Math.Min(maxWindowHeight, desiredDataGridHeight + 60);
            var dataGridHeight = windowHeight - 60;

            ClientSize = new Size(480, windowHeight);

            // Create DataGridView
            _dataGrid = new WinForms.DataGridView
            {
                Location = new System.Drawing.Point(12, 12),
                Size = new Size(456, dataGridHeight),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                BackgroundColor = SystemColors.Window,
                BorderStyle = WinForms.BorderStyle.Fixed3D,
                RowTemplate = { Height = 18 }
            };

            // Handle keyboard events
            _dataGrid.KeyDown += DataGrid_KeyDown;

            // Add columns
            var checkColumn = new WinForms.DataGridViewCheckBoxColumn
            {
                HeaderText = "Select",
                Name = "Selected",
                DataPropertyName = "Selected",
                Width = 60,
                ReadOnly = false
            };
            
            var nameColumn = new WinForms.DataGridViewTextBoxColumn
            {
                HeaderText = "Selection Set Name",
                Name = "Name",
                DataPropertyName = "Name",
                Width = 300,
                ReadOnly = true
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
                Size = new Size(110, 20)
            };
            
            _numCopies = new WinForms.NumericUpDown
            {
                Location = new System.Drawing.Point(130, dataGridHeight + 22),
                Size = new Size(60, 20),
                Minimum = 1,
                Maximum = 50,
                Value = 1
            };

            // OK and Cancel buttons
            var ok = new WinForms.Button 
            { 
                Text = "OK", 
                Location = new System.Drawing.Point(270, dataGridHeight + 20), 
                Size = new Size(80, 25), 
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

        private void DataGrid_KeyDown(object sender, WinForms.KeyEventArgs e)
        {
            if (e.KeyCode == WinForms.Keys.Tab && !e.Shift)
            {
                // Move focus to Number of Copies
                e.Handled = true;
                e.SuppressKeyPress = true;
                _numCopies.Focus();
            }
            else if (e.KeyCode == WinForms.Keys.Enter)
            {
                // Act as OK button
                e.Handled = true;
                e.SuppressKeyPress = true;
                DialogResult = WinForms.DialogResult.OK;
                Close();
            }
            else if (e.KeyCode == WinForms.Keys.Space)
            {
                // Toggle checkboxes for selected rows
                e.Handled = true;
                e.SuppressKeyPress = true;
                ToggleSelectedRows();
            }
            else if (e.Control && e.KeyCode == WinForms.Keys.A)
            {
                // Select all rows
                e.Handled = true;
                e.SuppressKeyPress = true;
                _dataGrid.SelectAll();
            }
        }

        private void ToggleSelectedRows()
        {
            // Get all selected rows
            var selectedRows = _dataGrid.SelectedRows;
            if (selectedRows.Count == 0) return;

            // Determine new state based on first selected row
            bool newState = true;
            foreach (WinForms.DataGridViewRow row in selectedRows)
            {
                if (row.DataBoundItem is SelectionSetInfo info)
                {
                    newState = !info.Selected;
                    break;
                }
            }

            // Apply new state to all selected rows
            foreach (WinForms.DataGridViewRow row in selectedRows)
            {
                if (row.DataBoundItem is SelectionSetInfo info)
                {
                    info.Selected = newState;
                }
            }
            
            _dataGrid.Refresh();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);
            if (DialogResult != WinForms.DialogResult.OK) return;

            SelectedSelectionSets = _selectionSets
                .Where(s => s.Selected)
                .Select(s => s.FilterElement)
                .ToList();

            if (SelectedSelectionSets.Count == 0)
            {
                WinForms.MessageBox.Show("Please select at least one selection set to duplicate.",
                    "No Selection", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            CopyCount = (int)_numCopies.Value;
        }
    }

    // ─────────────────────────────────────────────────────────────
    //  Core duplication helper
    // ─────────────────────────────────────────────────────────────
    internal static class SelectionSetDuplicator
    {
        private static readonly Regex _numberSuffix =
            new Regex(@"\s+(\d+)$", RegexOptions.Compiled);

        public static List<SelectionFilterElement> DuplicateSelectionSets(
            Document doc,
            IEnumerable<SelectionFilterElement> selectionSets,
            int copies)
        {
            var usedNames =
                new FilteredElementCollector(doc)
                    .OfClass(typeof(SelectionFilterElement))
                    .Cast<SelectionFilterElement>()
                    .Select(s => s.Name)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

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
    public class DuplicateSelectionSets : IExternalCommand
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
                TaskDialog.Show("Duplicate Selection Sets", 
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
                    ElementCount = set.GetElementIds().Count
                });
            }

            // Show dialog
            List<SelectionFilterElement> selectedSets;
            int copies;
            
            using (var dlg = new SelectionSetDuplicationForm(selectionSetInfos))
            {
                if (dlg.ShowDialog() != WinForms.DialogResult.OK) 
                    return Result.Cancelled;
                    
                selectedSets = dlg.SelectedSelectionSets;
                copies = dlg.CopyCount;
            }

            // Perform duplication
            using (var t = new Transaction(doc, "Duplicate Selection Sets"))
            {
                t.Start();
                List<SelectionFilterElement> createdSets = null;
                
                try
                {
                    createdSets = SelectionSetDuplicator.DuplicateSelectionSets(
                        doc, selectedSets, copies);
                    t.Commit();
                    
                    // Report results
                    string msg = $"Successfully created {createdSets.Count} selection set(s).";
                    if (createdSets.Count <= 5)
                    {
                        msg += "\n\nCreated sets:";
                        foreach (var set in createdSets)
                        {
                            msg += $"\n• {set.Name}";
                        }
                    }
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
