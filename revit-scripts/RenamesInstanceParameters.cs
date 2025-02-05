using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class RenamesInstanceParameters : IExternalCommand
{
    private class ParameterInfo
    {
        public string Name { get; set; }
        public bool IsShared { get; set; }
    }

    private class ParameterValueInfo
    {
        public Element Element { get; set; }
        public Parameter Parameter { get; set; }
        public string ParameterName { get; set; }
        public string CurrentValue { get; set; }
        public bool IsShared { get; set; }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        try
        {
            var selectedIds = uiDoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("Error", "Please select elements first.");
                return Result.Cancelled;
            }

            var selectedElements = selectedIds.Cast<ElementId>()
                .Select(id => doc.GetElement(id))
                .Where(e => e != null)
                .ToList();

            var paramNames = selectedElements
                .SelectMany(elem => elem.Parameters.Cast<Parameter>())
                .Where(p => !p.IsReadOnly && p.StorageType == StorageType.String)
                .GroupBy(p => p.Definition.Name)
                .Select(g => new ParameterInfo 
                { 
                    Name = g.Key,
                    IsShared = g.First().IsShared
                })
                .ToHashSet();

            if (paramNames.Count == 0)
            {
                TaskDialog.Show("Error", "No writable string parameters found.");
                return Result.Cancelled;
            }

            var paramList = paramNames.ToList();
            var selectedParams = CustomGUIs.DataGrid(
                paramList,
                new List<string> { "Name" },
                null,
                "Select Parameters to Rename"
            );

            if (selectedParams == null || selectedParams.Count == 0)
                return Result.Succeeded;

            var paramValues = new List<ParameterValueInfo>();
            foreach (var elem in selectedElements)
            {
                foreach (var paramInfo in selectedParams)
                {
                    var param = elem.LookupParameter(paramInfo.Name);
                    if (param?.StorageType == StorageType.String && !param.IsReadOnly)
                    {
                        paramValues.Add(new ParameterValueInfo
                        {
                            Element = elem,
                            Parameter = param,
                            ParameterName = paramInfo.Name,
                            CurrentValue = param.AsString() ?? "",
                            IsShared = param.IsShared
                        });
                    }
                }
            }

            if (paramValues.Count == 0)
            {
                TaskDialog.Show("Error", "No valid parameters found in selection.");
                return Result.Cancelled;
            }

            using (var form = new RenameParamForm(paramValues))
            {
                if (form.ShowDialog() != DialogResult.OK)
                    return Result.Succeeded;

                using (var tx = new Transaction(doc, "Rename Parameters"))
                {
                    tx.Start();
                    foreach (var pv in paramValues)
                    {
                        try
                        {
                            string currentValue = pv.CurrentValue;
                            
                            if (!string.IsNullOrEmpty(form.FindText) && 
                                currentValue.Contains(form.FindText))
                            {
                                currentValue = currentValue.Replace(form.FindText, form.ReplaceText);
                            }

                            if (!string.IsNullOrEmpty(form.PatternText))
                            {
                                currentValue = form.PatternText.Replace("{}", currentValue);
                            }

                            if (currentValue != pv.CurrentValue)
                            {
                                pv.Parameter.Set(currentValue);
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Error", $"Failed to update {pv.Element.Name}: {ex.Message}");
                        }
                    }
                    tx.Commit();
                }
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Critical Error", ex.ToString());
            return Result.Failed;
        }
    }

    private class RenameParamForm : System.Windows.Forms.Form
    {
        private readonly List<ParameterValueInfo> _paramValues;
        private System.Windows.Forms.TextBox _txtFind;
        private System.Windows.Forms.TextBox _txtReplace;
        private System.Windows.Forms.TextBox _txtPattern;
        private System.Windows.Forms.RichTextBox _rtbBefore;
        private System.Windows.Forms.RichTextBox _rtbAfter;

        public string FindText => _txtFind.Text;
        public string ReplaceText => _txtReplace.Text;
        public string PatternText => _txtPattern.Text;

        public RenameParamForm(List<ParameterValueInfo> paramValues)
        {
            _paramValues = paramValues;
            InitializeComponent();
            InitializeData();
            
            this.KeyDown += (s, e) => 
            {
                if (e.KeyCode == Keys.Escape)
                {
                    this.DialogResult = DialogResult.Cancel;
                    this.Close();
                }
            };
        }

        private void InitializeComponent()
        {
            this.Text = "Modify Parameters";
            this.Size = new System.Drawing.Size(600, 500);
            this.MinimumSize = new System.Drawing.Size(500, 400);
            this.Font = new System.Drawing.Font("Segoe UI", 9);
            this.KeyPreview = true;

            var mainLayout = new System.Windows.Forms.TableLayoutPanel
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6,
                Padding = new System.Windows.Forms.Padding(10),
                ColumnStyles = 
                {
                    new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80),
                    new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100)
                },
                RowStyles = 
                {
                    new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30),
                    new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30),
                    new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30),
                    new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20),
                    new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45),
                    new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45)
                }
            };

            mainLayout.Controls.Add(new System.Windows.Forms.Label { 
                Text = "Find:", 
                TextAlign = System.Drawing.ContentAlignment.MiddleRight 
            }, 0, 0);
            _txtFind = new System.Windows.Forms.TextBox { Dock = System.Windows.Forms.DockStyle.Fill };
            mainLayout.Controls.Add(_txtFind, 1, 0);

            mainLayout.Controls.Add(new System.Windows.Forms.Label { 
                Text = "Replace:", 
                TextAlign = System.Drawing.ContentAlignment.MiddleRight 
            }, 0, 1);
            _txtReplace = new System.Windows.Forms.TextBox { Dock = System.Windows.Forms.DockStyle.Fill };
            mainLayout.Controls.Add(_txtReplace, 1, 1);

            mainLayout.Controls.Add(new System.Windows.Forms.Label { 
                Text = "Pattern:", 
                TextAlign = System.Drawing.ContentAlignment.MiddleRight 
            }, 0, 2);
            _txtPattern = new System.Windows.Forms.TextBox { 
                Dock = System.Windows.Forms.DockStyle.Fill, 
                Text = "{}" 
            };
            mainLayout.Controls.Add(_txtPattern, 1, 2);

            var lblHint = new System.Windows.Forms.Label 
            { 
                Text = "Use {} to represent current value",
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Italic),
                ForeColor = System.Drawing.Color.Gray,
                Dock = System.Windows.Forms.DockStyle.Top
            };
            mainLayout.Controls.Add(lblHint, 1, 3);

            _rtbBefore = new System.Windows.Forms.RichTextBox { 
                ReadOnly = true, 
                Dock = System.Windows.Forms.DockStyle.Fill 
            };
            var groupBefore = new System.Windows.Forms.GroupBox { 
                Text = "Current Values", 
                Dock = System.Windows.Forms.DockStyle.Fill 
            };
            groupBefore.Controls.Add(_rtbBefore);
            mainLayout.Controls.Add(groupBefore, 0, 4);
            mainLayout.SetColumnSpan(groupBefore, 2);

            _rtbAfter = new System.Windows.Forms.RichTextBox { 
                ReadOnly = true, 
                Dock = System.Windows.Forms.DockStyle.Fill 
            };
            var groupAfter = new System.Windows.Forms.GroupBox { 
                Text = "Preview", 
                Dock = System.Windows.Forms.DockStyle.Fill 
            };
            groupAfter.Controls.Add(_rtbAfter);
            mainLayout.Controls.Add(groupAfter, 0, 5);
            mainLayout.SetColumnSpan(groupAfter, 2);

            var buttonPanel = new System.Windows.Forms.Panel { 
                Dock = System.Windows.Forms.DockStyle.Bottom, 
                Height = 40 
            };

            var btnCancel = new System.Windows.Forms.Button 
            { 
                Text = "Cancel",
                Size = new System.Drawing.Size(80, 30),
                DialogResult = System.Windows.Forms.DialogResult.Cancel
            };
            buttonPanel.Controls.Add(btnCancel);
            btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            
            var btnOk = new System.Windows.Forms.Button 
            { 
                Text = "OK",
                Size = new System.Drawing.Size(80, 30),
                DialogResult = System.Windows.Forms.DialogResult.OK
            };
            buttonPanel.Controls.Add(btnOk);
            
            btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            
            btnOk.Location = new System.Drawing.Point(
                buttonPanel.Width - btnOk.Width - 10, 
                buttonPanel.Height - btnOk.Height - 5
            );
            btnCancel.Location = new System.Drawing.Point(
                btnOk.Left - btnCancel.Width - 10,
                btnOk.Top
            );

            _txtFind.TextChanged += UpdatePreview;
            _txtReplace.TextChanged += UpdatePreview;
            _txtPattern.TextChanged += UpdatePreview;
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
            
            this.Controls.Add(mainLayout);
            this.Controls.Add(buttonPanel);
        }

        private void InitializeData()
        {
            foreach (var pv in _paramValues)
            {
                _rtbBefore.AppendText($"{pv.CurrentValue}{System.Environment.NewLine}");
            }
            UpdatePreview(this, EventArgs.Empty);
        }

        private void UpdatePreview(object sender, EventArgs e)
        {
            _rtbAfter.Clear();
            foreach (var pv in _paramValues)
            {
                string current = ApplyTransformations(pv.CurrentValue);
                _rtbAfter.AppendText($"{current}{System.Environment.NewLine}");
            }
        }

        private string ApplyTransformations(string input)
        {
            if (!string.IsNullOrEmpty(FindText) && input.Contains(FindText))
            {
                input = input.Replace(FindText, ReplaceText);
            }

            if (!string.IsNullOrEmpty(PatternText))
            {
                input = PatternText.Replace("{}", input);
            }

            return input;
        }
    }
}
