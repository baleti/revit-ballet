using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
// Alias WinForms to avoid ambiguity
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace MyRevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class RenameSelectedTexts : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Only TextNotes
            var textNotes = uiDoc.GetSelectionIds()
                              .Select(id => doc.GetElement(id) as TextNote)
                              .Where(tn => tn != null)
                              .ToList();

            if (textNotes.Count == 0)
            {
                WinForms.MessageBox.Show("Please select one or more text notes first.", "Error");
                return Result.Cancelled;
            }

            var originals = textNotes.Select(tn => tn.Text).ToList();

            using (var form = new RenameTextForm(originals))
            {
                if (form.ShowDialog() != WinForms.DialogResult.OK)
                    return Result.Succeeded;

                using (var tx = new Transaction(doc, "Rename Selected Texts"))
                {
                    tx.Start();
                    for (int i = 0; i < textNotes.Count; i++)
                    {
                        string newText = form.GetTransformedValue(originals[i]);
                        if (newText != originals[i])
                            textNotes[i].Text = newText;
                    }
                    tx.Commit();
                }
            }

            return Result.Succeeded;
        }
    }

    public class RenameTextForm : WinForms.Form
    {
        private readonly List<string> _originalValues;
        private WinForms.TextBox _txtFind, _txtReplace, _txtPattern, _txtMath;
        private WinForms.TextBox _txtBefore, _txtAfter;

        public string FindText    => _txtFind.Text;
        public string ReplaceText => _txtReplace.Text;
        public string PatternText => _txtPattern.Text;
        public string MathOpText  => _txtMath.Text;

        public RenameTextForm(List<string> originalValues)
        {
            _originalValues = originalValues;
            InitializeComponent();
            LoadOriginals();
            UpdatePreview(null, EventArgs.Empty);
            this.KeyDown += (s, e) =>
            {
                if (e.KeyCode == WinForms.Keys.Escape)
                {
                    this.DialogResult = WinForms.DialogResult.Cancel;
                    this.Close();
                }
            };
        }

        private void InitializeComponent()
        {
            this.Text        = "Rename Text Notes";
            this.Size        = new Drawing.Size(600, 600);
            this.MinimumSize = new Drawing.Size(500, 400);
            this.Font        = new Drawing.Font("Segoe UI", 9);
            this.KeyPreview  = true;

            var main = new WinForms.TableLayoutPanel
            {
                Dock       = WinForms.DockStyle.Fill,
                ColumnCount = 2,
                RowCount    = 8,
                Padding    = new WinForms.Padding(10)
            };
            main.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 80));
            main.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100));
            main.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30)); // Find
            main.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30)); // Replace
            main.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30)); // Pattern
            main.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 20)); // Pattern hint
            main.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30)); // Math
            main.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 20)); // Math hint
            main.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));  // Before
            main.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));  // After

            // Row 0: Find
            main.Controls.Add(new WinForms.Label
            {
                Text      = "Find:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock      = WinForms.DockStyle.Fill
            }, 0, 0);
            _txtFind = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            main.Controls.Add(_txtFind, 1, 0);

            // Row 1: Replace
            main.Controls.Add(new WinForms.Label
            {
                Text      = "Replace:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock      = WinForms.DockStyle.Fill
            }, 0, 1);
            _txtReplace = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            main.Controls.Add(_txtReplace, 1, 1);

            // Row 2: Pattern
            main.Controls.Add(new WinForms.Label
            {
                Text      = "Pattern:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock      = WinForms.DockStyle.Fill
            }, 0, 2);
            _txtPattern = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill, Text = "{}" };
            main.Controls.Add(_txtPattern, 1, 2);
            main.Controls.Add(new WinForms.Label
            {
                Text      = "Use {} to represent current value.",
                Font      = new Drawing.Font(this.Font, Drawing.FontStyle.Italic),
                ForeColor = Drawing.Color.Gray,
                Dock      = WinForms.DockStyle.Fill
            }, 1, 3);

            // Row 4: Math
            main.Controls.Add(new WinForms.Label
            {
                Text      = "Math:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock      = WinForms.DockStyle.Fill
            }, 0, 4);
            _txtMath = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            main.Controls.Add(_txtMath, 1, 4);
            main.Controls.Add(new WinForms.Label
            {
                Text      = "Use x (e.g. '2x','x/2','x+2','x-2','-x').",
                Font      = new Drawing.Font(this.Font, Drawing.FontStyle.Italic),
                ForeColor = Drawing.Color.Gray,
                Dock      = WinForms.DockStyle.Fill
            }, 1, 5);

            // Row 6: Current Text (multi-line TextBox)
            _txtBefore = new WinForms.TextBox
            {
                Multiline     = true,
                AcceptsReturn = true,
                ReadOnly      = true,
                ScrollBars    = WinForms.ScrollBars.Vertical,
                Dock          = WinForms.DockStyle.Fill
            };
            var grpBefore = new WinForms.GroupBox { Text = "Current Text", Dock = WinForms.DockStyle.Fill };
            grpBefore.Controls.Add(_txtBefore);
            main.Controls.Add(grpBefore, 0, 6);
            main.SetColumnSpan(grpBefore, 2);

            // Row 7: Preview (multi-line TextBox)
            _txtAfter = new WinForms.TextBox
            {
                Multiline     = true,
                AcceptsReturn = true,
                ReadOnly      = true,
                ScrollBars    = WinForms.ScrollBars.Vertical,
                Dock          = WinForms.DockStyle.Fill
            };
            var grpAfter = new WinForms.GroupBox { Text = "Preview", Dock = WinForms.DockStyle.Fill };
            grpAfter.Controls.Add(_txtAfter);
            main.Controls.Add(grpAfter, 0, 7);
            main.SetColumnSpan(grpAfter, 2);

            // Buttons
            var buttonPanel = new WinForms.Panel { Dock = WinForms.DockStyle.Bottom, Height = 40 };
            var btnCancel   = new WinForms.Button { Text = "Cancel", Size = new Drawing.Size(80, 30), DialogResult = WinForms.DialogResult.Cancel };
            var btnOk       = new WinForms.Button { Text = "OK",     Size = new Drawing.Size(80, 30), DialogResult = WinForms.DialogResult.OK };
            buttonPanel.Controls.Add(btnCancel);
            buttonPanel.Controls.Add(btnOk);
            btnOk.Anchor     = btnCancel.Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right;
            btnOk.Location   = new Drawing.Point(buttonPanel.Width - btnOk.Width - 10, buttonPanel.Height - btnOk.Height - 5);
            btnCancel.Location = new Drawing.Point(btnOk.Left - btnCancel.Width - 10, btnOk.Top);

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
            this.Controls.Add(main);
            this.Controls.Add(buttonPanel);

            // Preview updates
            _txtFind.TextChanged    += UpdatePreview;
            _txtReplace.TextChanged += UpdatePreview;
            _txtPattern.TextChanged += UpdatePreview;
            _txtMath.TextChanged    += UpdatePreview;
        }

        private void LoadOriginals()
        {
            foreach (var s in _originalValues)
                _txtBefore.AppendText(s + Environment.NewLine);
        }

        private void UpdatePreview(object sender, EventArgs e)
        {
            _txtAfter.Clear();
            foreach (var orig in _originalValues)
                _txtAfter.AppendText(GetTransformedValue(orig) + Environment.NewLine);
        }

        public string GetTransformedValue(string original)
        {
            string val = original;

            // Find / Replace
            if (!string.IsNullOrEmpty(FindText))
                val = val.Replace(FindText, ReplaceText);
            else if (!string.IsNullOrEmpty(ReplaceText))
                val = ReplaceText;

            // Pattern
            if (!string.IsNullOrEmpty(PatternText))
                val = PatternText.Replace("{}", val);

            // Math
            if (!string.IsNullOrWhiteSpace(MathOpText) &&
                double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out double x))
            {
                val = ApplyMath(x, MathOpText).ToString(CultureInfo.InvariantCulture);
            }

            return val;
        }

        private double ApplyMath(double x, string op)
        {
            op = op.Replace(" ", "");
            if (op.Equals("x",  StringComparison.OrdinalIgnoreCase)) return x;
            if (op.Equals("-x", StringComparison.OrdinalIgnoreCase)) return -x;

            // "2x" style
            if (op.EndsWith("x", StringComparison.OrdinalIgnoreCase) && op.Length > 1)
            {
                var multStr = op.Substring(0, op.Length - 1);
                if (double.TryParse(multStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double m))
                    return m * x;
            }

            // "x+2", "x-2", "x*2", "x/2"
            if (op.StartsWith("x", StringComparison.OrdinalIgnoreCase) && op.Length > 1)
            {
                var tail = op.Substring(1);
                if (tail.StartsWith("+") && double.TryParse(tail.Substring(1), out double a)) return x + a;
                if (tail.StartsWith("-") && double.TryParse(tail.Substring(1), out double s)) return x - s;
                if (tail.StartsWith("*") && double.TryParse(tail.Substring(1), out double m2)) return x * m2;
                if (tail.StartsWith("/") && double.TryParse(tail.Substring(1), out double d) && d != 0) return x / d;
            }

            return x;
        }
    }
}
