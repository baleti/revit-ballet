// *************************************************************************************************
//  RenameInstanceParametersOfSelectedElements.cs
//  Revit add-in (C# 7.3 / .NET 4.8 compliant)
// *************************************************************************************************
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

// Alias WinForms and Drawing to avoid type ambiguities.
using WinForms = System.Windows.Forms;
using Drawing  = System.Drawing;

namespace MyRevitAddin
{
    // =============================================================================================
    // 1. Shared helper – math, reflection and RenameParameters worker
    // =============================================================================================
    public static class ParameterRenamerHelper
    {
        #region DTOs ------------------------------------------------------------------------------

        public class ParameterInfo
        {
            public string Name;
            public bool   IsShared;
            public string Group;
        }

        public class ParameterValueInfo
        {
            public Element   Element;        // owning element
            public Parameter Parameter;      // null if pseudo parameter
            public string    ParameterName;
            public string    CurrentValue;
            public bool      IsShared;
        }

        #endregion

        #region Math helpers ----------------------------------------------------------------------

        /// <summary>Interprets <paramref name="mathOp"/> ("2x", "x/2", "x+1", …) and applies it to <paramref name="x"/>.</summary>
        public static double ApplyMathOperation(double x, string mathOp)
        {
            if (string.IsNullOrWhiteSpace(mathOp))
                return x;

            mathOp = mathOp.Replace(" ", string.Empty);

            // identity / negate
            if (mathOp.Equals("x",  StringComparison.OrdinalIgnoreCase)) return x;
            if (mathOp.Equals("-x", StringComparison.OrdinalIgnoreCase)) return -x;

            // "2x"  (multiplier before x)
            if (!mathOp.StartsWith("x", StringComparison.OrdinalIgnoreCase) &&
                 mathOp.EndsWith("x",  StringComparison.OrdinalIgnoreCase))
            {
                string multStr = mathOp.Substring(0, mathOp.Length - 1);
                if (double.TryParse(multStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double mult))
                    return mult * x;
            }

            // operations after x
            if (mathOp.StartsWith("x", StringComparison.OrdinalIgnoreCase) && mathOp.Length >= 3)
            {
                char   op  = mathOp[1];
                string num = mathOp.Substring(2);
                if (double.TryParse(num, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                {
                    switch (op)
                    {
                        case '+': return x + n;
                        case '-': return x - n;
                        case '*': return x * n;
                        case '/': return Math.Abs(n) < double.Epsilon ? x : x / n;
                    }
                }
            }
            return x;   // unrecognised -> unchanged
        }

        /// <summary>Applies <see cref="ApplyMathOperation"/> to every numeric token in <paramref name="input"/>.</summary>
        public static string ApplyMathToNumbersInString(string input, string mathOp)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrWhiteSpace(mathOp))
                return input;

            return Regex.Replace(
                input,
                @"-?\d+(?:\.\d+)?",                              // signed integers/decimals
                m =>
                {
                    if (!double.TryParse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                        return m.Value;                           // shouldn't happen

                    double res = ApplyMathOperation(n, mathOp);

                    // keep integer look-and-feel where possible
                    return Math.Abs(res % 1) < 1e-10
                         ? Math.Round(res).ToString(CultureInfo.InvariantCulture)
                         : res.ToString(CultureInfo.InvariantCulture);
                });
        }

        #endregion

        #region Reflection helper – Yes/No parameters ---------------------------------------------

        public static bool IsYesNoParameter(Parameter param)
        {
            if (param?.Definition == null) return false;

            PropertyInfo pi = param.Definition.GetType().GetProperty(
                                "ParameterType",
                                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (pi != null)
            {
                object val = pi.GetValue(param.Definition, null);
                if (val != null)
                {
                    string s = val.ToString().Replace("/", string.Empty).Replace(" ", string.Empty);
                    return s.Equals("YesNo", StringComparison.OrdinalIgnoreCase);
                }
            }
            return false;
        }

        #endregion

        #region Parameter value getter helper -----------------------------------------------------

        /// <summary>Gets parameter value as string</summary>
        private static string GetParameterValueAsString(Parameter p)
        {
            if (p == null) return string.Empty;
            return (p.StorageType == StorageType.String)
                ? (p.AsString() ?? string.Empty)
                : (p.AsValueString() ?? string.Empty);
        }

        #endregion

        #region Pattern parsing helper ------------------------------------------------------------

        /// <summary>
        /// Parses pattern string and replaces parameter references.
        /// Supports: {} for current value, $"Parameter Name" or $ParameterName for other parameters
        /// </summary>
        private static string ParsePatternWithParameterReferences(string pattern, string currentValue, Element element)
        {
            if (string.IsNullOrEmpty(pattern))
                return currentValue;

            // First replace {} with current value
            string result = pattern.Replace("{}", currentValue);

            // Regex to match $"Parameter Name" or $ParameterNameWithoutSpaces
            // This matches: $"anything in quotes" OR $word (alphanumeric and underscore)
            var regex = new Regex(@"\$""([^""]+)""|(?<!\$)\$(\w+)");
            
            result = regex.Replace(result, match =>
            {
                // Get the parameter name from either the quoted group or the unquoted group
                string paramName = !string.IsNullOrEmpty(match.Groups[1].Value) 
                    ? match.Groups[1].Value 
                    : match.Groups[2].Value;

                // Look up the parameter on the element
                Parameter param = element.LookupParameter(paramName);
                if (param != null)
                {
                    return GetParameterValueAsString(param);
                }
                
                // If parameter not found, return the original text
                return match.Value;
            });

            return result;
        }

        #endregion

        #region Value transformation --------------------------------------------------------------

        private static string TransformValue(string original, RenameParamForm form, Element element = null)
        {
            string value = original;

            // 1. Find / Replace
            if (!string.IsNullOrEmpty(form.FindText))
                value = value.Replace(form.FindText, form.ReplaceText);
            else if (!string.IsNullOrEmpty(form.ReplaceText))
                value = form.ReplaceText;

            // 2. Pattern (now with parameter reference support)
            if (!string.IsNullOrEmpty(form.PatternText))
            {
                if (element != null)
                {
                    // Use the new pattern parser that supports parameter references
                    value = ParsePatternWithParameterReferences(form.PatternText, value, element);
                }
                else
                {
                    // Fallback for preview mode (no element available)
                    value = form.PatternText.Replace("{}", value);
                }
            }

            // 3. Math
            if (!string.IsNullOrEmpty(form.MathOperationText))
            {
                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                    value = ApplyMathOperation(n, form.MathOperationText).ToString(CultureInfo.InvariantCulture);
                else
                    value = ApplyMathToNumbersInString(value, form.MathOperationText);
            }
            return value;
        }

        #endregion

        #region Parameter setter ------------------------------------------------------------------

        private static void SetParameterValue(Parameter p, string newVal, bool isMetric)
        {
            switch (p.StorageType)
            {
                case StorageType.String:
                    p.Set(newVal);
                    break;

                case StorageType.Double:
                    if (double.TryParse(newVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                    {
                        // convert mm → ft if necessary
                        ForgeTypeId ut = p.GetUnitTypeId();
                        if (isMetric &&
                            (ut == UnitTypeId.Millimeters || ut == UnitTypeId.Meters))
                            d = d / 304.8;
                        p.Set(d);
                    }
                    break;

                case StorageType.Integer:
                    if (IsYesNoParameter(p))
                    {
                        string s = newVal.Trim().ToLowerInvariant();
                        if (s == "yes" || s == "true" || s == "1")  p.Set(1);
                        if (s == "no"  || s == "false"|| s == "0") p.Set(0);
                    }
                    else if (int.TryParse(newVal, NumberStyles.Any, CultureInfo.InvariantCulture, out int i))
                    {
                        p.Set(i);
                    }
                    break;
            }
        }

        #endregion

        #region Main worker – RenameParameters (instance only) ------------------------------------

        /// <summary>
        /// Lets the user pick instance parameters and renames them.
        /// Two-pass scheme prevents temporary duplicates (e.g. when renumbering sheets).
        /// </summary>
        public static Result RenameInstanceParameters(Document doc, List<Element> elements)
        {
            try
            {
                // ---------- project units (needed for mm → ft conversion) ----------------------
                Units         projUnits = doc.GetUnits();
                FormatOptions lenOpts   = projUnits.GetFormatOptions(SpecTypeId.Length);
                bool isMetric = lenOpts.GetUnitTypeId() == UnitTypeId.Millimeters ||
                                lenOpts.GetUnitTypeId() == UnitTypeId.Meters;

                // helper – parameter ➔ string
                string AsString(Parameter pp) =>
                    pp == null ? string.Empty
                               : (pp.StorageType == StorageType.String
                                     ? (pp.AsString()      ?? string.Empty)
                                     : (pp.AsValueString() ?? string.Empty));

                // 1. Build list of editable parameters present on at least one element ----------
                List<ParameterInfo> pInfos =
                    elements.SelectMany(el =>
                    {
                        List<ParameterInfo> list = new List<ParameterInfo>();
                        foreach (Parameter p in el.Parameters)
                        {
                            if (p.IsReadOnly || p.StorageType == StorageType.ElementId) continue;
                            list.Add(new ParameterInfo
                            {
                                Name     = p.Definition.Name,
                                Group    = p.Definition.ParameterGroup.ToString(),
                                IsShared = p.IsShared
                            });
                        }
                        return list;
                    })
                    .GroupBy(info => info.Name)
                    .Select(g => g.First())
                    .OrderBy(info => info.Group)
                    .ThenBy(info => info.Name)
                    .ToList();

                if (pInfos.Count == 0)
                {
                    WinForms.MessageBox.Show("No editable parameters found.");
                    return Result.Cancelled;
                }

                // 2. Parameter pick UI (small grid) --------------------------------------------
                List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                foreach (ParameterInfo pi in pInfos)
                {
                    rows.Add(new Dictionary<string, object>
                    {
                        { "Group", pi.Group },
                        { "Name",  pi.Name  }
                    });
                }
                List<string> headers = new List<string> { "Group", "Name" };
                IList<Dictionary<string, object>> picked =
                    CustomGUIs.DataGrid(rows, headers, /*multiSelect=*/false);

                if (picked == null || picked.Count == 0)
                    return Result.Succeeded;

                // 3. Flatten to ParameterValueInfo list ----------------------------------------
                List<ParameterValueInfo> pVals = new List<ParameterValueInfo>();
                foreach (Element el in elements)
                {
                    foreach (Dictionary<string, object> d in picked)
                    {
                        string pName = d["Name"].ToString();
                        Parameter p  = el.LookupParameter(pName);
                        if (p != null && !p.IsReadOnly && p.StorageType != StorageType.ElementId)
                        {
                            pVals.Add(new ParameterValueInfo
                            {
                                Element       = el,
                                Parameter     = p,
                                ParameterName = pName,
                                CurrentValue  = AsString(p),
                                IsShared      = p.IsShared
                            });
                        }
                    }
                }
                if (pVals.Count == 0)
                {
                    WinForms.MessageBox.Show("No valid parameters in selection.");
                    return Result.Cancelled;
                }

                // 4. Rename UI + write-back -----------------------------------------------------
                using (RenameParamForm form = new RenameParamForm(pVals))
                {
                    if (form.ShowDialog() != WinForms.DialogResult.OK)
                        return Result.Succeeded;

                    // build update list BEFORE opening the transaction
                    var updates = new List<(ParameterValueInfo Pv, string NewVal)>();
                    foreach (ParameterValueInfo pv in pVals)
                    {
                        // Pass the element to TransformValue for parameter reference support
                        string newVal = TransformValue(pv.CurrentValue, form, pv.Element);
                        if (newVal != pv.CurrentValue)
                            updates.Add((pv, newVal));
                    }
                    if (updates.Count == 0) return Result.Succeeded;

                    using (Transaction tx = new Transaction(doc, "Rename Parameters"))
                    {
                        tx.Start();

                        // === PASS 1 – temporary unique placeholders (string params) ===============
                        foreach (var u in updates)
                        {
                            if (u.Pv.Parameter != null &&
                                u.Pv.Parameter.StorageType == StorageType.String)
                            {
                                SetParameterValue(u.Pv.Parameter,
                                                  $"TMP_{Guid.NewGuid():N}",
                                                  isMetric);
                            }
                        }

                        // === PASS 2 – final values ================================================
                        foreach (var u in updates)
                        {
                            try
                            {
                                SetParameterValue(u.Pv.Parameter, u.NewVal, isMetric);
                            }
                            catch (Exception ex)
                            {
                                WinForms.MessageBox.Show(ex.Message, "Error");
                            }
                        }

                        tx.Commit();
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show(ex.ToString(), "Critical Error");
                return Result.Failed;
            }
        }

        #endregion
    }

    // =============================================================================================
    // 2. WinForms preview dialog
    // =============================================================================================
    public class RenameParamForm : WinForms.Form
    {
        private readonly List<ParameterRenamerHelper.ParameterValueInfo> _paramVals;

        private WinForms.TextBox      _txtFind;
        private WinForms.TextBox      _txtReplace;
        private WinForms.TextBox      _txtPattern;
        private WinForms.TextBox      _txtMath;
        private WinForms.RichTextBox  _rtbBefore;
        private WinForms.RichTextBox  _rtbAfter;

        #region Exposed properties

        public string FindText          => _txtFind.Text;
        public string ReplaceText       => _txtReplace.Text;
        public string PatternText       => _txtPattern.Text;
        public string MathOperationText => _txtMath.Text;

        #endregion

        public RenameParamForm(List<ParameterRenamerHelper.ParameterValueInfo> paramVals)
        {
            _paramVals = paramVals;
            BuildUI();
            LoadCurrentValues();
        }

        private void BuildUI()
        {
            Text        = "Modify Parameters";
            Font        = new Drawing.Font("Segoe UI", 9);
            MinimumSize = new Drawing.Size(520, 480);
            Size        = new Drawing.Size(640, 560);
            KeyPreview  = true;
            KeyDown    += (s, e) => { if (e.KeyCode == WinForms.Keys.Escape) Close(); };

            // === layout ========================================================================
            WinForms.TableLayoutPanel grid = new WinForms.TableLayoutPanel
            {
                Dock        = WinForms.DockStyle.Fill,
                ColumnCount = 2,
                RowCount    = 8,
                Padding     = new WinForms.Padding(8)
            };
            grid.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 90));
            grid.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100));
            for (int r = 0; r < 6; ++r)
                grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, r % 2 == 0 ? 28 : 20));
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));

            // Find / Replace
            grid.Controls.Add(MakeLabel("Find:"), 0, 0);
            _txtFind    = MakeTextBox();
            grid.Controls.Add(_txtFind, 1, 0);

            grid.Controls.Add(MakeLabel("Replace:"), 0, 1);
            _txtReplace = MakeTextBox();
            grid.Controls.Add(_txtReplace, 1, 1);

            // Pattern
            grid.Controls.Add(MakeLabel("Pattern:"), 0, 2);
            _txtPattern = MakeTextBox("{}");   // default
            grid.Controls.Add(_txtPattern, 1, 2);

            grid.Controls.Add(MakeHint("Use {} for current value. Use $\"Parameter Name\" or $ParameterName to reference other parameters."), 1, 3);

            // Math
            grid.Controls.Add(MakeLabel("Math:"), 0, 4);
            _txtMath = MakeTextBox();
            grid.Controls.Add(_txtMath, 1, 4);

            grid.Controls.Add(MakeHint("Use x to represent current value (e.g. 2x, x/2, x+3, -x)."), 1, 5);

            // Before / After preview
            _rtbBefore = new WinForms.RichTextBox { ReadOnly = true, Dock = WinForms.DockStyle.Fill };
            WinForms.GroupBox grpBefore = new WinForms.GroupBox { Text = "Current Values", Dock = WinForms.DockStyle.Fill };
            grpBefore.Controls.Add(_rtbBefore);
            grid.Controls.Add(grpBefore, 0, 6);
            grid.SetColumnSpan(grpBefore, 2);

            _rtbAfter = new WinForms.RichTextBox { ReadOnly = true, Dock = WinForms.DockStyle.Fill };
            WinForms.GroupBox grpAfter = new WinForms.GroupBox { Text = "Preview", Dock = WinForms.DockStyle.Fill };
            grpAfter.Controls.Add(_rtbAfter);
            grid.Controls.Add(grpAfter, 0, 7);
            grid.SetColumnSpan(grpAfter, 2);

            // buttons
            WinForms.FlowLayoutPanel buttons = new WinForms.FlowLayoutPanel
            {
                FlowDirection = WinForms.FlowDirection.RightToLeft,
                Dock          = WinForms.DockStyle.Bottom,
                Padding       = new WinForms.Padding(8)
            };

            WinForms.Button btnOK     = new WinForms.Button { Text = "OK",     DialogResult = WinForms.DialogResult.OK };
            WinForms.Button btnCancel = new WinForms.Button { Text = "Cancel", DialogResult = WinForms.DialogResult.Cancel };
            buttons.Controls.Add(btnOK);
            buttons.Controls.Add(btnCancel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            // events
            _txtFind.TextChanged    += (s, e) => RefreshPreview();
            _txtReplace.TextChanged += (s, e) => RefreshPreview();
            _txtPattern.TextChanged += (s, e) => RefreshPreview();
            _txtMath.TextChanged    += (s, e) => RefreshPreview();

            Controls.Add(grid);
            Controls.Add(buttons);
        }

        private static WinForms.Label MakeLabel(string txt) =>
            new WinForms.Label
            {
                Text      = txt,
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock      = WinForms.DockStyle.Fill
            };

        private static WinForms.TextBox MakeTextBox(string initial = "") =>
            new WinForms.TextBox { Text = initial, Dock = WinForms.DockStyle.Fill };

        private static WinForms.Label MakeHint(string txt) =>
            new WinForms.Label
            {
                Text      = txt,
                Dock      = WinForms.DockStyle.Fill,
                ForeColor = Drawing.Color.Gray,
                Font      = new Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Italic)
            };

        private void LoadCurrentValues()
        {
            foreach (var pv in _paramVals)
                _rtbBefore.AppendText(pv.CurrentValue + Environment.NewLine);

            RefreshPreview();
        }

        private void RefreshPreview()
        {
            _rtbAfter.Clear();
            foreach (var pv in _paramVals)
            {
                string v = pv.CurrentValue;

                // Apply same transformations as helper
                if (!string.IsNullOrEmpty(FindText))
                    v = v.Replace(FindText, ReplaceText);
                else if (!string.IsNullOrEmpty(ReplaceText))
                    v = ReplaceText;

                if (!string.IsNullOrEmpty(PatternText))
                {
                    // For preview, we can't access other parameters, so show placeholder
                    v = PatternText.Replace("{}", v);
                    // Show parameter references as-is in preview
                    var regex = new Regex(@"\$""([^""]+)""|(?<!\$)\$(\w+)");
                    v = regex.Replace(v, match => "[" + match.Value + "]");
                }

                if (!string.IsNullOrEmpty(MathOperationText))
                {
                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                        v = ParameterRenamerHelper.ApplyMathOperation(n, MathOperationText).ToString(CultureInfo.InvariantCulture);
                    else
                        v = ParameterRenamerHelper.ApplyMathToNumbersInString(v, MathOperationText);
                }

                _rtbAfter.AppendText(v + Environment.NewLine);
            }
        }
    }

    // =============================================================================================
    // 3. External command – rename instance parameters of selection
    // =============================================================================================
    [Transaction(TransactionMode.Manual)]
    public class RenameInstanceParametersOfSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData cData, ref string msg, ElementSet set)
        {
            UIDocument uiDoc = cData.Application.ActiveUIDocument;
            Document   doc   = uiDoc.Document;

            IList<ElementId> selIds = uiDoc.GetSelectionIds().ToList();
            if (selIds.Count == 0)
            {
                WinForms.MessageBox.Show("Please select elements first.");
                return Result.Cancelled;
            }

            List<Element> elems = selIds
                                  .Select(id => doc.GetElement(id))
                                  .Where(e => e != null)
                                  .ToList();

            return ParameterRenamerHelper.RenameInstanceParameters(doc, elems);
        }
    }
}
