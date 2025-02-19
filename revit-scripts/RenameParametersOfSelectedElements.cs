using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

// Alias WinForms and Drawing to avoid type ambiguities.
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace MyRevitAddin
{
    /// <summary>
    /// Contains the shared logic for renaming parameters.
    /// </summary>
    public static class ParameterRenamerHelper
    {
        public class ParameterInfo
        {
            public string Name { get; set; }
            public bool IsShared { get; set; }
            /// <summary>
            /// A string representing the parameter’s group (for example, Constraints, Dimensions, Identity Data, etc.)
            /// </summary>
            public string Group { get; set; }
        }

        public class ParameterValueInfo
        {
            /// <summary>
            /// The element whose “parameter” is to be updated.
            /// For a pseudo parameter (such as “Type Name”), this is still the element.
            /// </summary>
            public Element Element { get; set; }
            /// <summary>
            /// The actual Revit Parameter. If null, this ParameterValueInfo represents a pseudo parameter.
            /// </summary>
            public Parameter Parameter { get; set; }
            /// <summary>
            /// The display name of the parameter.
            /// </summary>
            public string ParameterName { get; set; }
            /// <summary>
            /// The current value (as text) to be modified.
            /// </summary>
            public string CurrentValue { get; set; }
            public bool IsShared { get; set; }
        }

        /// <summary>
        /// Applies a simple mathematical operation on the input value using the given string.
        /// Supported operations include:
        ///   - "2x"  (multiplies x by 2)
        ///   - "x/2" (divides x by 2)
        ///   - "x+2" (adds 2)
        ///   - "x-2" (subtracts 2)
        ///   - "-x"  (negates x)
        /// The operation string is case-insensitive and should have no spaces.
        /// If the string is not recognized, the original value is returned.
        /// </summary>
        public static double ApplyMathOperation(double x, string mathOp)
        {
            if (string.IsNullOrWhiteSpace(mathOp))
                return x;

            mathOp = mathOp.Replace(" ", ""); // remove spaces
            double result = x;
            if (mathOp.Equals("x", StringComparison.OrdinalIgnoreCase))
            {
                return x;
            }
            else if (mathOp.Equals("-x", StringComparison.OrdinalIgnoreCase))
            {
                return -x;
            }
            else if (mathOp.EndsWith("x", StringComparison.OrdinalIgnoreCase) && !mathOp.StartsWith("x"))
            {
                // e.g. "2x"
                string multStr = mathOp.Substring(0, mathOp.Length - 1);
                if (double.TryParse(multStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double mult))
                    return mult * x;
            }
            else if (mathOp.StartsWith("x", StringComparison.OrdinalIgnoreCase))
            {
                string opPart = mathOp.Substring(1);
                if (opPart.StartsWith("+"))
                {
                    string num = opPart.Substring(1);
                    if (double.TryParse(num, NumberStyles.Any, CultureInfo.InvariantCulture, out double addVal))
                        return x + addVal;
                }
                else if (opPart.StartsWith("-"))
                {
                    string num = opPart.Substring(1);
                    if (double.TryParse(num, NumberStyles.Any, CultureInfo.InvariantCulture, out double subVal))
                        return x - subVal;
                }
                else if (opPart.StartsWith("*"))
                {
                    string num = opPart.Substring(1);
                    if (double.TryParse(num, NumberStyles.Any, CultureInfo.InvariantCulture, out double mult))
                        return x * mult;
                }
                else if (opPart.StartsWith("/"))
                {
                    string num = opPart.Substring(1);
                    if (double.TryParse(num, NumberStyles.Any, CultureInfo.InvariantCulture, out double div) && div != 0)
                        return x / div;
                }
            }
            return x;
        }

        /// <summary>
        /// Determines (via reflection) if a parameter is a Yes/No parameter.
        /// In Revit 2024 the internal ParameterType string may be "Yes/No",
        /// so we normalize the value before comparing.
        /// </summary>
        public static bool IsYesNoParameter(Parameter param)
        {
            if (param == null || param.Definition == null)
                return false;

            // Use reflection to get the (non-public) ParameterType property.
            var prop = param.Definition.GetType().GetProperty("ParameterType",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (prop != null)
            {
                var value = prop.GetValue(param.Definition, null);
                if (value != null)
                {
                    // Remove any slashes and spaces so "Yes/No" becomes "YesNo"
                    string typeName = value.ToString().Replace("/", "").Replace(" ", "");
                    if (typeName.Equals("YesNo", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Transforms the original string value using the provided form values.
        /// If a Find text is provided, it replaces it with the Replace text.
        /// Otherwise, if Replace text is provided (and Find is empty), that value is used.
        /// Then if a Pattern is provided, it is applied (by replacing "{}" with the current value).
        /// Finally, if a Math value is provided and the result is numeric, it is applied.
        /// </summary>
        private static string TransformValue(string original, RenameParamForm form)
        {
            string newValue = original;
            if (!string.IsNullOrEmpty(form.FindText))
            {
                newValue = newValue.Replace(form.FindText, form.ReplaceText);
            }
            else if (!string.IsNullOrEmpty(form.ReplaceText))
            {
                newValue = form.ReplaceText;
            }
            if (!string.IsNullOrEmpty(form.PatternText))
            {
                newValue = form.PatternText.Replace("{}", newValue);
            }
            if (!string.IsNullOrEmpty(form.MathOperationText) &&
                double.TryParse(newValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double x))
            {
                double result = ApplyMathOperation(x, form.MathOperationText);
                newValue = result.ToString(CultureInfo.InvariantCulture);
            }
            return newValue;
        }

        /// <summary>
        /// Shared routine for renaming parameter values.
        /// 
        /// - For type elements (includeTypeName==true) the routine builds a transposed grid
        ///   (rows = parameters; columns = family type names) and then, after the user selects
        ///   one or more rows, it converts the selection to a flat list of ParameterValueInfo objects
        ///   and calls the standard renaming dialog.
        /// 
        /// - For instance elements (includeTypeName==false) we now collect all editable parameters
        ///   (not just string parameters) from the selected instance elements. The union is sorted by
        ///   parameter group then by name. The grid shows two columns ("Group" and "Name") so the user
        ///   may select any instance parameter – such as constraints, dimensions, phasing, etc.
        ///   Then we build a flat list of ParameterValueInfo objects (one for each element/parameter)
        ///   and show the same renaming dialog.
        /// 
        /// IMPORTANT: Before updating a numerical parameter (Double), if the parameter represents a length
        /// and if the project’s length unit (obtained via doc.GetUnits().GetFormatOptions(SpecTypeId.Length)) 
        /// indicates metric (i.e. its unit type is UnitTypeId.Millimeters or UnitTypeId.Meters), then we assume 
        /// the user‐entered value is in millimeters and convert it to feet (by dividing by 304.8) before calling Set.
        /// </summary>
        public static Result RenameParameters(Document doc, List<Element> elements, bool includeTypeName = false)
        {
            try
            {
                // Determine if the project is metric.
                Units projectUnits = doc.GetUnits();
                FormatOptions lengthOptions = projectUnits.GetFormatOptions(SpecTypeId.Length);
                bool isMetric = (lengthOptions.GetUnitTypeId() == UnitTypeId.Millimeters ||
                                 lengthOptions.GetUnitTypeId() == UnitTypeId.Meters);

                if (includeTypeName)
                {
                    // === TYPE PARAMETERS (Matrix/Transposed View) ===
                    var unionParams = new Dictionary<(string Group, string Name), ParameterInfo>();

                    foreach (var elem in elements)
                    {
                        foreach (Parameter p in elem.Parameters)
                        {
                            if (!p.IsReadOnly && p.StorageType != StorageType.ElementId)
                            {
                                string paramName = p.Definition.Name;
                                string group = p.Definition.ParameterGroup.ToString();
                                var key = (group, paramName);
                                if (!unionParams.ContainsKey(key))
                                {
                                    unionParams[key] = new ParameterInfo
                                    {
                                        Name = paramName,
                                        IsShared = p.IsShared,
                                        Group = group
                                    };
                                }
                            }
                        }
                        // Add pseudo parameter "Type Name"
                        var key2 = ("Identity Data", "Type Name");
                        if (!unionParams.ContainsKey(key2))
                        {
                            unionParams[key2] = new ParameterInfo
                            {
                                Name = "Type Name",
                                IsShared = false,
                                Group = "Identity Data"
                            };
                        }
                    }

                    var sortedParamInfos = unionParams.Values
                        .OrderBy(pi => pi.Group)
                        .ThenBy(pi => pi.Name)
                        .ToList();

                    var typeNames = elements.Select(e => e.Name).ToList();
                    var propertyNames = new List<string> { "Group", "Name" };
                    propertyNames.AddRange(typeNames);

                    var gridEntries = new List<Dictionary<string, object>>();
                    foreach (var pi in sortedParamInfos)
                    {
                        var row = new Dictionary<string, object>();
                        row["Group"] = pi.Group;
                        row["Name"] = pi.Name;
                        foreach (var elem in elements)
                        {
                            string value = "";
                            if (pi.Name == "Type Name")
                            {
                                value = elem.Name;
                            }
                            else
                            {
                                var p = elem.LookupParameter(pi.Name);
                                if (p != null && !p.IsReadOnly && p.StorageType != StorageType.ElementId)
                                {
                                    value = p.StorageType == StorageType.String ? p.AsString() ?? "" : p.AsValueString() ?? "";
                                }
                            }
                            row[elem.Name] = value;
                        }
                        gridEntries.Add(row);
                    }

                    var selectedRows = CustomGUIs.DataGrid(gridEntries, propertyNames, false);
                    if (selectedRows == null || selectedRows.Count == 0)
                        return Result.Succeeded;

                    var paramValues = new List<ParameterValueInfo>();
                    foreach (var row in selectedRows)
                    {
                        string paramName = row["Name"]?.ToString();
                        foreach (var elem in elements)
                        {
                            if (paramName == "Type Name")
                            {
                                paramValues.Add(new ParameterValueInfo
                                {
                                    Element = elem,
                                    Parameter = null,
                                    ParameterName = "Type Name",
                                    CurrentValue = elem.Name,
                                    IsShared = false
                                });
                            }
                            else
                            {
                                var p = elem.LookupParameter(paramName);
                                if (p != null && !p.IsReadOnly && p.StorageType != StorageType.ElementId)
                                {
                                    string currentVal = p.StorageType == StorageType.String ? p.AsString() ?? "" : p.AsValueString() ?? "";
                                    paramValues.Add(new ParameterValueInfo
                                    {
                                        Element = elem,
                                        Parameter = p,
                                        ParameterName = paramName,
                                        CurrentValue = currentVal,
                                        IsShared = p.IsShared
                                    });
                                }
                            }
                        }
                    }

                    if (paramValues.Count == 0)
                    {
                        WinForms.MessageBox.Show("No valid parameters found in selection.", "Error");
                        return Result.Cancelled;
                    }

                    using (var form = new RenameParamForm(paramValues))
                    {
                        if (form.ShowDialog() != WinForms.DialogResult.OK)
                            return Result.Succeeded;

                        using (var tx = new Transaction(doc, "Rename Type Parameters"))
                        {
                            tx.Start();
                            foreach (var pv in paramValues)
                            {
                                string newValue = TransformValue(pv.CurrentValue, form);

                                if (newValue != pv.CurrentValue)
                                {
                                    if (pv.Parameter == null)
                                    {
                                        try { pv.Element.Name = newValue; }
                                        catch (Exception ex)
                                        {
                                            WinForms.MessageBox.Show($"Failed to update Type Name for {pv.Element.Id}:\n{ex.Message}", "Error");
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            if (pv.Parameter.StorageType == StorageType.String)
                                            {
                                                pv.Parameter.Set(newValue);
                                            }
                                            else if (pv.Parameter.StorageType == StorageType.Double)
                                            {
                                                if (double.TryParse(newValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double newDouble))
                                                {
                                                    // Convert metric length values from mm to feet.
                                                    if ((pv.Parameter.GetUnitTypeId() == UnitTypeId.Millimeters ||
                                                         pv.Parameter.GetUnitTypeId() == UnitTypeId.Meters) && isMetric)
                                                        newDouble = newDouble / 304.8;
                                                    pv.Parameter.Set(newDouble);
                                                }
                                            }
                                            else if (pv.Parameter.StorageType == StorageType.Integer)
                                            {
                                                // For Yes/No parameters, use our reflection-based check.
                                                if (IsYesNoParameter(pv.Parameter))
                                                {
                                                    string lower = newValue.Trim().ToLowerInvariant();
                                                    bool? boolValue = null;
                                                    if (lower == "yes" || lower == "true" || lower == "1")
                                                        boolValue = true;
                                                    else if (lower == "no" || lower == "false" || lower == "0")
                                                        boolValue = false;
                                                    if (boolValue.HasValue)
                                                        pv.Parameter.Set(boolValue.Value ? 1 : 0);
                                                }
                                                else
                                                {
                                                    if (int.TryParse(newValue, NumberStyles.Any, CultureInfo.InvariantCulture, out int newInt))
                                                    {
                                                        pv.Parameter.Set(newInt);
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            WinForms.MessageBox.Show($"Failed to update parameter {pv.ParameterName} on element {pv.Element.Id}:\n{ex.Message}", "Error");
                                        }
                                    }
                                }
                            }
                            tx.Commit();
                        }
                    }
                    return Result.Succeeded;
                }
                else
                {
                    // === INSTANCE PARAMETERS (Standard View) ===
                    var paramInfos = elements
                        .SelectMany(elem => elem.Parameters.Cast<Parameter>())
                        .Where(p => !p.IsReadOnly && p.StorageType != StorageType.ElementId)
                        .GroupBy(p => p.Definition.Name)
                        .Select(g => new ParameterInfo
                        {
                            Name = g.Key,
                            IsShared = g.First().IsShared,
                            Group = g.First().Definition.ParameterGroup.ToString()
                        })
                        .OrderBy(pi => pi.Group)
                        .ThenBy(pi => pi.Name)
                        .ToList();

                    // Add pseudo‑parameters only if at least one FamilyInstance supports the flip.
                    if (elements.Any(e => e is FamilyInstance fi && fi.CanFlipFacing))
                    {
                        paramInfos.Add(new ParameterInfo { Name = "FacingFlipped", IsShared = false, Group = "FamilyInstance Properties" });
                    }
                    if (elements.Any(e => e is FamilyInstance fi && fi.CanFlipHand))
                    {
                        paramInfos.Add(new ParameterInfo { Name = "HandFlipped", IsShared = false, Group = "FamilyInstance Properties" });
                    }

                    if (paramInfos.Count == 0)
                    {
                        WinForms.MessageBox.Show("No editable parameters found.", "Error");
                        return Result.Cancelled;
                    }

                    var gridEntries = paramInfos.Select(pi => new Dictionary<string, object>
                    {
                        { "Group", pi.Group },
                        { "Name", pi.Name }
                    }).ToList();

                    var selectedParams = CustomGUIs.DataGrid(gridEntries, new List<string> { "Group", "Name" }, false);
                    if (selectedParams == null || selectedParams.Count == 0)
                        return Result.Succeeded;

                    var paramValues = new List<ParameterValueInfo>();
                    foreach (var elem in elements)
                    {
                        foreach (var dict in selectedParams)
                        {
                            string paramName = dict["Name"]?.ToString();
                            // Handle the pseudo‑parameters explicitly.
                            if (paramName == "FacingFlipped")
                            {
                                if (elem is FamilyInstance fi && fi.CanFlipFacing)
                                {
                                    string currentVal = fi.FacingFlipped ? "True" : "False";
                                    paramValues.Add(new ParameterValueInfo
                                    {
                                        Element = elem,
                                        Parameter = null,
                                        ParameterName = "FacingFlipped",
                                        CurrentValue = currentVal,
                                        IsShared = false
                                    });
                                }
                            }
                            else if (paramName == "HandFlipped")
                            {
                                if (elem is FamilyInstance fi && fi.CanFlipHand)
                                {
                                    string currentVal = fi.HandFlipped ? "True" : "False";
                                    paramValues.Add(new ParameterValueInfo
                                    {
                                        Element = elem,
                                        Parameter = null,
                                        ParameterName = "HandFlipped",
                                        CurrentValue = currentVal,
                                        IsShared = false
                                    });
                                }
                            }
                            else
                            {
                                var p = elem.LookupParameter(paramName);
                                if (p != null && !p.IsReadOnly && p.StorageType != StorageType.ElementId)
                                {
                                    string currentVal = p.StorageType == StorageType.String ? p.AsString() ?? "" : p.AsValueString() ?? "";
                                    paramValues.Add(new ParameterValueInfo
                                    {
                                        Element = elem,
                                        Parameter = p,
                                        ParameterName = paramName,
                                        CurrentValue = currentVal,
                                        IsShared = p.IsShared
                                    });
                                }
                            }
                        }
                    }

                    if (paramValues.Count == 0)
                    {
                        WinForms.MessageBox.Show("No valid parameters found in selection.", "Error");
                        return Result.Cancelled;
                    }

                    using (var form = new RenameParamForm(paramValues))
                    {
                        if (form.ShowDialog() != WinForms.DialogResult.OK)
                            return Result.Succeeded;

                        using (var tx = new Transaction(doc, "Rename Parameters"))
                        {
                            tx.Start();
                            foreach (var pv in paramValues)
                            {
                                string newValue = TransformValue(pv.CurrentValue, form);
                                if (newValue != pv.CurrentValue)
                                {
                                    try
                                    {
                                        if (pv.Parameter == null)
                                        {
                                            if (pv.ParameterName == "FacingFlipped" || pv.ParameterName == "HandFlipped")
                                            {
                                                if (pv.Element is FamilyInstance fi)
                                                {
                                                    string lower = newValue.Trim().ToLowerInvariant();
                                                    bool? desiredValue = null;
                                                    if (lower == "yes" || lower == "true" || lower == "1")
                                                        desiredValue = true;
                                                    else if (lower == "no" || lower == "false" || lower == "0")
                                                        desiredValue = false;
                                                    if (desiredValue.HasValue)
                                                    {
                                                        if (pv.ParameterName == "FacingFlipped")
                                                        {
                                                            if (fi.FacingFlipped != desiredValue.Value && fi.CanFlipFacing)
                                                                fi.flipFacing();
                                                        }
                                                        else if (pv.ParameterName == "HandFlipped")
                                                        {
                                                            if (fi.HandFlipped != desiredValue.Value && fi.CanFlipHand)
                                                                fi.flipHand();
                                                        }
                                                    }
                                                }
                                            }
                                            else if (pv.ParameterName == "Type Name")
                                            {
                                                pv.Element.Name = newValue;
                                            }
                                        }
                                        else if (pv.Parameter.StorageType == StorageType.String)
                                        {
                                            pv.Parameter.Set(newValue);
                                        }
                                        else if (pv.Parameter.StorageType == StorageType.Double)
                                        {
                                            if (double.TryParse(newValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double newDouble))
                                            {
                                                if ((pv.Parameter.GetUnitTypeId() == UnitTypeId.Millimeters ||
                                                     pv.Parameter.GetUnitTypeId() == UnitTypeId.Meters) && isMetric)
                                                    newDouble = newDouble / 304.8;
                                                pv.Parameter.Set(newDouble);
                                            }
                                        }
                                        else if (pv.Parameter.StorageType == StorageType.Integer)
                                        {
                                            if (IsYesNoParameter(pv.Parameter))
                                            {
                                                string lower = newValue.Trim().ToLowerInvariant();
                                                bool? boolValue = null;
                                                if (lower == "yes" || lower == "true" || lower == "1")
                                                    boolValue = true;
                                                else if (lower == "no" || lower == "false" || lower == "0")
                                                    boolValue = false;
                                                if (boolValue.HasValue)
                                                    pv.Parameter.Set(boolValue.Value ? 1 : 0);
                                            }
                                            else
                                            {
                                                if (int.TryParse(newValue, NumberStyles.Any, CultureInfo.InvariantCulture, out int newInt))
                                                {
                                                    pv.Parameter.Set(newInt);
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        WinForms.MessageBox.Show($"Failed to update {pv.Element.Name}:\n{ex.Message}", "Error");
                                    }
                                }
                            }
                            tx.Commit();
                        }
                    }
                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show(ex.ToString(), "Critical Error");
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// The form used for previewing and applying parameter renaming.
    /// This dialog is used by both instance and type parameter commands.
    /// </summary>
    public class RenameParamForm : WinForms.Form
    {
        private readonly List<ParameterRenamerHelper.ParameterValueInfo> _paramValues;
        private WinForms.TextBox _txtFind;
        private WinForms.TextBox _txtReplace;
        private WinForms.TextBox _txtPattern;
        private WinForms.TextBox _txtMathOperation; // Renamed field ("Math")
        private WinForms.RichTextBox _rtbBefore;
        private WinForms.RichTextBox _rtbAfter;

        public string FindText => _txtFind.Text;
        public string ReplaceText => _txtReplace.Text;
        public string PatternText => _txtPattern.Text;
        public string MathOperationText => _txtMathOperation.Text;

        public RenameParamForm(List<ParameterRenamerHelper.ParameterValueInfo> paramValues)
        {
            _paramValues = paramValues;
            InitializeComponent();
            InitializeData();
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
            this.Text = "Modify Parameters";
            this.Size = new Drawing.Size(600, 600);
            this.MinimumSize = new Drawing.Size(500, 400);
            this.Font = new Drawing.Font("Segoe UI", 9);
            this.KeyPreview = true;

            // Updated layout: 8 rows total.
            // Rows:
            // 0: Find field
            // 1: Replace field
            // 2: Pattern label & textbox
            // 3: Pattern hint ("Use {} to represent current value.")
            // 4: Math label (renamed from "Math Op") & textbox
            // 5: Math hint ("Use x to represent current value (e.g., '2x', 'x/2', 'x+2', 'x-2', '-x').")
            // 6: Current Values (group box)
            // 7: Preview (group box)
            var mainLayout = new WinForms.TableLayoutPanel
            {
                Dock = WinForms.DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 8,
                Padding = new WinForms.Padding(10)
            };
            mainLayout.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 80));
            mainLayout.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100));
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30)); // Row 0: Find
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30)); // Row 1: Replace
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30)); // Row 2: Pattern
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 20)); // Row 3: Pattern Hint
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30)); // Row 4: Math
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 20)); // Row 5: Math Hint
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));  // Row 6: Current Values
            mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));  // Row 7: Preview

            // Row 0: Find field
            var lblFind = new WinForms.Label
            {
                Text = "Find:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock = WinForms.DockStyle.Fill
            };
            mainLayout.Controls.Add(lblFind, 0, 0);
            _txtFind = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            mainLayout.Controls.Add(_txtFind, 1, 0);

            // Row 1: Replace field
            var lblReplace = new WinForms.Label
            {
                Text = "Replace:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock = WinForms.DockStyle.Fill
            };
            mainLayout.Controls.Add(lblReplace, 0, 1);
            _txtReplace = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            mainLayout.Controls.Add(_txtReplace, 1, 1);

            // Row 2: Pattern field
            var lblPattern = new WinForms.Label
            {
                Text = "Pattern:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock = WinForms.DockStyle.Fill
            };
            mainLayout.Controls.Add(lblPattern, 0, 2);
            _txtPattern = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill, Text = "{}" };
            mainLayout.Controls.Add(_txtPattern, 1, 2);

            // Row 3: Pattern hint
            var lblPatternHint = new WinForms.Label
            {
                Text = "Use {} to represent current value.",
                Font = new Drawing.Font(this.Font, Drawing.FontStyle.Italic),
                ForeColor = Drawing.Color.Gray,
                Dock = WinForms.DockStyle.Fill
            };
            mainLayout.Controls.Add(lblPatternHint, 1, 3);

            // Row 4: Math field (renamed from "Math Op" to "Math")
            var lblMath = new WinForms.Label
            {
                Text = "Math:",
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock = WinForms.DockStyle.Fill
            };
            mainLayout.Controls.Add(lblMath, 0, 4);
            _txtMathOperation = new WinForms.TextBox { Dock = WinForms.DockStyle.Fill };
            mainLayout.Controls.Add(_txtMathOperation, 1, 4);

            // Row 5: Math hint
            var lblMathHint = new WinForms.Label
            {
                Text = "Use x to represent current value (e.g., '2x', 'x/2', 'x+2', 'x-2', '-x').",
                Font = new Drawing.Font(this.Font, Drawing.FontStyle.Italic),
                ForeColor = Drawing.Color.Gray,
                Dock = WinForms.DockStyle.Fill
            };
            mainLayout.Controls.Add(lblMathHint, 1, 5);

            // Row 6: Current Values
            _rtbBefore = new WinForms.RichTextBox { ReadOnly = true, Dock = WinForms.DockStyle.Fill };
            var groupBefore = new WinForms.GroupBox { Text = "Current Values", Dock = WinForms.DockStyle.Fill };
            groupBefore.Controls.Add(_rtbBefore);
            mainLayout.Controls.Add(groupBefore, 0, 6);
            mainLayout.SetColumnSpan(groupBefore, 2);

            // Row 7: Preview
            _rtbAfter = new WinForms.RichTextBox { ReadOnly = true, Dock = WinForms.DockStyle.Fill };
            var groupAfter = new WinForms.GroupBox { Text = "Preview", Dock = WinForms.DockStyle.Fill };
            groupAfter.Controls.Add(_rtbAfter);
            mainLayout.Controls.Add(groupAfter, 0, 7);
            mainLayout.SetColumnSpan(groupAfter, 2);

            // Button Panel at the bottom.
            var buttonPanel = new WinForms.Panel { Dock = WinForms.DockStyle.Bottom, Height = 40 };
            var btnCancel = new WinForms.Button { Text = "Cancel", Size = new Drawing.Size(80, 30), DialogResult = WinForms.DialogResult.Cancel };
            buttonPanel.Controls.Add(btnCancel);
            btnCancel.Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right;
            var btnOk = new WinForms.Button { Text = "OK", Size = new Drawing.Size(80, 30), DialogResult = WinForms.DialogResult.OK };
            buttonPanel.Controls.Add(btnOk);
            btnOk.Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right;
            btnOk.Location = new Drawing.Point(buttonPanel.Width - btnOk.Width - 10, buttonPanel.Height - btnOk.Height - 5);
            btnCancel.Location = new Drawing.Point(btnOk.Left - btnCancel.Width - 10, btnOk.Top);

            // Update preview when any textbox changes.
            _txtFind.TextChanged += UpdatePreview;
            _txtReplace.TextChanged += UpdatePreview;
            _txtPattern.TextChanged += UpdatePreview;
            _txtMathOperation.TextChanged += UpdatePreview;

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;

            this.Controls.Add(mainLayout);
            this.Controls.Add(buttonPanel);
        }

        private void InitializeData()
        {
            foreach (var pv in _paramValues)
                _rtbBefore.AppendText($"{pv.CurrentValue}{Environment.NewLine}");
            UpdatePreview(this, EventArgs.Empty);
        }

        private void UpdatePreview(object sender, EventArgs e)
        {
            _rtbAfter.Clear();
            foreach (var pv in _paramValues)
            {
                string current = ApplyTransformations(pv.CurrentValue);
                _rtbAfter.AppendText($"{current}{Environment.NewLine}");
            }
        }

        private string ApplyTransformations(string input)
        {
            string output = input;
            if (!string.IsNullOrEmpty(FindText))
            {
                output = output.Replace(FindText, ReplaceText);
            }
            else if (!string.IsNullOrEmpty(ReplaceText))
            {
                output = ReplaceText;
            }
            if (!string.IsNullOrEmpty(PatternText))
            {
                output = PatternText.Replace("{}", output);
            }
            if (!string.IsNullOrEmpty(MathOperationText) &&
                double.TryParse(output, NumberStyles.Any, CultureInfo.InvariantCulture, out double x))
            {
                double result = ParameterRenamerHelper.ApplyMathOperation(x, MathOperationText);
                output = result.ToString(CultureInfo.InvariantCulture);
            }
            return output;
        }
    }

    /// <summary>
    /// Command for renaming instance parameters of selected elements.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RenameInstanceParametersOfSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            var selectedIds = uiDoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                WinForms.MessageBox.Show("Please select elements first.", "Error");
                return Result.Cancelled;
            }
            var selectedElements = selectedIds
                .Cast<ElementId>()
                .Select(id => doc.GetElement(id))
                .Where(e => e != null)
                .ToList();
            return ParameterRenamerHelper.RenameParameters(doc, selectedElements, includeTypeName: false);
        }
    }

    /// <summary>
    /// Command for renaming type parameters of selected elements (or types).
    /// In addition to the editable type parameters, this command now displays a transposed grid:
    /// each row shows a parameter (with its Group and Name) and each additional column shows the value
    /// for a given family type. The rows are sorted first by Group then by Name.
    /// After the grid selection, the same renaming dialog (with preview, find/replace, pattern, and math) is shown.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RenameTypeParametersOfSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            var selectedIds = uiDoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                WinForms.MessageBox.Show("Please select elements or types first.", "Error");
                return Result.Cancelled;
            }
            var selectedElements = selectedIds
                .Cast<ElementId>()
                .Select(id => doc.GetElement(id))
                .Where(e => e != null)
                .ToList();
            var typeElements = new List<Element>();
            foreach (var elem in selectedElements)
            {
                if (elem is ElementType)
                    typeElements.Add(elem);
                else
                {
                    Element typeElem = doc.GetElement(elem.GetTypeId());
                    if (typeElem != null)
                        typeElements.Add(typeElem);
                }
            }
            typeElements = typeElements.GroupBy(e => e.Id).Select(g => g.First()).ToList();
            if (typeElements.Count == 0)
            {
                WinForms.MessageBox.Show("No type elements found in selection.", "Error");
                return Result.Cancelled;
            }
            return ParameterRenamerHelper.RenameParameters(doc, typeElements, includeTypeName: true);
        }
    }
}
