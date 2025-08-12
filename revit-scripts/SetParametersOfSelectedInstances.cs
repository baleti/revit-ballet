#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

[Transaction(TransactionMode.Manual)]
public class SetParametersOfSelectedInstances : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the Revit Document and UI Document
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get currently selected elements
        var selectedIds = uidoc.GetSelectionIds();
        if (!selectedIds.Any())
        {
            TaskDialog.Show("No Selection", "Please select one or more elements.");
            return Result.Cancelled;
        }

        List<Element> selectedElements = selectedIds.Select(id => doc.GetElement(id)).Where(e => e != null).ToList();

        // Gather instance parameters from all selected elements
        // Only include non-read-only parameters
        var allParameters = new List<ParameterEntryForMultiple>();
        foreach (var elem in selectedElements)
        {
            string elementName = GetElementDisplayName(elem);
            foreach (Parameter param in elem.Parameters)
            {
                if (!param.IsReadOnly)
                {
                    allParameters.Add(new ParameterEntryForMultiple
                    {
                        Element = elem,
                        ElementName = elementName,
                        ParameterName = param.Definition.Name,
                        CurrentValue = GetParameterValue(param),
                        Parameter = param
                    });
                }
            }
        }

        if (!allParameters.Any())
        {
            TaskDialog.Show("No Editable Parameters", "The selected elements have no editable instance parameters.");
            return Result.Cancelled;
        }

        // Sort by ElementName so they're grouped by their originating element
        allParameters = allParameters.OrderBy(p => p.ElementName).ToList();

        // Show DataGrid: ElementName, ParameterName, CurrentValue
        var chosenParameterEntries = CustomGUIs.DataGrid(allParameters, new List<string> { "ElementName", "ParameterName", "CurrentValue" }, null, "Select Instance Parameters");
        if (!chosenParameterEntries.Any()) return Result.Cancelled;

        // Ask user for new value
        string newValue = ShowInputDialog("Set Parameter Value", "Enter new value for the selected parameters:");
        if (newValue == null) return Result.Cancelled; // User cancelled

        // Set the parameter's value for each selected parameter/element
        using (Transaction tx = new Transaction(doc, "Set Instance Parameters"))
        {
            tx.Start();
            foreach (var paramEntry in chosenParameterEntries)
            {
                if (!SetParameterValue(paramEntry.Parameter, newValue))
                {
                    TaskDialog.Show("Error", $"Could not set parameter '{paramEntry.ParameterName}' on '{paramEntry.ElementName}' to the specified value.");
                    tx.RollBack();
                    return Result.Cancelled;
                }
            }
            tx.Commit();
        }

        TaskDialog.Show("Success", "Parameter values updated successfully for all selected parameters on the selected elements.");
        return Result.Succeeded;
    }

    private string GetElementDisplayName(Element elem)
    {
        // Try a known parameter
        var nameParam = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (nameParam != null && !string.IsNullOrEmpty(nameParam.AsString()))
            return nameParam.AsString();

        // If that fails, fallback to the Element's Name property
        return elem.Name;
    }

    private string GetParameterValue(Parameter param)
    {
        switch (param.StorageType)
        {
            case StorageType.String:
                return param.AsString() ?? "";
            case StorageType.Double:
                return param.AsDouble().ToString();
            case StorageType.Integer:
                return param.AsInteger().ToString();
            case StorageType.ElementId:
                ElementId id = param.AsElementId();
                if (id == ElementId.InvalidElementId) return "";
                return id.IntegerValue.ToString();
            default:
                return "";
        }
    }

    private bool SetParameterValue(Parameter param, string value)
    {
        try
        {
            switch (param.StorageType)
            {
                case StorageType.String:
                    param.Set(value);
                    return true;
                case StorageType.Double:
                    if (double.TryParse(value, out double dVal))
                    {
                        param.Set(dVal);
                        return true;
                    }
                    return false;
                case StorageType.Integer:
                    if (int.TryParse(value, out int iVal))
                    {
                        param.Set(iVal);
                        return true;
                    }
                    return false;
                case StorageType.ElementId:
                    if (int.TryParse(value, out int elemIdVal))
                    {
                        param.Set(new ElementId(elemIdVal));
                        return true;
                    }
                    return false;
                default:
                    return false;
            }
        }
        catch
        {
            return false;
        }
    }

    private string ShowInputDialog(string title, string promptText)
    {
        System.Windows.Forms.Form form = new System.Windows.Forms.Form();
        Label label = new Label();
        System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
        Button buttonOk = new Button();
        Button buttonCancel = new Button();

        form.Text = title;
        label.Text = promptText;
        textBox.Text = "";

        label.SetBounds(9, 20, 372, 13);
        textBox.SetBounds(12, 50, 372, 20);
        buttonOk.SetBounds(228, 80, 75, 23);
        buttonCancel.SetBounds(309, 80, 75, 23);

        buttonOk.Text = "OK";
        buttonCancel.Text = "Cancel";
        buttonOk.DialogResult = DialogResult.OK;
        buttonCancel.DialogResult = DialogResult.Cancel;

        label.AutoSize = true;
        textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
        buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

        form.ClientSize = new System.Drawing.Size(396, 120);
        form.Controls.AddRange(new System.Windows.Forms.Control[] { label, textBox, buttonOk, buttonCancel });
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
        form.AcceptButton = buttonOk;
        form.CancelButton = buttonCancel;

        DialogResult dialogResult = form.ShowDialog();
        return dialogResult == DialogResult.OK ? textBox.Text : null;
    }

    public class ParameterEntryForMultiple
    {
        public Element Element { get; set; }
        public string ElementName { get; set; }
        public string ParameterName { get; set; }
        public string CurrentValue { get; set; }
        public Parameter Parameter { get; set; }
    }
}
