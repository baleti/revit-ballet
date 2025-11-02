using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System;

[Transaction(TransactionMode.Manual)]
public class SetParameterValueToSelectedFamilyTypesInCurrentFamily : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the active document
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Check if running in the Family Editor
        if (!doc.IsFamilyDocument)
        {
            TaskDialog.Show("Error", "This command can only be run in the Family Editor.");
            return Result.Cancelled;
        }

        // Get the FamilyManager
        FamilyManager familyManager = doc.FamilyManager;
        var familyTypes = familyManager.Types.Cast<FamilyType>().ToList();

        if (familyTypes.Count == 0)
        {
            TaskDialog.Show("Error", "No family types found in the current family.");
            return Result.Cancelled;
        }

        // Prepare data for family types selection
        List<Dictionary<string, object>> typeEntries = familyTypes.Select(ft => new Dictionary<string, object>
        {
            { "Name", ft.Name }
        }).ToList();

        List<string> typePropertyNames = new List<string> { "Name" };

        // Prompt user to select family types
        var selectedTypeEntries = CustomGUIs.DataGrid(typeEntries, typePropertyNames, false);

        if (selectedTypeEntries == null || selectedTypeEntries.Count == 0)
        {
            TaskDialog.Show("Error", "No family types selected.");
            return Result.Cancelled;
        }

        // Get selected FamilyTypes
        var selectedFamilyTypes = selectedTypeEntries
            .Select(entry => familyTypes.FirstOrDefault(ft => ft.Name == entry["Name"].ToString()))
            .Where(ft => ft != null)
            .ToList();

        // Collect modifiable parameters from the family
        List<FamilyParameter> familyParameters = familyManager.Parameters
            .Cast<FamilyParameter>()
            .Where(p => !p.IsReadOnly)
            .ToList();

        if (familyParameters.Count == 0)
        {
            TaskDialog.Show("Error", "No writable parameters found in the current family.");
            return Result.Cancelled;
        }

        // Prepare data for parameter selection
        List<Dictionary<string, object>> paramEntries = familyParameters.Select(p => new Dictionary<string, object>
        {
            { "Name", p.Definition.Name }
        }).ToList();

        List<string> paramPropertyNames = new List<string> { "Name" };

        // Prompt user to select a parameter
        var selectedParamEntries = CustomGUIs.DataGrid(paramEntries, paramPropertyNames, false);

        if (selectedParamEntries == null || selectedParamEntries.Count == 0)
        {
            TaskDialog.Show("Error", "No parameter selected.");
            return Result.Cancelled;
        }

        string selectedParamName = selectedParamEntries.First()["Name"].ToString();

        // Get the selected FamilyParameter
        FamilyParameter selectedParameter = familyParameters.FirstOrDefault(p => p.Definition.Name == selectedParamName);

        if (selectedParameter == null)
        {
            TaskDialog.Show("Error", "Selected parameter not found.");
            return Result.Cancelled;
        }

        // Prompt user to input a value
        string inputValue = PromptForValue(selectedParamName);

        if (inputValue == null)
        {
            TaskDialog.Show("Error", "No value entered.");
            return Result.Cancelled;
        }

        // Start transaction to set parameter values
        using (Transaction trans = new Transaction(doc, "Set Parameter Value"))
        {
            trans.Start();

            foreach (var familyType in selectedFamilyTypes)
            {
                try
                {
                    familyManager.CurrentType = familyType;

                    switch (selectedParameter.StorageType)
                    {
                        case StorageType.Double:
                            if (double.TryParse(inputValue, out double doubleValue))
                            {
                                ForgeTypeId unitTypeId = selectedParameter.GetUnitTypeId();
                                double internalValue = doubleValue;

                                if (unitTypeId != null)
                                {
                                    // Convert to internal units
                                    internalValue = UnitUtils.ConvertToInternalUnits(doubleValue, unitTypeId);
                                }

                                familyManager.Set(selectedParameter, internalValue);
                            }
                            else
                            {
                                TaskDialog.Show("Error", $"Invalid value '{inputValue}' for parameter '{selectedParamName}'.");
                            }
                            break;
                        case StorageType.Integer:
                            if (int.TryParse(inputValue, out int intValue))
                            {
                                familyManager.Set(selectedParameter, intValue);
                            }
                            else
                            {
                                TaskDialog.Show("Error", $"Invalid value '{inputValue}' for parameter '{selectedParamName}'.");
                            }
                            break;
                        case StorageType.String:
                            familyManager.Set(selectedParameter, inputValue);
                            break;
                        case StorageType.ElementId:
                            // Handle ElementId if needed
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", $"Failed to set value for parameter '{selectedParamName}' in family type '{familyType.Name}': {ex.Message}");
                }
            }

            trans.Commit();
        }

        TaskDialog.Show("Success", "Parameter values have been updated.");
        return Result.Succeeded;
    }

    private string PromptForValue(string paramName)
    {
        using (System.Windows.Forms.Form form = new System.Windows.Forms.Form())
        {
            form.Text = "Enter Parameter Value";
            System.Windows.Forms.Label label = new System.Windows.Forms.Label() { Left = 10, Top = 20, Text = $"Value for '{paramName}':" };
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox() { Left = 10, Top = 50, Width = 260 };
            System.Windows.Forms.Button buttonOk = new System.Windows.Forms.Button() { Text = "OK", Left = 110, Width = 80, Top = 80 };
            System.Windows.Forms.Button buttonCancel = new System.Windows.Forms.Button() { Text = "Cancel", Left = 190, Width = 80, Top = 80 };

            buttonOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;

            form.Controls.Add(label);
            form.Controls.Add(textBox);
            form.Controls.Add(buttonOk);
            form.Controls.Add(buttonCancel);
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;
            form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            form.ClientSize = new System.Drawing.Size(280, 120);

            System.Windows.Forms.DialogResult result = form.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                return textBox.Text;
            }
        }
        return null;
    }
}
