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
public class SetParametersOfFamilyTypes : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Step 1: Display list of categories that support parameters (multi-select)
        var categories = doc.Settings.Categories
            .Cast<Category>()
            .Where(c => c.CategoryType == CategoryType.Model && c.AllowsBoundParameters)
            .OrderBy(c => c.Name)
            .ToList();

        var categoryEntries = categories.Select(c => new CategoryEntry { Name = c.Name, Category = c }).ToList();
        var chosenCategoryEntries = CustomGUIs.DataGrid(categoryEntries, new List<string> { "Name" }, null, "Select One or More Categories");
        
        if (!chosenCategoryEntries.Any()) return Result.Cancelled;

        // Map known system family categories to their corresponding ElementType classes
        var categoryToClass = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase)
        {
            { "Walls", typeof(WallType) },
            { "Floors", typeof(FloorType) },
            { "Roofs", typeof(RoofType) },
            { "Ceilings", typeof(CeilingType) }
            // Add more system families as needed...
        };

        // Step 2: For each chosen category, gather family types or materials into a combined list
        var allChosenElements = new List<Element>();

        foreach (var catEntry in chosenCategoryEntries)
        {
            Category chosenCategory = catEntry.Category;
            bool isMaterialCategory = chosenCategory.Name.Equals("Materials", StringComparison.OrdinalIgnoreCase);

            if (isMaterialCategory)
            {
                // Handle Materials
                var materialCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Material))
                    .Cast<Material>()
                    .OrderBy(m => m.Name)
                    .ToList();

                // Add all materials to combined list
                allChosenElements.AddRange(materialCollector);
            }
            else if (categoryToClass.TryGetValue(chosenCategory.Name, out Type systemFamilyType))
            {
                // System Family Category: use the corresponding system family type class
                var systemFamilyCollector = new FilteredElementCollector(doc)
                    .OfClass(systemFamilyType)
                    .WhereElementIsElementType()
                    .ToList();

                allChosenElements.AddRange(systemFamilyCollector);
            }
            else
            {
                // Loadable families with FamilySymbol
                var collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType();

                var familySymbols = collector
                    .Cast<FamilySymbol>()
                    .Where(fs => fs.Category != null && fs.Category.Id.IntegerValue == chosenCategory.Id.IntegerValue)
                    .ToList();

                allChosenElements.AddRange(familySymbols);
            }
        }

        allChosenElements = allChosenElements.Distinct().ToList();

        if (!allChosenElements.Any())
        {
            TaskDialog.Show("No Types", "No applicable types or materials found for the selected categories.");
            return Result.Cancelled;
        }

        // Sort elements by their display name
        var elementEntries = allChosenElements
            .Select(elem => new GenericElementEntry
            {
                ElementName = GetElementDisplayName(elem),
                Element = elem
            })
            .OrderBy(e => e.ElementName)
            .ToList();

        // Step 3: Show combined list of elements (family types or materials)
        var chosenElementEntries = CustomGUIs.DataGrid(elementEntries, new List<string> { "ElementName" }, null, "Select Family Types/Materials (You can select multiple)");
        var chosenElements = chosenElementEntries.Select(e => e.Element).ToList();
        if (!chosenElements.Any()) return Result.Cancelled;

        // Step 4: Show parameters of all chosen elements combined
        var allParameters = new List<ParameterEntryForMultiple>();
        foreach (var elem in chosenElements)
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
            TaskDialog.Show("No Editable Parameters", "These elements have no editable parameters.");
            return Result.Cancelled;
        }

        // Sort by ElementName so they're grouped by their originating element
        allParameters = allParameters.OrderBy(p => p.ElementName).ToList();

        // Step 5: Show parameters for selection
        var chosenParameterEntries = CustomGUIs.DataGrid(allParameters, new List<string> { "ElementName", "ParameterName", "CurrentValue" }, null, "Select Parameters (Multiple elements combined)");
        if (!chosenParameterEntries.Any()) return Result.Cancelled;

        // Step 6: Ask user for new value
        string newValue = ShowInputDialog("Set Parameter Value", "Enter new value for the selected parameters:");
        if (newValue == null) return Result.Cancelled; // User cancelled

        // Step 7: Set the parameter's value for each selected parameter/element
        using (Transaction tx = new Transaction(doc, "Set Parameters"))
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

        return Result.Succeeded;
    }

    private string GetElementDisplayName(Element elem)
    {
        if (elem is FamilySymbol fs)
        {
            return $"{fs.Family.Name} : {fs.Name}";
        }
        else
        {
            // Try using the Element's Name property or the ALL_MODEL_TYPE_NAME parameter
            var nameParam = elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME);
            if (nameParam != null && !string.IsNullOrEmpty(nameParam.AsString()))
                return nameParam.AsString();

            // If that fails, fallback to elem.Name
            return elem.Name;
        }
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
        form.ClientSize = new System.Drawing.Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
        form.AcceptButton = buttonOk;
        form.CancelButton = buttonCancel;

        DialogResult dialogResult = form.ShowDialog();
        return dialogResult == DialogResult.OK ? textBox.Text : null;
    }

    public class CategoryEntry
    {
        public string Name { get; set; }
        public Category Category { get; set; }
    }

    public class FamilyTypeEntry
    {
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public Element Element { get; set; }
    }

    public class GenericElementEntry
    {
        public string ElementName { get; set; }
        public Element Element { get; set; }
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
