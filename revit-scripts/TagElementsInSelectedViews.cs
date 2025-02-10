using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Globalization;

[Transaction(TransactionMode.Manual)]
public class TagElementsInSelectedViews : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the selected elements
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
        if (selectedIds.Count == 0)
        {
            return Result.Failed;
        }

        List<Autodesk.Revit.DB.View> selectedViews = new List<Autodesk.Revit.DB.View>();
        foreach (ElementId id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element is Viewport viewport)
            {
                // Get the view associated with the viewport
                ElementId viewId = viewport.ViewId;
                Autodesk.Revit.DB.View view = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                if (view != null && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
                {
                    selectedViews.Add(view);
                }
            }
            else if (element is Autodesk.Revit.DB.View view && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
            {
                selectedViews.Add(view);
            }
        }

        if (selectedViews.Count == 0)
        {
            return Result.Failed;
        }

        // Gather all family types present in the selected views
        Dictionary<ElementId, (string CategoryName, string FamilyName, string TypeName)> familyTypes = new Dictionary<ElementId, (string, string, string)>();
        foreach (Autodesk.Revit.DB.View view in selectedViews)
        {
            var elementsToTag = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && e.CanBeTagged(view))
                .ToList();

            foreach (Element element in elementsToTag)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId && !familyTypes.ContainsKey(typeId))
                {
                    ElementType elementType = doc.GetElement(typeId) as ElementType;
                    if (elementType != null)
                    {
                        familyTypes[typeId] = (element.Category.Name, elementType.FamilyName, elementType.Name);
                    }
                }
            }
        }

        // Prompt user to choose family types to tag
        List<Dictionary<string, object>> entries = familyTypes
            .Select(kv => new Dictionary<string, object> { { "Category", kv.Value.CategoryName }, { "Family", kv.Value.FamilyName }, { "Type", kv.Value.TypeName } })
            .ToList();
        List<string> propertyNames = new List<string> { "Category", "Family", "Type" };
        List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);
        if (selectedEntries.Count == 0)
        {
            System.Windows.Forms.MessageBox.Show("No family types selected to tag.", "Error");
            return Result.Failed;
        }

        HashSet<ElementId> selectedTypeIds = new HashSet<ElementId>(
            selectedEntries.Select(e =>
                familyTypes.FirstOrDefault(kv => kv.Value.CategoryName == e["Category"].ToString() &&
                                                  kv.Value.FamilyName == e["Family"].ToString() &&
                                                  kv.Value.TypeName == e["Type"].ToString()).Key));

        // Show a new dialog to allow the user to specify the offset factors.
        using (OffsetInputDialog offsetDialog = new OffsetInputDialog())
        {
            if (offsetDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                System.Windows.Forms.MessageBox.Show("Operation cancelled by the user.", "Cancelled");
                return Result.Failed;
            }
            // Retrieve user-defined offset factors.
            double userOffsetFactorX = offsetDialog.OffsetX;
            double userOffsetFactorY = offsetDialog.OffsetY;

            using (Transaction trans = new Transaction(doc, "Place Tags"))
            {
                trans.Start();
                try
                {
                    int tagNumber = 1;

                    foreach (Autodesk.Revit.DB.View view in selectedViews)
                    {
                        var elementsToTag = new FilteredElementCollector(doc, view.Id)
                            .WhereElementIsNotElementType()
                            .Where(e => e.Category != null && e.CanBeTagged(view) && selectedTypeIds.Contains(e.GetTypeId()))
                            .ToList();

                        foreach (Element element in elementsToTag)
                        {
                            LocationPoint location = element.Location as LocationPoint;
                            XYZ originalPosition = location?.Point;
                            BoundingBoxXYZ boundingBox = element.get_BoundingBox(null);

                            if (boundingBox != null && originalPosition != null)
                            {
                                // Get the current view scale
                                Autodesk.Revit.DB.View activeView = view;
                                int viewScale = activeView.Scale;

                                // Calculate the offsets using the user-defined factors:
                                double offsetX = userOffsetFactorX * viewScale;
                                double offsetY = userOffsetFactorY * viewScale;

                                // Calculate the tag position with offset
                                XYZ minPoint = boundingBox.Min;
                                XYZ tagPosition = new XYZ(originalPosition.X + offsetX, minPoint.Y + offsetY, originalPosition.Z);

                                IndependentTag newTag = IndependentTag.Create(doc, activeView.Id, new Reference(element), false, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, tagPosition);
                                Parameter typeMarkParam = newTag.LookupParameter("Type Mark");

                                if (typeMarkParam != null && !typeMarkParam.IsReadOnly)
                                {
                                    typeMarkParam.Set(tagNumber.ToString());
                                }

                                tagNumber++;
                            }
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    trans.RollBack();
                    return Result.Failed;
                }
            }
        }

        return Result.Succeeded;
    }
}

// Extension method to check if an element can be tagged in a specific view
public static class ElementExtensions
{
    public static bool CanBeTagged(this Element element, Autodesk.Revit.DB.View view)
    {
        // Check if the element can be tagged in the given view.
        // You can expand this logic as needed for your specific requirements.
        return element.Category != null &&
               element.Category.HasMaterialQuantities &&
               element.Category.CanAddSubcategory &&
               !element.Category.IsTagCategory;
    }
}

// A new dialog form for the offset input.
public class OffsetInputDialog : System.Windows.Forms.Form
{
    private System.Windows.Forms.Label labelX;
    private System.Windows.Forms.Label labelY;
    private System.Windows.Forms.TextBox textBoxX;
    private System.Windows.Forms.TextBox textBoxY;
    private System.Windows.Forms.Button okButton;
    private System.Windows.Forms.Button cancelButton;

    // Properties to retrieve the user-entered offset factors.
    public double OffsetX { get; private set; }
    public double OffsetY { get; private set; }

    public OffsetInputDialog()
    {
        InitializeComponent();
        LoadLastOffsets();
    }

    private void InitializeComponent()
    {
        this.labelX = new System.Windows.Forms.Label();
        this.labelY = new System.Windows.Forms.Label();
        this.textBoxX = new System.Windows.Forms.TextBox();
        this.textBoxY = new System.Windows.Forms.TextBox();
        this.okButton = new System.Windows.Forms.Button();
        this.cancelButton = new System.Windows.Forms.Button();

        // labelX
        this.labelX.AutoSize = true;
        this.labelX.Location = new System.Drawing.Point(12, 15);
        this.labelX.Name = "labelX";
        this.labelX.Size = new System.Drawing.Size(70, 13);
        this.labelX.TabIndex = 0;
        this.labelX.Text = "Offset X Factor:";

        // textBoxX
        this.textBoxX.Location = new System.Drawing.Point(100, 12);
        this.textBoxX.Name = "textBoxX";
        this.textBoxX.Size = new System.Drawing.Size(100, 20);
        this.textBoxX.TabIndex = 1;
        this.textBoxX.Text = "0.009"; // default value

        // labelY
        this.labelY.AutoSize = true;
        this.labelY.Location = new System.Drawing.Point(12, 45);
        this.labelY.Name = "labelY";
        this.labelY.Size = new System.Drawing.Size(70, 13);
        this.labelY.TabIndex = 2;
        this.labelY.Text = "Offset Y Factor:";

        // textBoxY
        this.textBoxY.Location = new System.Drawing.Point(100, 42);
        this.textBoxY.Name = "textBoxY";
        this.textBoxY.Size = new System.Drawing.Size(100, 20);
        this.textBoxY.TabIndex = 3;
        this.textBoxY.Text = "0.009"; // default value

        // okButton
        this.okButton.Location = new System.Drawing.Point(44, 80);
        this.okButton.Name = "okButton";
        this.okButton.Size = new System.Drawing.Size(75, 23);
        this.okButton.TabIndex = 4;
        this.okButton.Text = "OK";
        this.okButton.UseVisualStyleBackColor = true;
        this.okButton.Click += new System.EventHandler(this.OkButton_Click);

        // cancelButton
        this.cancelButton.Location = new System.Drawing.Point(125, 80);
        this.cancelButton.Name = "cancelButton";
        this.cancelButton.Size = new System.Drawing.Size(75, 23);
        this.cancelButton.TabIndex = 5;
        this.cancelButton.Text = "Cancel";
        this.cancelButton.UseVisualStyleBackColor = true;
        this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);

        // Set Accept and Cancel buttons for key handling.
        this.AcceptButton = this.okButton;
        this.CancelButton = this.cancelButton;

        // OffsetInputDialog form settings
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(220, 120);
        this.Controls.Add(this.labelX);
        this.Controls.Add(this.textBoxX);
        this.Controls.Add(this.labelY);
        this.Controls.Add(this.textBoxY);
        this.Controls.Add(this.okButton);
        this.Controls.Add(this.cancelButton);
        this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
        this.Text = "Enter Offset Factors";
    }

    private void OkButton_Click(object sender, System.EventArgs e)
    {
        if (double.TryParse(this.textBoxX.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double offsetX) &&
            double.TryParse(this.textBoxY.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double offsetY))
        {
            this.OffsetX = offsetX;
            this.OffsetY = offsetY;
            SaveLastOffsets(offsetX, offsetY);
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
        else
        {
            System.Windows.Forms.MessageBox.Show("Please enter valid numeric values for both offsets.",
                                                  "Invalid Input",
                                                  System.Windows.Forms.MessageBoxButtons.OK,
                                                  System.Windows.Forms.MessageBoxIcon.Error);
        }
    }

    private void CancelButton_Click(object sender, System.EventArgs e)
    {
        this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        this.Close();
    }

    /// <summary>
    /// Loads previously saved offset values (if available) from
    /// %appdata%\revit-scripts\TagElementsInSelectedViews-last-offsets.
    /// </summary>
    private void LoadLastOffsets()
    {
        try
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath = Path.Combine(appData, "revit-scripts");
            string filePath = Path.Combine(folderPath, "TagElementsInSelectedViews-last-offsets");
            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                if (lines.Length >= 2)
                {
                    if (double.TryParse(lines[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double lastOffsetX))
                    {
                        this.textBoxX.Text = lastOffsetX.ToString(CultureInfo.InvariantCulture);
                    }
                    if (double.TryParse(lines[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double lastOffsetY))
                    {
                        this.textBoxY.Text = lastOffsetY.ToString(CultureInfo.InvariantCulture);
                    }
                }
            }
        }
        catch
        {
            // If loading fails, leave the default values.
        }
    }

    /// <summary>
    /// Saves the provided offset values to
    /// %appdata%\revit-scripts\TagElementsInSelectedViews-last-offsets.
    /// </summary>
    private void SaveLastOffsets(double offsetX, double offsetY)
    {
        try
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath = Path.Combine(appData, "revit-scripts");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            string filePath = Path.Combine(folderPath, "TagElementsInSelectedViews-last-offsets");
            string[] lines = new string[2];
            lines[0] = offsetX.ToString(CultureInfo.InvariantCulture);
            lines[1] = offsetY.ToString(CultureInfo.InvariantCulture);
            File.WriteAllLines(filePath, lines);
        }
        catch
        {
            // Optionally handle errors here.
        }
    }
}
