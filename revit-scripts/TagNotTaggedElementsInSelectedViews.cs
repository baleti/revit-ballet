using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Globalization;

[Transaction(TransactionMode.Manual)]
public class TagNotTaggedElementsInSelectedViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the selected views (or viewports that reference a view)
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
        if (selectedIds.Count == 0)
            return Result.Failed;

        List<View> selectedViews = new List<View>();
        foreach (ElementId id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element is Viewport viewport)
            {
                ElementId viewId = viewport.ViewId;
                View view = doc.GetElement(viewId) as View;
                if (view != null && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
                    selectedViews.Add(view);
            }
            else if (element is View view && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
            {
                selectedViews.Add(view);
            }
        }
        if (selectedViews.Count == 0)
            return Result.Failed;

        // Build a dictionary of family types from elements in the selected views.
        // (Note: No crop region or heavy geometry tests are done here.)
        Dictionary<ElementId, (string CategoryName, string FamilyName, string TypeName)> familyTypes =
            new Dictionary<ElementId, (string, string, string)>();

        foreach (View view in selectedViews)
        {
            var elementsInView = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && e.CanBeTagged(view))
                .ToList();

            foreach (Element element in elementsInView)
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
        if (familyTypes.Count == 0)
        {
            TaskDialog.Show("Warning", "No taggable elements found in the selected views.");
            return Result.Failed;
        }

        // Compute counts for each family type across all views.
        Dictionary<ElementId, HashSet<ElementId>> allElementsByType = new Dictionary<ElementId, HashSet<ElementId>>();
        Dictionary<ElementId, HashSet<ElementId>> taggedElementsByType = new Dictionary<ElementId, HashSet<ElementId>>();

        foreach (View view in selectedViews)
        {
            // Get existing tags in this view.
            var existingTags = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(IndependentTag))
                .Cast<IndependentTag>()
                .ToList();
            HashSet<ElementId> alreadyTaggedIds = new HashSet<ElementId>(
                existingTags.SelectMany(tag => tag.GetTaggedElementIds()
                    .Where(linkId => linkId.LinkInstanceId == ElementId.InvalidElementId)
                    .Select(linkId => linkId.HostElementId)));

            var elementsInView = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && e.CanBeTagged(view) && familyTypes.ContainsKey(e.GetTypeId()))
                .ToList();

            foreach (Element element in elementsInView)
            {
                ElementId typeId = element.GetTypeId();
                if (!allElementsByType.ContainsKey(typeId))
                    allElementsByType[typeId] = new HashSet<ElementId>();
                allElementsByType[typeId].Add(element.Id);

                if (alreadyTaggedIds.Contains(element.Id))
                {
                    if (!taggedElementsByType.ContainsKey(typeId))
                        taggedElementsByType[typeId] = new HashSet<ElementId>();
                    taggedElementsByType[typeId].Add(element.Id);
                }
            }
        }

        // Build the DataGrid entries.
        List<Dictionary<string, object>> entries = familyTypes
            .Select(kv =>
            {
                int totalCount = allElementsByType.ContainsKey(kv.Key) ? allElementsByType[kv.Key].Count : 0;
                int taggedCount = taggedElementsByType.ContainsKey(kv.Key) ? taggedElementsByType[kv.Key].Count : 0;
                int untaggedCount = totalCount - taggedCount;
                return new Dictionary<string, object>
                {
                    { "Category", kv.Value.CategoryName },
                    { "Family", kv.Value.FamilyName },
                    { "Type", kv.Value.TypeName },
                    { "Tagged (Estimate)", taggedCount },
                    { "Untagged (Estimate)", untaggedCount }
                };
            })
            .ToList();

        if (entries.Count == 0)
        {
            TaskDialog.Show("Warning", "No family types available to tag.");
            return Result.Failed;
        }

        List<string> propertyNames = new List<string> { "Category", "Family", "Type", "Tagged (Estimate)", "Untagged (Estimate)" };
        // Display the DataGrid without any heavy crop region processing.
        List<Dictionary<string, object>> selectedEntries =
            CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);
        if (selectedEntries == null || selectedEntries.Count == 0)
        {
            TaskDialog.Show("Information", "No family types selected to tag.");
            return Result.Failed;
        }

        // Determine the selected family type IDs.
        HashSet<ElementId> selectedTypeIds = new HashSet<ElementId>(
            selectedEntries.Select(e =>
                familyTypes.FirstOrDefault(kv => kv.Value.CategoryName == e["Category"].ToString() &&
                                                  kv.Value.FamilyName == e["Family"].ToString() &&
                                                  kv.Value.TypeName == e["Type"].ToString()).Key));

        // Now prompt the user for offset factors.
        using (OffsetInputDialog offsetDialog = new OffsetInputDialog())
        {
            if (offsetDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                System.Windows.Forms.MessageBox.Show("Operation cancelled by the user.", "Cancelled");
                return Result.Failed;
            }
            double userOffsetFactorX = offsetDialog.OffsetX;
            double userOffsetFactorY = offsetDialog.OffsetY;

            // Begin transaction to create tags. Now, for each element the crop region is retrieved
            // and processed before placing a tag.
            using (Transaction trans = new Transaction(doc, "Place Tags"))
            {
                trans.Start();
                try
                {
                    int tagNumber = 1;
                    foreach (View view in selectedViews)
                    {
                        // Get existing tags in the view.
                        var existingTags = new FilteredElementCollector(doc, view.Id)
                            .OfClass(typeof(IndependentTag))
                            .Cast<IndependentTag>()
                            .ToList();
                        HashSet<ElementId> alreadyTaggedIds = new HashSet<ElementId>(
                            existingTags.SelectMany(tag => tag.GetTaggedElementIds()
                                .Where(linkId => linkId.LinkInstanceId == ElementId.InvalidElementId)
                                .Select(linkId => linkId.HostElementId)));

                        // Retrieve (and cache) the crop region curves once per view.
                        IList<CurveLoop> cropLoops = null;
                        if (view.CropBoxActive && view.CropBox != null)
                        {
                            ViewCropRegionShapeManager cropManager = view.GetCropRegionShapeManager();
                            cropLoops = cropManager.GetCropShape();
                        }

                        // Collect taggable elements that match the selected family types and are not already tagged.
                        var elementsToTag = new FilteredElementCollector(doc, view.Id)
                            .WhereElementIsNotElementType()
                            .Where(e => e.Category != null && e.CanBeTagged(view) && selectedTypeIds.Contains(e.GetTypeId()))
                            .Where(e => !alreadyTaggedIds.Contains(e.Id))
                            .ToList();

                        foreach (Element element in elementsToTag)
                        {
                            // Process only elements with a LocationPoint.
                            LocationPoint location = element.Location as LocationPoint;
                            if (location == null)
                                continue;

                            XYZ elementPoint = location.Point;

                            // If the view has an active crop region, perform the point-in-polygon test.
                            if (cropLoops != null && cropLoops.Count > 0)
                            {
                                // Transform the element’s point from model space into crop-region space.
                                Transform inverseTransform = view.CropBox.Transform.Inverse;
                                XYZ elementInCropSpace = inverseTransform.OfPoint(elementPoint);
                                UV elementUV = new UV(elementInCropSpace.X, elementInCropSpace.Y);

                                // Assume the first loop is the outer boundary.
                                List<UV> outerPolygon = GetVerticesFromCurveLoop(cropLoops[0]);
                                bool isInside = IsPointInsidePolygon(elementUV, outerPolygon);

                                // If additional loops exist, assume they are holes.
                                if (isInside && cropLoops.Count > 1)
                                {
                                    for (int i = 1; i < cropLoops.Count; i++)
                                    {
                                        List<UV> holePolygon = GetVerticesFromCurveLoop(cropLoops[i]);
                                        if (IsPointInsidePolygon(elementUV, holePolygon))
                                        {
                                            isInside = false;
                                            break;
                                        }
                                    }
                                }
                                if (!isInside)
                                {
                                    // Skip this element if it lies outside the crop region.
                                    continue;
                                }
                            }

                            // Determine tag placement using the element’s bounding box and user-defined offsets.
                            BoundingBoxXYZ boundingBox = element.get_BoundingBox(null);
                            if (boundingBox != null)
                            {
                                int viewScale = view.Scale;
                                double offsetX = userOffsetFactorX * viewScale;
                                double offsetY = userOffsetFactorY * viewScale;

                                XYZ centerPoint = (boundingBox.Min + boundingBox.Max) / 2.0;
                                XYZ tagPosition = new XYZ(elementPoint.X + offsetX, centerPoint.Y + offsetY, elementPoint.Z);

                                IndependentTag newTag = IndependentTag.Create(
                                    doc,
                                    view.Id,
                                    new Reference(element),
                                    false,
                                    TagMode.TM_ADDBY_CATEGORY,
                                    TagOrientation.Horizontal,
                                    tagPosition);
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

    /// <summary>
    /// Helper method: Extracts a list of UV points from a CurveLoop.
    /// </summary>
    private static List<UV> GetVerticesFromCurveLoop(CurveLoop loop)
    {
        List<UV> vertices = new List<UV>();
        foreach (Curve curve in loop)
        {
            // For a closed loop, the first point of each curve defines the outline.
            XYZ p = curve.GetEndPoint(0);
            vertices.Add(new UV(p.X, p.Y));
        }
        return vertices;
    }

    /// <summary>
    /// Helper method: Determines if a given 2D point (UV) lies inside a polygon defined by a list of UV points.
    /// Uses the ray-casting algorithm.
    /// </summary>
    private static bool IsPointInsidePolygon(UV point, List<UV> polygon)
    {
        bool inside = false;
        int count = polygon.Count;
        for (int i = 0, j = count - 1; i < count; j = i++)
        {
            if (((polygon[i].V > point.V) != (polygon[j].V > point.V)) &&
                (point.U < (polygon[j].U - polygon[i].U) * (point.V - polygon[i].V) / (polygon[j].V - polygon[i].V) + polygon[i].U))
            {
                inside = !inside;
            }
        }
        return inside;
    }
}

// Extension method for checking whether an element can be tagged in a view.
public static class ElementExtensions
{
    public static bool CanBeTagged(this Element element, View view)
    {
        return element.Category != null &&
               element.Category.HasMaterialQuantities &&
               element.Category.CanAddSubcategory &&
               !element.Category.IsTagCategory;
    }
}

// A dialog form for entering offset factors.
public class OffsetInputDialog : System.Windows.Forms.Form
{
    private System.Windows.Forms.Label labelX;
    private System.Windows.Forms.Label labelY;
    private System.Windows.Forms.TextBox textBoxX;
    private System.Windows.Forms.TextBox textBoxY;
    private System.Windows.Forms.Button okButton;
    private System.Windows.Forms.Button cancelButton;

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

        this.AcceptButton = this.okButton;
        this.CancelButton = this.cancelButton;
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

    private void OkButton_Click(object sender, EventArgs e)
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

    private void CancelButton_Click(object sender, EventArgs e)
    {
        this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        this.Close();
    }

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
            // Ignore errors if loading fails.
        }
    }

    private void SaveLastOffsets(double offsetX, double offsetY)
    {
        try
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath = Path.Combine(appData, "revit-scripts");
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);
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
