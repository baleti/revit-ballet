using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Globalization;
using WinForms = System.Windows.Forms;  // Alias for Windows Forms

// Enum for tag anchor choices.
public enum TagAnchor
{
    Center,
    Top,
    Bottom,
    Left,
    Right
}

[Transaction(TransactionMode.Manual)]
public class TagNotTaggedElementsInSelectedViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the selected views (or viewports that reference a view)
        ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
        if (selectedIds.Count == 0)
            return Result.Failed;

        List<Autodesk.Revit.DB.View> selectedViews = new List<Autodesk.Revit.DB.View>();
        foreach (ElementId id in selectedIds)
        {
            Element element = doc.GetElement(id);
            if (element is Viewport viewport)
            {
                ElementId viewId = viewport.ViewId;
                Autodesk.Revit.DB.View view = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                if (view != null && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
                    selectedViews.Add(view);
            }
            else if (element is Autodesk.Revit.DB.View view && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
            {
                selectedViews.Add(view);
            }
        }
        if (selectedViews.Count == 0)
            return Result.Failed;

        // Build a dictionary of family types from elements in the selected views.
        Dictionary<ElementId, (string CategoryName, string FamilyName, string TypeName)> familyTypes =
            new Dictionary<ElementId, (string, string, string)>();

        foreach (Autodesk.Revit.DB.View view in selectedViews)
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

        // Compute counts for each family type (fast estimate).
        Dictionary<ElementId, int> totalCounts = new Dictionary<ElementId, int>();
        Dictionary<ElementId, int> taggedCounts = new Dictionary<ElementId, int>();

        foreach (Autodesk.Revit.DB.View view in selectedViews)
        {
            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && e.CanBeTagged(view) && familyTypes.ContainsKey(e.GetTypeId()));

            foreach (var group in collector.GroupBy(e => e.GetTypeId()))
            {
                if (!totalCounts.ContainsKey(group.Key))
                    totalCounts[group.Key] = 0;
                totalCounts[group.Key] += group.Count();
            }

            var taggedIds = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(IndependentTag))
                .Cast<IndependentTag>()
                .SelectMany(tag => tag.GetTaggedElementIds()
                                      .Where(linkId => linkId.LinkInstanceId == ElementId.InvalidElementId)
                                      .Select(linkId => linkId.HostElementId))
                .ToHashSet();

            foreach (var group in collector.GroupBy(e => e.GetTypeId()))
            {
                int countTagged = group.Count(e => taggedIds.Contains(e.Id));
                if (!taggedCounts.ContainsKey(group.Key))
                    taggedCounts[group.Key] = 0;
                taggedCounts[group.Key] += countTagged;
            }
        }

        // Build the DataGrid entries.
        List<Dictionary<string, object>> entries = familyTypes
            .Select(kv =>
            {
                int totalCount = totalCounts.ContainsKey(kv.Key) ? totalCounts[kv.Key] : 0;
                int taggedCount = taggedCounts.ContainsKey(kv.Key) ? taggedCounts[kv.Key] : 0;
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
        // Display the DataGrid (assumed to be a custom UI helper).
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

        // Prompt the user for offset factors, orientation, and tag anchor.
        using (OffsetInputDialog offsetDialog = new OffsetInputDialog())
        {
            if (offsetDialog.ShowDialog() != WinForms.DialogResult.OK)
            {
                WinForms.MessageBox.Show("Operation cancelled by the user.", "Cancelled");
                return Result.Failed;
            }
            double userOffsetFactorX = offsetDialog.OffsetX;
            double userOffsetFactorY = offsetDialog.OffsetY;
            bool orientToObject = offsetDialog.OrientToObject;
            TagAnchor tagAnchorOption = offsetDialog.SelectedTagAnchor;

            // Conversion factor from millimeters to feet.
            const double mmToFeet = 1.0 / 304.8;

            // Begin transaction to create tags.
            using (Transaction trans = new Transaction(doc, "Place Tags"))
            {
                trans.Start();
                try
                {
                    int tagNumber = 1;
                    foreach (Autodesk.Revit.DB.View view in selectedViews)
                    {
                        var existingTags = new FilteredElementCollector(doc, view.Id)
                            .OfClass(typeof(IndependentTag))
                            .Cast<IndependentTag>()
                            .ToList();
                        HashSet<ElementId> alreadyTaggedIds = new HashSet<ElementId>(
                            existingTags.SelectMany(tag => tag.GetTaggedElementIds()
                                .Where(linkId => linkId.LinkInstanceId == ElementId.InvalidElementId)
                                .Select(linkId => linkId.HostElementId)));

                        IList<CurveLoop> cropLoops = null;
                        if (view.CropBoxActive && view.CropBox != null)
                        {
                            ViewCropRegionShapeManager cropManager = view.GetCropRegionShapeManager();
                            cropLoops = cropManager.GetCropShape();
                        }

                        var elementsToTag = new FilteredElementCollector(doc, view.Id)
                            .WhereElementIsNotElementType()
                            .Where(e => e.Category != null && e.CanBeTagged(view) && selectedTypeIds.Contains(e.GetTypeId()))
                            .Where(e => !alreadyTaggedIds.Contains(e.Id))
                            .ToList();

                        foreach (Element element in elementsToTag)
                        {
                            LocationPoint location = element.Location as LocationPoint;
                            if (location == null)
                                continue;

                            BoundingBoxXYZ bbox = element.get_BoundingBox(null);
                            if (bbox == null)
                                continue;

                            // --- Crop Region Bounding Box Check ---
                            if (cropLoops != null && cropLoops.Count > 0)
                            {
                                Transform invTransform = view.CropBox.Transform.Inverse;
                                List<UV> outerPolygon = GetVerticesFromCurveLoop(cropLoops[0]);

                                XYZ pMin = bbox.Min;
                                XYZ pMax = bbox.Max;

                                XYZ corner1 = ProjectPointToCropPlane(new XYZ(pMin.X, pMin.Y, pMin.Z), view.CropBox.Transform);
                                XYZ corner2 = ProjectPointToCropPlane(new XYZ(pMax.X, pMin.Y, pMin.Z), view.CropBox.Transform);
                                XYZ corner3 = ProjectPointToCropPlane(new XYZ(pMax.X, pMax.Y, pMax.Z), view.CropBox.Transform);
                                XYZ corner4 = ProjectPointToCropPlane(new XYZ(pMin.X, pMax.Y, pMax.Z), view.CropBox.Transform);

                                XYZ cp1 = invTransform.OfPoint(corner1);
                                XYZ cp2 = invTransform.OfPoint(corner2);
                                XYZ cp3 = invTransform.OfPoint(corner3);
                                XYZ cp4 = invTransform.OfPoint(corner4);

                                UV uv1 = new UV(cp1.X, cp1.Y);
                                UV uv2 = new UV(cp2.X, cp2.Y);
                                UV uv3 = new UV(cp3.X, cp3.Y);
                                UV uv4 = new UV(cp4.X, cp4.Y);

                                bool bboxInside = true;
                                foreach (UV uv in new List<UV> { uv1, uv2, uv3, uv4 })
                                {
                                    if (!IsPointInsidePolygon(uv, outerPolygon))
                                    {
                                        bboxInside = false;
                                        break;
                                    }
                                    if (cropLoops.Count > 1)
                                    {
                                        for (int i = 1; i < cropLoops.Count; i++)
                                        {
                                            List<UV> holePolygon = GetVerticesFromCurveLoop(cropLoops[i]);
                                            if (IsPointInsidePolygon(uv, holePolygon))
                                            {
                                                bboxInside = false;
                                                break;
                                            }
                                        }
                                        if (!bboxInside)
                                            break;
                                    }
                                }
                                if (!bboxInside)
                                    continue;
                            }
                            // --- End Crop Region Check ---

                            XYZ tagPosition = null;
                            if (!orientToObject)
                            {
                                // Non-orienting branch – compute tag anchor based on the selected option.
                                XYZ anchorPoint;
                                if (tagAnchorOption == TagAnchor.Center)
                                {
                                    anchorPoint = (bbox.Min + bbox.Max) / 2.0;
                                }
                                else
                                {
                                    Transform cropTransform = view.CropBox.Transform;
                                    Transform invT = cropTransform.Inverse;
                                    XYZ projMin = ProjectPointToCropPlane(bbox.Min, cropTransform);
                                    XYZ projMax = ProjectPointToCropPlane(bbox.Max, cropTransform);
                                    XYZ bboxMinView = invT.OfPoint(projMin);
                                    XYZ bboxMaxView = invT.OfPoint(projMax);
                                    double xCenter = (bboxMinView.X + bboxMaxView.X) / 2.0;
                                    double yCenter = (bboxMinView.Y + bboxMaxView.Y) / 2.0;
                                    XYZ anchorView;
                                    switch (tagAnchorOption)
                                    {
                                        case TagAnchor.Top:
                                            anchorView = new XYZ(xCenter, bboxMaxView.Y, 0);
                                            break;
                                        case TagAnchor.Bottom:
                                            anchorView = new XYZ(xCenter, bboxMinView.Y, 0);
                                            break;
                                        case TagAnchor.Left:
                                            anchorView = new XYZ(bboxMinView.X, yCenter, 0);
                                            break;
                                        case TagAnchor.Right:
                                            anchorView = new XYZ(bboxMaxView.X, yCenter, 0);
                                            break;
                                        default:
                                            anchorView = new XYZ(xCenter, yCenter, 0);
                                            break;
                                    }
                                    XYZ anchorModel = cropTransform.OfPoint(anchorView);
                                    double z = (bbox.Min.Z + bbox.Max.Z) / 2.0;
                                    anchorPoint = new XYZ(anchorModel.X, anchorModel.Y, z);
                                }
                                double offsetX = userOffsetFactorX * mmToFeet;
                                double offsetY = userOffsetFactorY * mmToFeet;
                                tagPosition = new XYZ(anchorPoint.X + offsetX, anchorPoint.Y + offsetY, anchorPoint.Z);
                            }
                            else
                            {
                                // Orient-to-object branch remains unchanged.
                                double offsetX = userOffsetFactorX * mmToFeet;
                                double offsetY = userOffsetFactorY * mmToFeet;
                                XYZ basePoint = (bbox.Min + bbox.Max) / 2.0;
                                XYZ facing = XYZ.BasisY;
                                if (element is FamilyInstance fi)
                                {
                                    facing = fi.FacingOrientation;
                                    if (facing.IsAlmostEqualTo(XYZ.Zero))
                                        facing = XYZ.BasisY;
                                    else
                                        facing = facing.Normalize();
                                }
                                XYZ left = new XYZ(-facing.Y, facing.X, 0).Normalize();
                                tagPosition = basePoint + (facing * offsetX) + (left * offsetY);
                            }

                            if (tagPosition != null)
                            {
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
            XYZ p = curve.GetEndPoint(0);
            vertices.Add(new UV(p.X, p.Y));
        }
        return vertices;
    }

    /// <summary>
    /// Projects a 3D point onto the crop plane defined by the crop transform.
    /// </summary>
    private static XYZ ProjectPointToCropPlane(XYZ point, Transform cropTransform)
    {
        XYZ planeOrigin = cropTransform.Origin;
        XYZ normal = cropTransform.BasisZ;
        double distance = (point - planeOrigin).DotProduct(normal);
        return point - distance * normal;
    }

    /// <summary>
    /// Determines if a given 2D point (UV) lies inside a polygon defined by a list of UV points.
    /// Points on an edge (within a small tolerance) are considered inside.
    /// Uses the ray-casting algorithm.
    /// </summary>
    private static bool IsPointInsidePolygon(UV point, List<UV> polygon)
    {
        double tol = 1e-6;
        int count = polygon.Count;
        for (int i = 0; i < count; i++)
        {
            int j = (i + 1) % count;
            if (IsPointOnLineSegment(point, polygon[i], polygon[j], tol))
                return true;
        }
        bool inside = false;
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

    /// <summary>
    /// Determines if a point lies on the line segment between two points, within a tolerance.
    /// </summary>
    private static bool IsPointOnLineSegment(UV p, UV a, UV b, double tol = 1e-6)
    {
        double cross = Math.Abs((p.V - a.V) * (b.U - a.U) - (p.U - a.U) * (b.V - a.V));
        if (cross > tol)
            return false;
        double dot = (p.U - a.U) * (b.U - a.U) + (p.V - a.V) * (b.V - a.V);
        if (dot < 0)
            return false;
        double lenSq = (b.U - a.U) * (b.U - a.U) + (b.V - a.V) * (b.V - a.V);
        if (dot > lenSq)
            return false;
        return true;
    }
}

// Extension method for checking whether an element can be tagged in a view.
public static class ElementExtensions
{
    public static bool CanBeTagged(this Element element, Autodesk.Revit.DB.View view)
    {
        return element.Category != null &&
               element.Category.HasMaterialQuantities &&
               element.Category.CanAddSubcategory &&
               !element.Category.IsTagCategory;
    }
}

// A dialog form for entering offset factors, orientation, and tag anchor.
public class OffsetInputDialog : WinForms.Form
{
    private WinForms.Label labelX;
    private WinForms.Label labelY;
    private WinForms.TextBox textBoxX;
    private WinForms.TextBox textBoxY;
    private WinForms.CheckBox checkBoxOrient;
    private WinForms.Label labelTagAnchor;
    private WinForms.ComboBox comboBoxTagAnchor;
    private WinForms.Button okButton;
    private WinForms.Button cancelButton;

    public double OffsetX { get; private set; }
    public double OffsetY { get; private set; }
    public bool OrientToObject
    {
        get { return checkBoxOrient.Checked; }
    }
    public TagAnchor SelectedTagAnchor { get; private set; }

    public OffsetInputDialog()
    {
        InitializeComponent();
        LoadLastOffsets();
    }

    private void InitializeComponent()
    {
        this.labelX = new WinForms.Label();
        this.labelY = new WinForms.Label();
        this.textBoxX = new WinForms.TextBox();
        this.textBoxY = new WinForms.TextBox();
        this.checkBoxOrient = new WinForms.CheckBox();
        this.labelTagAnchor = new WinForms.Label();
        this.comboBoxTagAnchor = new WinForms.ComboBox();
        this.okButton = new WinForms.Button();
        this.cancelButton = new WinForms.Button();
        // 
        // labelX
        // 
        this.labelX.AutoSize = true;
        this.labelX.Location = new System.Drawing.Point(12, 15);
        this.labelX.Name = "labelX";
        this.labelX.Size = new System.Drawing.Size(70, 13);
        this.labelX.TabIndex = 0;
        this.labelX.Text = "Offset X (mm):";
        // 
        // textBoxX
        // 
        this.textBoxX.Location = new System.Drawing.Point(100, 12);
        this.textBoxX.Name = "textBoxX";
        this.textBoxX.Size = new System.Drawing.Size(100, 20);
        this.textBoxX.TabIndex = 1;
        this.textBoxX.Text = "10";
        // 
        // labelY
        // 
        this.labelY.AutoSize = true;
        this.labelY.Location = new System.Drawing.Point(12, 45);
        this.labelY.Name = "labelY";
        this.labelY.Size = new System.Drawing.Size(70, 13);
        this.labelY.TabIndex = 2;
        this.labelY.Text = "Offset Y (mm):";
        // 
        // textBoxY
        // 
        this.textBoxY.Location = new System.Drawing.Point(100, 42);
        this.textBoxY.Name = "textBoxY";
        this.textBoxY.Size = new System.Drawing.Size(100, 20);
        this.textBoxY.TabIndex = 3;
        this.textBoxY.Text = "10";
        // 
        // checkBoxOrient
        // 
        this.checkBoxOrient.AutoSize = true;
        this.checkBoxOrient.Location = new System.Drawing.Point(12, 75);
        this.checkBoxOrient.Name = "checkBoxOrient";
        this.checkBoxOrient.Size = new System.Drawing.Size(100, 17);
        this.checkBoxOrient.TabIndex = 4;
        this.checkBoxOrient.Text = "Orient to object";
        this.checkBoxOrient.UseVisualStyleBackColor = true;
        this.checkBoxOrient.Checked = false;
        // 
        // labelTagAnchor
        // 
        this.labelTagAnchor.AutoSize = true;
        this.labelTagAnchor.Location = new System.Drawing.Point(12, 105);
        this.labelTagAnchor.Name = "labelTagAnchor";
        this.labelTagAnchor.Size = new System.Drawing.Size(67, 13);
        this.labelTagAnchor.TabIndex = 5;
        this.labelTagAnchor.Text = "Tag Anchor:";
        // 
        // comboBoxTagAnchor
        // 
        this.comboBoxTagAnchor.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;
        this.comboBoxTagAnchor.FormattingEnabled = true;
        this.comboBoxTagAnchor.Items.AddRange(new object[] {
            "Center",
            "Top",
            "Bottom",
            "Left",
            "Right"});
        this.comboBoxTagAnchor.Location = new System.Drawing.Point(100, 102);
        this.comboBoxTagAnchor.Name = "comboBoxTagAnchor";
        this.comboBoxTagAnchor.Size = new System.Drawing.Size(100, 21);
        this.comboBoxTagAnchor.TabIndex = 6;
        this.comboBoxTagAnchor.SelectedIndex = 0;
        // 
        // okButton
        // 
        this.okButton.Location = new System.Drawing.Point(30, 140);
        this.okButton.Name = "okButton";
        this.okButton.Size = new System.Drawing.Size(75, 23);
        this.okButton.TabIndex = 7;
        this.okButton.Text = "OK";
        this.okButton.UseVisualStyleBackColor = true;
        this.okButton.Click += new System.EventHandler(this.OkButton_Click);
        // 
        // cancelButton
        // 
        this.cancelButton.Location = new System.Drawing.Point(115, 140);
        this.cancelButton.Name = "cancelButton";
        this.cancelButton.Size = new System.Drawing.Size(75, 23);
        this.cancelButton.TabIndex = 8;
        this.cancelButton.Text = "Cancel";
        this.cancelButton.UseVisualStyleBackColor = true;
        this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
        // 
        // OffsetInputDialog
        // 
        this.AcceptButton = this.okButton;
        this.CancelButton = this.cancelButton;
        this.ClientSize = new System.Drawing.Size(220, 180);
        this.Controls.Add(this.labelX);
        this.Controls.Add(this.textBoxX);
        this.Controls.Add(this.labelY);
        this.Controls.Add(this.textBoxY);
        this.Controls.Add(this.checkBoxOrient);
        this.Controls.Add(this.labelTagAnchor);
        this.Controls.Add(this.comboBoxTagAnchor);
        this.Controls.Add(this.okButton);
        this.Controls.Add(this.cancelButton);
        this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.StartPosition = WinForms.FormStartPosition.CenterParent;
        this.Text = "Enter Offset Factors";
    }

    private void OkButton_Click(object sender, EventArgs e)
    {
        if (double.TryParse(this.textBoxX.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double offsetX) &&
            double.TryParse(this.textBoxY.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double offsetY))
        {
            this.OffsetX = offsetX;
            this.OffsetY = offsetY;
            switch (this.comboBoxTagAnchor.SelectedItem.ToString())
            {
                case "Top":
                    SelectedTagAnchor = TagAnchor.Top;
                    break;
                case "Bottom":
                    SelectedTagAnchor = TagAnchor.Bottom;
                    break;
                case "Left":
                    SelectedTagAnchor = TagAnchor.Left;
                    break;
                case "Right":
                    SelectedTagAnchor = TagAnchor.Right;
                    break;
                default:
                    SelectedTagAnchor = TagAnchor.Center;
                    break;
            }
            SaveLastOffsets(offsetX, offsetY);
            this.DialogResult = WinForms.DialogResult.OK;
            this.Close();
        }
        else
        {
            WinForms.MessageBox.Show("Please enter valid numeric values for the offsets.",
                            "Invalid Input",
                            WinForms.MessageBoxButtons.OK,
                            WinForms.MessageBoxIcon.Error);
        }
    }

    private void CancelButton_Click(object sender, EventArgs e)
    {
        this.DialogResult = WinForms.DialogResult.Cancel;
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
                if (lines.Length >= 3)
                {
                    if (bool.TryParse(lines[2], out bool lastOrient))
                    {
                        this.checkBoxOrient.Checked = lastOrient;
                    }
                }
                if (lines.Length >= 4)
                {
                    string lastAnchor = lines[3];
                    if (!string.IsNullOrEmpty(lastAnchor))
                    {
                        this.comboBoxTagAnchor.SelectedItem = lastAnchor;
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
            string[] lines = new string[4];
            lines[0] = offsetX.ToString(CultureInfo.InvariantCulture);
            lines[1] = offsetY.ToString(CultureInfo.InvariantCulture);
            lines[2] = checkBoxOrient.Checked.ToString();
            lines[3] = this.comboBoxTagAnchor.SelectedItem.ToString();
            File.WriteAllLines(filePath, lines);
        }
        catch
        {
            // Optionally handle errors here.
        }
    }
}
