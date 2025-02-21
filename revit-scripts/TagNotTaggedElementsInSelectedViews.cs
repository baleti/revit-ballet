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

        // --- PATCHED SECTION: Compute counts for each family type using a fast estimate ---
        Dictionary<ElementId, int> totalCounts = new Dictionary<ElementId, int>();
        Dictionary<ElementId, int> taggedCounts = new Dictionary<ElementId, int>();

        foreach (View view in selectedViews)
        {
            // Get all taggable elements in the view that belong to a known family type.
            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && e.CanBeTagged(view) && familyTypes.ContainsKey(e.GetTypeId()));

            // Group by element type and sum counts.
            foreach (var group in collector.GroupBy(e => e.GetTypeId()))
            {
                if (!totalCounts.ContainsKey(group.Key))
                    totalCounts[group.Key] = 0;
                totalCounts[group.Key] += group.Count();
            }

            // Get tagged elements in this view.
            var taggedIds = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(IndependentTag))
                .Cast<IndependentTag>()
                .SelectMany(tag => tag.GetTaggedElementIds()
                                      .Where(linkId => linkId.LinkInstanceId == ElementId.InvalidElementId)
                                      .Select(linkId => linkId.HostElementId))
                .ToHashSet();

            // Group again to count tagged elements.
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
        // Display the DataGrid (this is assumed to be a custom UI helper).
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

        // Now prompt the user for offset factors and whether to orient the tag to the object.
        using (OffsetInputDialog offsetDialog = new OffsetInputDialog())
        {
            if (offsetDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                System.Windows.Forms.MessageBox.Show("Operation cancelled by the user.", "Cancelled");
                return Result.Failed;
            }
            double userOffsetFactorX = offsetDialog.OffsetX;
            double userOffsetFactorY = offsetDialog.OffsetY;
            bool orientToObject = offsetDialog.OrientToObject;

            // Define conversion factor from millimeters to feet.
            const double mmToFeet = 1.0 / 304.8;

            // Begin transaction to create tags.
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

                            // Retrieve the element's bounding box.
                            BoundingBoxXYZ bbox = element.get_BoundingBox(null);
                            if (bbox == null)
                                continue;

                            // --- Crop Region Bounding Box Check ---
                            // Ensure that all four corners of the element's bounding box (projected onto the crop plane)
                            // lie completely within the crop region.
                            if (cropLoops != null && cropLoops.Count > 0)
                            {
                                Transform invTransform = view.CropBox.Transform.Inverse;
                                List<UV> outerPolygon = GetVerticesFromCurveLoop(cropLoops[0]);

                                XYZ pMin = bbox.Min;
                                XYZ pMax = bbox.Max;

                                // Project the bounding box corners onto the crop plane using the crop box transform.
                                XYZ corner1 = ProjectPointToCropPlane(new XYZ(pMin.X, pMin.Y, pMin.Z), view.CropBox.Transform);
                                XYZ corner2 = ProjectPointToCropPlane(new XYZ(pMax.X, pMin.Y, pMin.Z), view.CropBox.Transform);
                                XYZ corner3 = ProjectPointToCropPlane(new XYZ(pMax.X, pMax.Y, pMax.Z), view.CropBox.Transform);
                                XYZ corner4 = ProjectPointToCropPlane(new XYZ(pMin.X, pMax.Y, pMax.Z), view.CropBox.Transform);

                                // Transform each corner into crop region space.
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
                                    // Each corner must be inside the outer crop polygon.
                                    if (!IsPointInsidePolygon(uv, outerPolygon))
                                    {
                                        bboxInside = false;
                                        break;
                                    }
                                    // Also, if the crop region has holes, the corner must not lie inside any hole.
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
                                    continue; // Skip element if its bounding box is not completely within the crop region.
                            }
                            // --- End Crop Region Check ---

                            // Determine tag position based on user input.
                            XYZ tagPosition = null;
                            if (!orientToObject)
                            {
                                // When orientation is off, use the center of the bounding box.
                                double offsetX = userOffsetFactorX * mmToFeet;
                                double offsetY = userOffsetFactorY * mmToFeet;
                                XYZ centerPoint = (bbox.Min + bbox.Max) / 2.0;
                                tagPosition = new XYZ(centerPoint.X + offsetX, centerPoint.Y + offsetY, centerPoint.Z);
                            }
                            else
                            {
                                // When orientation is on, apply the offsets along the element's facing and left directions.
                                double offsetX = userOffsetFactorX * mmToFeet;
                                double offsetY = userOffsetFactorY * mmToFeet;
                                XYZ basePoint = (bbox.Min + bbox.Max) / 2.0;
                                XYZ facing = XYZ.BasisY; // default if no orientation available.
                                if (element is Autodesk.Revit.DB.FamilyInstance fi)
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
    /// Helper method: Projects a 3D point onto the crop plane defined by the crop transform.
    /// </summary>
    private static XYZ ProjectPointToCropPlane(XYZ point, Transform cropTransform)
    {
        XYZ planeOrigin = cropTransform.Origin;
        XYZ normal = cropTransform.BasisZ;
        double distance = (point - planeOrigin).DotProduct(normal);
        return point - distance * normal;
    }

    /// <summary>
    /// Helper method: Determines if a given 2D point (UV) lies inside a polygon defined by a list of UV points.
    /// Points exactly on a polygon edge (within a small tolerance) are considered inside.
    /// Uses the ray-casting algorithm.
    /// </summary>
    private static bool IsPointInsidePolygon(UV point, List<UV> polygon)
    {
        double tol = 1e-6;
        int count = polygon.Count;
        // Check if the point lies exactly on any edge.
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
    /// Helper method: Determines if a point p lies on the line segment between a and b, within a tolerance.
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
    public static bool CanBeTagged(this Element element, View view)
    {
        return element.Category != null &&
               element.Category.HasMaterialQuantities &&
               element.Category.CanAddSubcategory &&
               !element.Category.IsTagCategory;
    }
}

// A dialog form for entering offset factors and whether to orient tags to objects.
public class OffsetInputDialog : System.Windows.Forms.Form
{
    private System.Windows.Forms.Label labelX;
    private System.Windows.Forms.Label labelY;
    private System.Windows.Forms.TextBox textBoxX;
    private System.Windows.Forms.TextBox textBoxY;
    private System.Windows.Forms.CheckBox checkBoxOrient;
    private System.Windows.Forms.Button okButton;
    private System.Windows.Forms.Button cancelButton;

    public double OffsetX { get; private set; }
    public double OffsetY { get; private set; }
    /// <summary>
    /// When true, tag placement will be aligned to each element's orientation.
    /// </summary>
    public bool OrientToObject
    {
        get { return checkBoxOrient.Checked; }
    }

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
        this.checkBoxOrient = new System.Windows.Forms.CheckBox();
        this.okButton = new System.Windows.Forms.Button();
        this.cancelButton = new System.Windows.Forms.Button();

        // labelX
        this.labelX.AutoSize = true;
        this.labelX.Location = new System.Drawing.Point(12, 15);
        this.labelX.Name = "labelX";
        this.labelX.Size = new System.Drawing.Size(70, 13);
        this.labelX.TabIndex = 0;
        this.labelX.Text = "Offset X (mm):";

        // textBoxX
        this.textBoxX.Location = new System.Drawing.Point(100, 12);
        this.textBoxX.Name = "textBoxX";
        this.textBoxX.Size = new System.Drawing.Size(100, 20);
        this.textBoxX.TabIndex = 1;
        this.textBoxX.Text = "10"; // default value in millimeters

        // labelY
        this.labelY.AutoSize = true;
        this.labelY.Location = new System.Drawing.Point(12, 45);
        this.labelY.Name = "labelY";
        this.labelY.Size = new System.Drawing.Size(70, 13);
        this.labelY.TabIndex = 2;
        this.labelY.Text = "Offset Y (mm):";

        // textBoxY
        this.textBoxY.Location = new System.Drawing.Point(100, 42);
        this.textBoxY.Name = "textBoxY";
        this.textBoxY.Size = new System.Drawing.Size(100, 20);
        this.textBoxY.TabIndex = 3;
        this.textBoxY.Text = "10"; // default value in millimeters

        // checkBoxOrient
        this.checkBoxOrient.AutoSize = true;
        this.checkBoxOrient.Location = new System.Drawing.Point(12, 75);
        this.checkBoxOrient.Name = "checkBoxOrient";
        this.checkBoxOrient.Size = new System.Drawing.Size(100, 17);
        this.checkBoxOrient.TabIndex = 4;
        this.checkBoxOrient.Text = "Orient to object";
        this.checkBoxOrient.UseVisualStyleBackColor = true;
        this.checkBoxOrient.Checked = false;  // off by default

        // okButton
        this.okButton.Location = new System.Drawing.Point(30, 110);
        this.okButton.Name = "okButton";
        this.okButton.Size = new System.Drawing.Size(75, 23);
        this.okButton.TabIndex = 5;
        this.okButton.Text = "OK";
        this.okButton.UseVisualStyleBackColor = true;
        this.okButton.Click += new System.EventHandler(this.OkButton_Click);

        // cancelButton
        this.cancelButton.Location = new System.Drawing.Point(115, 110);
        this.cancelButton.Name = "cancelButton";
        this.cancelButton.Size = new System.Drawing.Size(75, 23);
        this.cancelButton.TabIndex = 6;
        this.cancelButton.Text = "Cancel";
        this.cancelButton.UseVisualStyleBackColor = true;
        this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);

        this.AcceptButton = this.okButton;
        this.CancelButton = this.cancelButton;
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(220, 150);
        this.Controls.Add(this.labelX);
        this.Controls.Add(this.textBoxX);
        this.Controls.Add(this.labelY);
        this.Controls.Add(this.textBoxY);
        this.Controls.Add(this.checkBoxOrient);
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
            System.Windows.Forms.MessageBox.Show("Please enter valid numeric values for the offsets.",
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
                if (lines.Length >= 3)
                {
                    if (bool.TryParse(lines[2], out bool lastOrient))
                    {
                        this.checkBoxOrient.Checked = lastOrient;
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
            string[] lines = new string[3];
            lines[0] = offsetX.ToString(CultureInfo.InvariantCulture);
            lines[1] = offsetY.ToString(CultureInfo.InvariantCulture);
            lines[2] = checkBoxOrient.Checked.ToString();
            File.WriteAllLines(filePath, lines);
        }
        catch
        {
            // Optionally handle errors here.
        }
    }
}
