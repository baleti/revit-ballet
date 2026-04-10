// CeilingPlanFromFloorPlan.cs
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using RevitView = Autodesk.Revit.DB.View;
using RevitViewport = Autodesk.Revit.DB.Viewport;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace RevitAddin
{
    internal class CeilingPlanOptionsForm : System.Windows.Forms.Form
    {
        private RadioButton radioWithoutDetailing;
        private RadioButton radioWithDetailing;
        private TextBox txtPrefix;
        private TextBox txtSuffix;
        private Button okButton;
        private Button cancelButton;

        public bool WithDetailing { get; private set; } = false;
        public string Prefix { get; private set; } = "";
        public string Suffix { get; private set; } = "";

        public CeilingPlanOptionsForm()
        {
            Text = "Ceiling Plan from Floor Plan";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(310, 220);
            MaximizeBox = MinimizeBox = false;

            radioWithoutDetailing = new RadioButton
            {
                Text = "Without Detailing",
                Left = 20, Top = 20, Width = 270, Checked = true
            };
            radioWithDetailing = new RadioButton
            {
                Text = "With Detailing (copy annotations from source)",
                Left = 20, Top = 46, Width = 270
            };

            var lblPrefix = new Label { Text = "Prefix:", Left = 20, Top = 90, Width = 55, TextAlign = ContentAlignment.MiddleLeft };
            txtPrefix = new TextBox { Left = 80, Top = 87, Width = 210 };

            var lblSuffix = new Label { Text = "Suffix:", Left = 20, Top = 122, Width = 55, TextAlign = ContentAlignment.MiddleLeft };
            txtSuffix = new TextBox { Left = 80, Top = 119, Width = 210 };

            okButton = new Button
            {
                Text = "OK", Left = 70, Top = 170, Width = 70,
                DialogResult = DialogResult.OK
            };
            cancelButton = new Button
            {
                Text = "Cancel", Left = 160, Top = 170, Width = 70,
                DialogResult = DialogResult.Cancel
            };

            AcceptButton = okButton;
            CancelButton = cancelButton;

            Controls.AddRange(new Control[]
            {
                radioWithoutDetailing, radioWithDetailing,
                lblPrefix, txtPrefix,
                lblSuffix, txtSuffix,
                okButton, cancelButton
            });
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);
            if (DialogResult != DialogResult.OK) return;

            WithDetailing = radioWithDetailing.Checked;
            Prefix = txtPrefix.Text;
            Suffix = txtSuffix.Text;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class CeilingPlanFromFloorPlan : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Collect floor plan views; track which were reached via viewport selection
            var floorPlanViews = new List<ViewPlan>();
            var selectedViewports = new Dictionary<ElementId, List<RevitViewport>>();

            foreach (ElementId id in uiDoc.GetSelectionIds())
            {
                Element elem = doc.GetElement(id);
                if (elem is ViewPlan vp && vp.ViewType == ViewType.FloorPlan)
                {
                    if (!floorPlanViews.Any(v => v.Id == vp.Id))
                        floorPlanViews.Add(vp);
                }
                else if (elem is RevitViewport viewport)
                {
                    var view = doc.GetElement(viewport.ViewId) as ViewPlan;
                    if (view != null && view.ViewType == ViewType.FloorPlan)
                    {
                        if (!floorPlanViews.Any(v => v.Id == view.Id))
                            floorPlanViews.Add(view);
                        if (!selectedViewports.ContainsKey(view.Id))
                            selectedViewports[view.Id] = new List<RevitViewport>();
                        selectedViewports[view.Id].Add(viewport);
                    }
                }
            }

            // Fall back to active view when nothing is selected
            if (floorPlanViews.Count == 0)
            {
                if (uiDoc.ActiveView is ViewPlan activeVp && activeVp.ViewType == ViewType.FloorPlan)
                    floorPlanViews.Add(activeVp);
                else
                {
                    TaskDialog.Show("Ceiling Plan from Floor Plan",
                        "No floor plan views or viewports selected, and the active view is not a floor plan.");
                    return Result.Cancelled;
                }
            }

            // Find ceiling plan view family type
            var ceilingPlanFamilyType = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.CeilingPlan);

            if (ceilingPlanFamilyType == null)
            {
                TaskDialog.Show("Ceiling Plan from Floor Plan",
                    "No Ceiling Plan view family type found in the project.");
                return Result.Cancelled;
            }

            // Show options dialog
            bool withDetailing;
            string prefix, suffix;
            using (var form = new CeilingPlanOptionsForm())
            {
                if (form.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;
                withDetailing = form.WithDetailing;
                prefix = form.Prefix;
                suffix = form.Suffix;
            }

            // Snapshot of existing view names for uniqueness checks
            var existingNames = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitView))
                    .Cast<RevitView>()
                    .Select(v => v.Name),
                StringComparer.OrdinalIgnoreCase);

            using (var trans = new Transaction(doc, "Create Ceiling Plans from Floor Plans"))
            {
                trans.Start();
                try
                {
                    foreach (var floorPlan in floorPlanViews)
                    {
                        // Create the ceiling plan on the same level
                        ViewPlan ceilingPlan = ViewPlan.Create(doc, ceilingPlanFamilyType.Id, floorPlan.GenLevel.Id);

                        // Name it
                        string baseName = floorPlan.Name.Trim('{', '}');
                        string newName = BuildName(baseName, prefix, suffix);
                        newName = EnsureUniqueName(newName, existingNames);
                        existingNames.Add(newName);
                        ceilingPlan.Name = newName;

                        // Copy crop region
                        if (floorPlan.CropBoxActive)
                        {
                            ceilingPlan.CropBoxActive = true;
                            try
                            {
                                var sourceManager = floorPlan.GetCropRegionShapeManager();
                                if (sourceManager.ShapeSet)
                                {
                                    var cropShape = sourceManager.GetCropShape();
                                    if (cropShape != null && cropShape.Count > 0)
                                        ceilingPlan.GetCropRegionShapeManager().SetCropShape(cropShape[0]);
                                }
                                else
                                {
                                    ceilingPlan.CropBox = floorPlan.CropBox;
                                }
                            }
                            catch
                            {
                                try { ceilingPlan.CropBox = floorPlan.CropBox; } catch { }
                            }
                        }

                        // Optionally copy view-owned annotation/detail elements
                        if (withDetailing)
                            CopyViewOwnedElements(doc, floorPlan, ceilingPlan);

                        // Viewport replacement: only when viewports were explicitly selected
                        if (selectedViewports.TryGetValue(floorPlan.Id, out var viewportList))
                        {
                            foreach (var sourceVp in viewportList)
                            {
                                ElementId sheetId = sourceVp.SheetId;
                                XYZ position = sourceVp.GetBoxCenter();

                                // Place new ceiling plan at the same position on the same sheet
                                try { Viewport.Create(doc, sheetId, ceilingPlan.Id, position); } catch { }

                                // Remove original floor plan viewport (never the view itself)
                                try { doc.Delete(sourceVp.Id); } catch { }
                            }
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    message = ex.Message;
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }

        private static void CopyViewOwnedElements(Document doc, RevitView source, RevitView target)
        {
            var elementIds = new List<ElementId>();

            foreach (Element e in new FilteredElementCollector(doc, source.Id).WhereElementIsNotElementType())
            {
                if (e.OwnerViewId != source.Id) continue;

                bool canCopy = false;
                if (e.Category != null && e.Category.CategoryType == CategoryType.Annotation)
                    canCopy = true;
                else if (e.Category != null && (
                    e.Category.Id.AsLong() == (long)BuiltInCategory.OST_DetailComponents ||
                    e.Category.Id.AsLong() == (long)BuiltInCategory.OST_IOSDetailGroups ||
                    e.Category.Id.AsLong() == (long)BuiltInCategory.OST_Lines ||
                    e.Category.Id.AsLong() == (long)BuiltInCategory.OST_RasterImages ||
                    e.Category.Id.AsLong() == (long)BuiltInCategory.OST_InsulationLines))
                    canCopy = true;
                else if (e is ImportInstance imp && imp.ViewSpecific)
                    canCopy = true;

                if (canCopy)
                    elementIds.Add(e.Id);
            }

            if (elementIds.Count > 0)
            {
                try
                {
                    ElementTransformUtils.CopyElements(source, elementIds, target, Transform.Identity, null);
                }
                catch { }
            }
        }

        private static string BuildName(string baseName, string prefix, string suffix)
        {
            bool hasPrefix = !string.IsNullOrEmpty(prefix);
            bool hasSuffix = !string.IsNullOrEmpty(suffix);
            if (!hasPrefix && !hasSuffix)
                return baseName + " - Ceiling Plan";
            return (hasPrefix ? prefix : "") + baseName + (hasSuffix ? suffix : "");
        }

        private static string EnsureUniqueName(string name, HashSet<string> existingNames)
        {
            if (!existingNames.Contains(name))
                return name;
            int counter = 2;
            while (existingNames.Contains(name + " " + counter))
                counter++;
            return name + " " + counter;
        }
    }
}
