// CeilingPlanFromFloorPlan.cs
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinFormsControl = System.Windows.Forms.Control;

using RevitView = Autodesk.Revit.DB.View;
using RevitViewport = Autodesk.Revit.DB.Viewport;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace RevitAddin
{
    internal class CeilingPlanOptionsForm : System.Windows.Forms.Form
    {
        private RadioButton radioWithoutDetailing;
        private RadioButton radioWithDetailing;
        private Button okButton;
        private Button cancelButton;

        public bool WithDetailing { get; private set; } = false;

        public CeilingPlanOptionsForm()
        {
            Text = "Ceiling Plan from Floor Plan";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(310, 120);
            MaximizeBox = MinimizeBox = false;

            radioWithoutDetailing = new RadioButton
            {
                Text = "Without Detailing",
                Left = 20, Top = 15, Width = 270, Checked = true
            };
            radioWithDetailing = new RadioButton
            {
                Text = "With Detailing (copy annotations from source)",
                Left = 20, Top = 41, Width = 270
            };

            okButton = new Button
            {
                Text = "OK", Left = 70, Top = 78, Width = 70,
                DialogResult = DialogResult.OK
            };
            cancelButton = new Button
            {
                Text = "Cancel", Left = 160, Top = 78, Width = 70,
                DialogResult = DialogResult.Cancel
            };

            AcceptButton = okButton;
            CancelButton = cancelButton;

            Controls.AddRange(new WinFormsControl[]
            {
                radioWithoutDetailing, radioWithDetailing,
                okButton, cancelButton
            });
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);
            if (DialogResult == DialogResult.OK)
                WithDetailing = radioWithDetailing.Checked;
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

            // Step 1: detailing options
            bool withDetailing;
            using (var form = new CeilingPlanOptionsForm())
            {
                if (form.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;
                withDetailing = form.WithDetailing;
            }

            // Step 2: rename dialog — default pattern appends timestamp + " - Ceiling Plan"
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string initialPattern = "{}_" + timestamp + " - Ceiling Plan";
            var sourceNames = floorPlanViews.Select(v => v.Name.Trim('{', '}')).ToList();
            List<string> newNames;
            using (var renameDialog = new RevitCommands.AdvancedEditDialog(sourceNames, null, "Name Ceiling Plans", initialPattern))
            {
                if (renameDialog.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;
                newNames = renameDialog.GetTransformedValues();
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
                    for (int i = 0; i < floorPlanViews.Count; i++)
                    {
                        var floorPlan = floorPlanViews[i];

                        // Create the ceiling plan on the same level
                        ViewPlan ceilingPlan = ViewPlan.Create(doc, ceilingPlanFamilyType.Id, floorPlan.GenLevel.Id);

                        // Name it (use rename-dialog result, ensure unique)
                        string newName = EnsureUniqueName(newNames[i], existingNames);
                        existingNames.Add(newName);
                        ceilingPlan.Name = newName;

                        // Copy view template (carries scale, visibility, etc.)
                        if (floorPlan.ViewTemplateId != ElementId.InvalidElementId)
                        {
                            try { ceilingPlan.ViewTemplateId = floorPlan.ViewTemplateId; } catch { }
                        }

                        // Copy scale — works when no template controls it, harmless otherwise
                        try { ceilingPlan.Scale = floorPlan.Scale; } catch { }

                        // Copy crop region — always via SetCropShape with a CurveLoop so the
                        // shape maps correctly into the ceiling plan coordinate system.
                        // (Directly copying CropBox fails because floor/ceiling plans have
                        // different view coordinate frames.)
                        if (floorPlan.CropBoxActive)
                        {
                            ceilingPlan.CropBoxActive = true;
                            try
                            {
                                CurveLoop cropLoop = null;

                                // Prefer the actual (possibly non-rectangular) crop shape
                                try
                                {
                                    var sourceManager = floorPlan.GetCropRegionShapeManager();
                                    if (sourceManager.ShapeSet)
                                        cropLoop = sourceManager.GetCropShape().FirstOrDefault();
                                }
                                catch (Autodesk.Revit.Exceptions.InvalidOperationException) { }

                                // Fall back to a rectangular loop built from CropBox corners (Z=0)
                                if (cropLoop == null)
                                {
                                    BoundingBoxXYZ cb = floorPlan.CropBox;
                                    XYZ p1 = new XYZ(cb.Min.X, cb.Min.Y, 0);
                                    XYZ p2 = new XYZ(cb.Max.X, cb.Min.Y, 0);
                                    XYZ p3 = new XYZ(cb.Max.X, cb.Max.Y, 0);
                                    XYZ p4 = new XYZ(cb.Min.X, cb.Max.Y, 0);
                                    cropLoop = new CurveLoop();
                                    cropLoop.Append(Line.CreateBound(p1, p2));
                                    cropLoop.Append(Line.CreateBound(p2, p3));
                                    cropLoop.Append(Line.CreateBound(p3, p4));
                                    cropLoop.Append(Line.CreateBound(p4, p1));
                                }

                                ceilingPlan.GetCropRegionShapeManager().SetCropShape(cropLoop);
                            }
                            catch { }
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
