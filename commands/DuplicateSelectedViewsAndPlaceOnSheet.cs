using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

// Aliases to disambiguate Revit types from Windows Forms types
using RevitView = Autodesk.Revit.DB.View;
using RevitViewport = Autodesk.Revit.DB.Viewport;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace RevitAddin
{
    /// <summary>
    /// A Windows Form that prompts the user to select a duplication option (without number of duplicates).
    /// Number of duplicates is determined by the number of selected sheets.
    /// </summary>
    internal class DuplicationOptionsFormForSheets : System.Windows.Forms.Form
    {
        public DuplicationMode SelectedMode { get; private set; } = DuplicationMode.WithoutDetailing;

        private RadioButton radioWithoutDetailing;
        private RadioButton radioWithDetailing;
        private RadioButton radioDependent;
        private Button okButton;
        private Button cancelButton;

        public DuplicationOptionsFormForSheets()
        {
            // Set basic form properties
            this.Text = "Duplication Options";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(250, 140);
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Create radio buttons
            radioWithoutDetailing = new RadioButton()
            {
                Text = "Without Detailing",
                Left = 20,
                Top = 20,
                Width = 200,
                Checked = true
            };
            radioWithDetailing = new RadioButton()
            {
                Text = "With Detailing",
                Left = 20,
                Top = 50,
                Width = 200
            };
            radioDependent = new RadioButton()
            {
                Text = "Dependent",
                Left = 20,
                Top = 80,
                Width = 200
            };

            // Create OK and Cancel buttons
            okButton = new Button()
            {
                Text = "OK",
                Left = 50,
                Width = 70,
                Top = 110,
                DialogResult = DialogResult.OK
            };
            cancelButton = new Button()
            {
                Text = "Cancel",
                Left = 130,
                Width = 70,
                Top = 110,
                DialogResult = DialogResult.Cancel
            };

            // Add controls to the form
            this.Controls.Add(radioWithoutDetailing);
            this.Controls.Add(radioWithDetailing);
            this.Controls.Add(radioDependent);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);

            // Set the Accept and Cancel buttons
            this.AcceptButton = okButton;
            this.CancelButton = cancelButton;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);

            // When closing with OK, store the selected duplication mode
            if (this.DialogResult == DialogResult.OK)
            {
                if (radioWithoutDetailing.Checked)
                    SelectedMode = DuplicationMode.WithoutDetailing;
                else if (radioWithDetailing.Checked)
                    SelectedMode = DuplicationMode.WithDetailing;
                else if (radioDependent.Checked)
                    SelectedMode = DuplicationMode.Dependent;
            }
        }
    }

    /// <summary>
    /// Command that duplicates selected views and places them on selected sheets.
    /// If both views and sheets are selected, uses them directly.
    /// Otherwise, shows a DataGrid to select target sheets.
    /// Number of duplicates = number of selected sheets.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class DuplicateSelectedViewsAndPlaceOnSheet : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();

            // ─────────────────────────────────────────────────────────────
            // 1. Collect views and sheets from current selection
            // ─────────────────────────────────────────────────────────────
            List<RevitView> selectedViews = new List<RevitView>();
            List<ViewSheet> selectedSheets = new List<ViewSheet>();

            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element is ViewSheet sheet)
                {
                    selectedSheets.Add(sheet);
                }
                else if (element is RevitView view && !(view is ViewSheet))
                {
                    selectedViews.Add(view);
                }
                else if (element is RevitViewport viewport)
                {
                    // If a viewport is selected, add its corresponding view
                    RevitView viewFromViewport = doc.GetElement(viewport.ViewId) as RevitView;
                    if (viewFromViewport != null && !(viewFromViewport is ViewSheet))
                    {
                        selectedViews.Add(viewFromViewport);
                    }
                }
            }

            if (selectedViews.Count == 0)
            {
                TaskDialog.Show("Error", "No valid views selected.");
                return Result.Cancelled;
            }

            // ─────────────────────────────────────────────────────────────
            // 2. If no sheets selected, show DataGrid to select sheets
            // ─────────────────────────────────────────────────────────────
            if (selectedSheets.Count == 0)
            {
                // Collect all sheets in the document
                List<ViewSheet> allSheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .ToList();

                if (allSheets.Count == 0)
                {
                    TaskDialog.Show("Error", "No sheets found in the document.");
                    return Result.Cancelled;
                }

                // Get browser organization columns for sheets
                List<BrowserOrganizationHelper.BrowserColumn> browserColumns =
                    BrowserOrganizationHelper.GetBrowserColumnsForViews(doc, allSheets.Cast<RevitView>());

                // Prepare data for the grid
                List<Dictionary<string, object>> gridData = new List<Dictionary<string, object>>();

                foreach (ViewSheet sheet in allSheets)
                {
                    var dict = new Dictionary<string, object>();

                    // Add browser organization columns first
                    BrowserOrganizationHelper.AddBrowserColumnsToDict(dict, sheet, doc, browserColumns);

                    // Then add standard columns
                    dict["SheetNumber"] = sheet.SheetNumber;
                    dict["Name"] = sheet.Name;
                    dict["ElementIdObject"] = sheet.Id;
                    dict["__OriginalObject"] = sheet;

                    gridData.Add(dict);
                }

                // Sort by browser organization columns
                if (browserColumns != null && browserColumns.Count > 0)
                {
                    gridData = BrowserOrganizationHelper.SortByBrowserColumns(gridData, browserColumns);
                }
                else
                {
                    gridData = gridData.OrderBy(row =>
                    {
                        if (row.ContainsKey("__OriginalObject") && row["__OriginalObject"] is ViewSheet s)
                            return s.SheetNumber;
                        return "";
                    }).ToList();
                }

                // Column headers
                List<string> columns = new List<string>();
                columns.AddRange(browserColumns.Select(bc => bc.Name));
                columns.Add("SheetNumber");
                columns.Add("Name");

                // Find sheets where selected views are currently placed
                List<int> initialSelectionIndices = new List<int>();
                HashSet<ElementId> sheetsWithSelectedViews = new HashSet<ElementId>();

                foreach (RevitView view in selectedViews)
                {
                    var viewport = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitViewport))
                        .Cast<RevitViewport>()
                        .FirstOrDefault(vp => vp.ViewId == view.Id);

                    if (viewport != null)
                    {
                        sheetsWithSelectedViews.Add(viewport.SheetId);
                    }
                }

                // Find indices of these sheets in gridData
                for (int i = 0; i < gridData.Count; i++)
                {
                    if (gridData[i].ContainsKey("ElementIdObject") &&
                        gridData[i]["ElementIdObject"] is ElementId sheetId &&
                        sheetsWithSelectedViews.Contains(sheetId))
                    {
                        initialSelectionIndices.Add(i);
                    }
                }

                // Show the grid with initial selection
                CustomGUIs.SetCurrentUIDocument(uidoc);
                List<Dictionary<string, object>> selectedRows =
                    CustomGUIs.DataGrid(gridData, columns, false, initialSelectionIndices);

                if (selectedRows == null || selectedRows.Count == 0)
                {
                    return Result.Cancelled;
                }

                // Extract selected sheets
                selectedSheets = CustomGUIs.ExtractOriginalObjects<ViewSheet>(selectedRows);
            }

            if (selectedSheets.Count == 0)
            {
                TaskDialog.Show("Error", "No sheets selected.");
                return Result.Cancelled;
            }

            // ─────────────────────────────────────────────────────────────
            // 3. Prompt the user for duplication options
            // ─────────────────────────────────────────────────────────────
            DuplicationMode mode;
            using (var form = new DuplicationOptionsFormForSheets())
            {
                if (form.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;

                mode = form.SelectedMode;
            }

            // ─────────────────────────────────────────────────────────────
            // 4. Duplicate views and place on sheets
            // ─────────────────────────────────────────────────────────────
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            int duplicateCount = selectedSheets.Count;

            using (Transaction trans = new Transaction(doc, "Duplicate Views and Place on Sheets"))
            {
                trans.Start();
                try
                {
                    foreach (var sourceView in selectedViews)
                    {
                        if (sourceView == null)
                            continue;

                        // Determine the duplication option based on the selected mode
                        ViewDuplicateOption duplicateOption;
                        switch (mode)
                        {
                            case DuplicationMode.Dependent:
                                duplicateOption = ViewDuplicateOption.AsDependent;
                                break;
                            case DuplicationMode.WithDetailing:
                                duplicateOption = ViewDuplicateOption.WithDetailing;
                                break;
                            default:
                                duplicateOption = ViewDuplicateOption.Duplicate;
                                break;
                        }

                        // Check if the view supports duplication with the selected option
                        if (!sourceView.CanViewBeDuplicated(duplicateOption))
                        {
                            continue;
                        }

                        // Get the position for placing duplicates
                        XYZ placementPosition = GetViewPlacementPosition(doc, sourceView);

                        // Create duplicates and place on sheets
                        for (int i = 0; i < duplicateCount; i++)
                        {
                            // Duplicate the view
                            ElementId newViewId = sourceView.Duplicate(duplicateOption);
                            RevitView dupView = doc.GetElement(newViewId) as RevitView;

                            if (dupView != null)
                            {
                                // Rename the new view
                                string baseName = sourceView.Name.Trim(new char[] { '{', '}' });

                                if (duplicateCount == 1)
                                {
                                    dupView.Name = $"{baseName} - Copy {timestamp}";
                                }
                                else
                                {
                                    dupView.Name = $"{baseName} - Copy {timestamp}_{i + 1}";
                                }

                                // Place the duplicate on the corresponding sheet
                                ViewSheet targetSheet = selectedSheets[i];

                                // Skip if view cannot be added to sheet
                                if (!RevitViewport.CanAddViewToSheet(doc, targetSheet.Id, dupView.Id))
                                {
                                    continue;
                                }

                                // Determine placement position for this sheet
                                XYZ sheetPlacementPosition;
                                if (placementPosition != null)
                                {
                                    // Use the same position as source view
                                    sheetPlacementPosition = placementPosition;
                                }
                                else
                                {
                                    // Use title block center or sheet center
                                    sheetPlacementPosition = GetTitleBlockCenter(doc, targetSheet)
                                        ?? GetSheetCenter(targetSheet);
                                }

                                // Create viewport on sheet
                                RevitViewport.Create(doc, targetSheet.Id, dupView.Id, sheetPlacementPosition);
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

        /// <summary>
        /// Gets the placement position for a view.
        /// If the view is placed on a sheet, returns its viewport center.
        /// Otherwise returns null (caller should use title block center).
        /// </summary>
        private XYZ GetViewPlacementPosition(Document doc, RevitView view)
        {
            var viewport = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitViewport))
                .Cast<RevitViewport>()
                .FirstOrDefault(vp => vp.ViewId == view.Id);

            if (viewport != null)
            {
                return viewport.GetBoxCenter();
            }

            return null;
        }

        /// <summary>
        /// Gets the center of the title block on a sheet.
        /// </summary>
        private XYZ GetTitleBlockCenter(Document doc, ViewSheet sheet)
        {
            var titleBlock = new FilteredElementCollector(doc)
                .OwnedByView(sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilyInstance>()
                .FirstOrDefault();

            if (titleBlock == null)
                return null;

            BoundingBoxXYZ bb = titleBlock.get_BoundingBox(sheet);
            if (bb == null)
                return null;

            return (bb.Min + bb.Max) * 0.5;
        }

        /// <summary>
        /// Gets the center of a sheet.
        /// </summary>
        private XYZ GetSheetCenter(ViewSheet sheet)
        {
            return new XYZ(
                (sheet.Outline.Max.U + sheet.Outline.Min.U) / 2,
                (sheet.Outline.Max.V + sheet.Outline.Min.V) / 2,
                0);
        }
    }
}
