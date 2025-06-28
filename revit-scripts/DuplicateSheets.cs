// DuplicateSelectedSheets.cs  –  Revit 2024 / C# 7.3 compatible
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

namespace RevitAddin
{
    // ─────────────────────────────────────────────────────────────
    //  Custom Sheet Duplication Options
    // ─────────────────────────────────────────────────────────────
    public enum CustomSheetDuplicateOption
    {
        EmptySheet,
        WithSheetDetailing,
        WithSheetDetailingViewsWithoutDetailing,
        WithSheetDetailingViewsWithDetailing
    }
    
    // ─────────────────────────────────────────────────────────────
    //  WinForm
    // ─────────────────────────────────────────────────────────────
    internal class SheetDuplicationOptionsForm : WinForms.Form
    {
        private readonly WinForms.RadioButton _radioEmpty;
        private readonly WinForms.RadioButton _radioWithDetailing;
        private readonly WinForms.RadioButton _radioViewsWithoutDetailing;
        private readonly WinForms.RadioButton _radioViewsWithDetailing;
        private readonly WinForms.NumericUpDown _numCopies;

        public CustomSheetDuplicateOption SelectedOption { get; private set; } =
            CustomSheetDuplicateOption.EmptySheet;

        public int CopyCount { get; private set; } = 1;

        public SheetDuplicationOptionsForm()
        {
            Text            = "Duplicate Sheet Options";
            FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            StartPosition   = WinForms.FormStartPosition.CenterScreen;
            ClientSize      = new Size(340, 220);
            MaximizeBox = MinimizeBox = false;

            _radioEmpty = new WinForms.RadioButton
            {
                Text    = "Empty Sheet",
                Left    = 20,
                Top     = 20,
                Width   = 300,
                Checked = true
            };
            _radioWithDetailing = new WinForms.RadioButton
            {
                Text  = "With Sheet Detailing",
                Left  = 20,
                Top   = 45,
                Width = 300
            };
            _radioViewsWithoutDetailing = new WinForms.RadioButton
            {
                Text    = "With Sheet Detailing, Views Without Detailing",
                Left    = 20,
                Top     = 70,
                Width   = 300
            };
            _radioViewsWithDetailing = new WinForms.RadioButton
            {
                Text  = "With Sheet Detailing, Views With Detailing",
                Left  = 20,
                Top   = 95,
                Width = 300
            };

            var lblCopies = new WinForms.Label
            {
                Text  = "Number of copies:",
                Left  = 20,
                Top   = 135,
                Width = 120
            };
            _numCopies = new WinForms.NumericUpDown
            {
                Left    = 150,
                Top     = 130,
                Width   = 60,
                Minimum = 1,
                Maximum = 50,
                Value   = 1
            };

            var ok     = new WinForms.Button { Text = "OK",     Left = 90,  Width = 70, Top = 170, DialogResult = WinForms.DialogResult.OK };
            var cancel = new WinForms.Button { Text = "Cancel", Left = 170, Width = 70, Top = 170, DialogResult = WinForms.DialogResult.Cancel };

            AcceptButton = ok;
            CancelButton = cancel;

            Controls.AddRange(new WinForms.Control[]
            {
                _radioEmpty, _radioWithDetailing, _radioViewsWithoutDetailing, _radioViewsWithDetailing, 
                lblCopies, _numCopies, ok, cancel
            });
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);
            if (DialogResult != WinForms.DialogResult.OK) return;

            if (_radioEmpty.Checked)
                SelectedOption = CustomSheetDuplicateOption.EmptySheet;
            else if (_radioWithDetailing.Checked)
                SelectedOption = CustomSheetDuplicateOption.WithSheetDetailing;
            else if (_radioViewsWithoutDetailing.Checked)
                SelectedOption = CustomSheetDuplicateOption.WithSheetDetailingViewsWithoutDetailing;
            else if (_radioViewsWithDetailing.Checked)
                SelectedOption = CustomSheetDuplicateOption.WithSheetDetailingViewsWithDetailing;

            CopyCount = (int)_numCopies.Value;
        }
    }

    // ─────────────────────────────────────────────────────────────
    //  Core duplication helper
    // ─────────────────────────────────────────────────────────────
    internal static class SheetDuplicator
    {
        private static readonly Regex _lastDigits =
            new Regex(@"(\d+)(?!.*\d)", RegexOptions.Compiled);

        public static List<ViewSheet> DuplicateSheets(
            Document doc,
            IEnumerable<ViewSheet> sheets,
            CustomSheetDuplicateOption option,
            int copies)
        {
            var usedNumbers =
                new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Select(s => s.SheetNumber)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var createdSheets = new List<ViewSheet>();

            foreach (ViewSheet sheet in sheets)
            {
                for (int n = 0; n < copies; ++n)
                {
                    // Determine the Revit sheet duplication option
                    SheetDuplicateOption sheetDupOption = 
                        (option == CustomSheetDuplicateOption.EmptySheet) 
                        ? SheetDuplicateOption.DuplicateEmptySheet 
                        : SheetDuplicateOption.DuplicateSheetWithDetailing;

                    ElementId newSheetId = sheet.Duplicate(sheetDupOption);
                    if (newSheetId == ElementId.InvalidElementId) continue;

                    var dupSheet = doc.GetElement(newSheetId) as ViewSheet;
                    if (dupSheet == null) continue;

                    string nextNo = NextAvailableNumber(sheet.SheetNumber, usedNumbers);
                    usedNumbers.Add(nextNo);

                    dupSheet.SheetNumber = nextNo;
                    dupSheet.Name = sheet.Name;

                    // Handle view duplication for the new options
                    if (option == CustomSheetDuplicateOption.WithSheetDetailingViewsWithoutDetailing ||
                        option == CustomSheetDuplicateOption.WithSheetDetailingViewsWithDetailing)
                    {
                        DuplicateAndPlaceViews(doc, sheet, dupSheet, option);
                    }

                    createdSheets.Add(dupSheet);
                }
            }

            return createdSheets;
        }

        private static void DuplicateAndPlaceViews(
            Document doc, 
            ViewSheet originalSheet, 
            ViewSheet newSheet,
            CustomSheetDuplicateOption option)
        {
            // Get all viewports on the original sheet
            var originalViewports = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .Where(vp => vp.SheetId == originalSheet.Id)
                .ToList();

            // Delete all viewports from the new sheet (they were copied with sheet detailing)
            var newSheetViewports = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .Where(vp => vp.SheetId == newSheet.Id)
                .ToList();

            foreach (var vp in newSheetViewports)
            {
                doc.Delete(vp.Id);
            }

            // Get timestamp for view naming
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // Collect views to duplicate and their positions
            var viewsToDuplicate = new List<(View view, XYZ position)>();
            
            foreach (var viewport in originalViewports)
            {
                View originalView = doc.GetElement(viewport.ViewId) as View;
                if (originalView == null) continue;

                // Skip schedules and other non-placeable views
                if (originalView.ViewType == ViewType.Schedule ||
                    originalView.ViewType == ViewType.DrawingSheet)
                    continue;

                XYZ position = viewport.GetBoxCenter();
                viewsToDuplicate.Add((originalView, position));
            }

            // Now duplicate all views first (fast operation)
            var duplicatedViews = new List<(View original, View duplicate, XYZ position)>();
            
            foreach (var (originalView, position) in viewsToDuplicate)
            {
                // Check if it's a legend - legends don't need to be duplicated
                if (originalView.ViewType == ViewType.Legend)
                {
                    duplicatedViews.Add((originalView, originalView, position));
                    continue;
                }

                // Determine duplication option
                ViewDuplicateOption dupOption = 
                    option == CustomSheetDuplicateOption.WithSheetDetailingViewsWithDetailing
                    ? ViewDuplicateOption.WithDetailing
                    : ViewDuplicateOption.Duplicate;

                // Check if view can be duplicated with selected option
                if (!originalView.CanViewBeDuplicated(dupOption))
                    continue;

                // Duplicate the view
                ElementId newViewId = originalView.Duplicate(dupOption);
                if (newViewId == ElementId.InvalidElementId) continue;

                View dupView = doc.GetElement(newViewId) as View;
                if (dupView == null) continue;

                // Rename the duplicated view with timestamp
                try
                {
                    dupView.Name = originalView.Name + "_" + timestamp;
                }
                catch
                {
                    // If name exists, add counter
                    int counter = 1;
                    while (counter < 100)
                    {
                        try
                        {
                            dupView.Name = originalView.Name + "_" + timestamp + "_" + counter;
                            break;
                        }
                        catch
                        {
                            counter++;
                        }
                    }
                }

                duplicatedViews.Add((originalView, dupView, position));
            }

            // Now place all duplicated views on the new sheet
            foreach (var (original, duplicate, position) in duplicatedViews)
            {
                try
                {
                    Viewport.Create(doc, newSheet.Id, duplicate.Id, position);
                }
                catch
                {
                    // Some views might not be placeable, skip them
                }
            }
        }

        private static string NextAvailableNumber(
            string source,
            HashSet<string> used)
        {
            Match m = _lastDigits.Match(source);

            // No digits → append _1, _2, …
            if (!m.Success)
            {
                int suffix = 1;
                string numCandidate = source + "_" + suffix;
                while (used.Contains(numCandidate, StringComparer.OrdinalIgnoreCase))
                {
                    suffix++;
                    numCandidate = source + "_" + suffix;
                }
                return numCandidate;
            }

            int start  = m.Index;
            int len    = m.Length;
            string pre = source.Substring(0, start);
            string digits = m.Value;
            string post = source.Substring(start + len);

            int number = int.Parse(digits);
            int width  = len;              // keep zero-padding
            int next   = number + 1;

            string candidate = Build(pre, next, width, post);
            while (used.Contains(candidate, StringComparer.OrdinalIgnoreCase))
            {
                next++;
                candidate = Build(pre, next, width, post);
            }
            return candidate;
        }

        private static string Build(string pre, int num, int width, string post)
        {
            return pre + num.ToString().PadLeft(width, '0') + post;
        }
    }

    // ─────────────────────────────────────────────────────────────
    //  Revit external command
    // ─────────────────────────────────────────────────────────────
    [Transaction(TransactionMode.Manual)]
    public class DuplicateSelectedSheets : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uiDoc.Document;
            var selIds = uiDoc.GetSelectionIds();
            var sheets = new List<ViewSheet>();

            foreach (ElementId id in selIds)
            {
                Element e = doc.GetElement(id);
                if (e is ViewSheet vs)
                {
                    sheets.Add(vs);
                }
                else if (e is Viewport vp && vp.SheetId != ElementId.InvalidElementId)
                {
                    var sheet = doc.GetElement(vp.SheetId) as ViewSheet;
                    if (sheet != null && !sheets.Contains(sheet)) 
                        sheets.Add(sheet);
                }
            }

            if (sheets.Count == 0 && uiDoc.ActiveView is ViewSheet activeSheet)
                sheets.Add(activeSheet);

            if (sheets.Count == 0)
            {
                TaskDialog.Show("Duplicate Sheets", "No sheet view selected.");
                return Result.Cancelled;
            }

            CustomSheetDuplicateOption option;
            int copies;
            using (var dlg = new SheetDuplicationOptionsForm())
            {
                if (dlg.ShowDialog() != WinForms.DialogResult.OK) return Result.Cancelled;
                option = dlg.SelectedOption;
                copies = dlg.CopyCount;
            }

            using (var t = new Transaction(doc, "Duplicate Selected Sheets"))
            {
                t.Start();
                List<ViewSheet> createdSheets = null;
                try
                {
                    createdSheets = SheetDuplicator.DuplicateSheets(doc, sheets, option, copies);
                    t.Commit();
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    message = ex.Message;
                    return Result.Failed;
                }

                // Open all newly created sheets
                if (createdSheets != null && createdSheets.Count > 0)
                {
                    // Open the first created sheet as the active view
                    uiDoc.ActiveView = createdSheets[0];
                    
                    // Open remaining sheets in new UI views if there are multiple
                    for (int i = 1; i < createdSheets.Count; i++)
                    {
                        try
                        {
                            // Try to open in a new view window
                            UIView uiView = uiDoc.GetOpenUIViews()
                                .FirstOrDefault(v => v.ViewId == createdSheets[i].Id);
                            
                            if (uiView == null)
                            {
                                // If not already open, make it active temporarily
                                // This will open it in a new tab/window depending on Revit settings
                                uiDoc.ActiveView = createdSheets[i];
                            }
                        }
                        catch
                        {
                            // If opening fails, continue with next sheet
                        }
                    }
                    
                    // Set the first sheet as active again
                    if (createdSheets.Count > 1)
                    {
                        uiDoc.ActiveView = createdSheets[0];
                    }
                }
            }
            return Result.Succeeded;
        }
    }
}
