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
    //  WinForm
    // ─────────────────────────────────────────────────────────────
    internal class SheetDuplicationOptionsForm : WinForms.Form
    {
        private readonly WinForms.RadioButton _radioEmpty;
        private readonly WinForms.RadioButton _radioWithDetailing;
        private readonly WinForms.NumericUpDown _numCopies;

        public SheetDuplicateOption SelectedOption { get; private set; } =
            SheetDuplicateOption.DuplicateEmptySheet;

        public int CopyCount { get; private set; } = 1;

        public SheetDuplicationOptionsForm()
        {
            Text            = "Duplicate Sheet Options";
            FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            StartPosition   = WinForms.FormStartPosition.CenterScreen;
            ClientSize      = new Size(260, 170);
            MaximizeBox = MinimizeBox = false;

            _radioEmpty = new WinForms.RadioButton
            {
                Text    = "Empty Sheet",
                Left    = 20,
                Top     = 20,
                Width   = 200,
                Checked = true
            };
            _radioWithDetailing = new WinForms.RadioButton
            {
                Text  = "With Sheet Detailing",
                Left  = 20,
                Top   = 50,
                Width = 200
            };

            var lblCopies = new WinForms.Label
            {
                Text  = "Number of copies:",
                Left  = 20,
                Top   = 85,
                Width = 120
            };
            _numCopies = new WinForms.NumericUpDown
            {
                Left    = 150,
                Top     = 80,
                Width   = 60,
                Minimum = 1,
                Maximum = 50,
                Value   = 1
            };

            var ok     = new WinForms.Button { Text = "OK",     Left = 50,  Width = 70, Top = 120, DialogResult = WinForms.DialogResult.OK };
            var cancel = new WinForms.Button { Text = "Cancel", Left = 130, Width = 70, Top = 120, DialogResult = WinForms.DialogResult.Cancel };

            AcceptButton = ok;
            CancelButton = cancel;

            Controls.AddRange(new WinForms.Control[]
            {
                _radioEmpty, _radioWithDetailing, lblCopies, _numCopies, ok, cancel
            });
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);
            if (DialogResult != WinForms.DialogResult.OK) return;

            SelectedOption = _radioWithDetailing.Checked
                ? SheetDuplicateOption.DuplicateSheetWithDetailing
                : SheetDuplicateOption.DuplicateEmptySheet;

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

        public static void DuplicateSheets(
            Document doc,
            IEnumerable<ViewSheet> sheets,
            SheetDuplicateOption option,
            int copies)
        {
            var usedNumbers =
                new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Select(s => s.SheetNumber)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (ViewSheet sheet in sheets)
            {
                for (int n = 0; n < copies; ++n)
                {
                    ElementId newId = sheet.Duplicate(option);
                    if (newId == ElementId.InvalidElementId) continue;

                    var dup = doc.GetElement(newId) as ViewSheet;
                    if (dup == null) continue;

                    string nextNo = NextAvailableNumber(sheet.SheetNumber, usedNumbers);
                    usedNumbers.Add(nextNo);

                    dup.SheetNumber = nextNo;
                    dup.Name        = sheet.Name;
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
            var selIds = uiDoc.Selection.GetElementIds();
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
                    if (sheet != null) sheets.Add(sheet);
                }
            }

            if (sheets.Count == 0 && uiDoc.ActiveView is ViewSheet activeSheet)
                sheets.Add(activeSheet);

            if (sheets.Count == 0)
            {
                TaskDialog.Show("Duplicate Sheets", "No sheet view selected.");
                return Result.Cancelled;
            }

            SheetDuplicateOption option;
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
                try
                {
                    SheetDuplicator.DuplicateSheets(doc, sheets, option, copies);
                    t.Commit();
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    message = ex.Message;
                    return Result.Failed;
                }
            }
            return Result.Succeeded;
        }
    }
}
