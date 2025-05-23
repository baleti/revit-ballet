// ──────────────────────────────────────────────────────────
//  Aliases to avoid Autodesk ↔ WinForms name collisions
// ──────────────────────────────────────────────────────────
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using WinForms = System.Windows.Forms;
using DB       = Autodesk.Revit.DB;

using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class CombineViewSetsIntoANewSet : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData,
                          ref string message,
                          DB.ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument    uidoc = uiapp.ActiveUIDocument;
        DB.Document   doc   = uidoc.Document;

        // 1️⃣ grab ALL view-sheet sets in the document
        List<DB.ViewSheetSet> allSets =
            new DB.FilteredElementCollector(doc)
                  .OfClass(typeof(DB.ViewSheetSet))
                  .Cast<DB.ViewSheetSet>()
                  .ToList();

        if (!allSets.Any())
        {
            TaskDialog.Show("Combine View Sets",
                            "There are no view-sheet sets in this project.");
            return Result.Cancelled;
        }

        // 2️⃣ build rows for your CustomGUIs.DataGrid
        var rows = new List<Dictionary<string, object>>();

        foreach (DB.ViewSheetSet set in allSets)
        {
            var views      = set.Views.Cast<DB.View>().ToList();
            int viewCount  = views.Count;
            int sheetCount = views.Count(v => v is DB.ViewSheet);

            rows.Add(new Dictionary<string, object>
            {
                ["Name"]        = set.Name,
                ["Description"] = $"{viewCount} view(s), {sheetCount} sheet(s)",
                ["View Count"]  = viewCount
            });
        }

        var cols = new List<string> { "Name", "Description", "View Count" };

        IList<Dictionary<string, object>> picked =
            CustomGUIs.DataGrid(rows, cols, false);

        if (picked == null || picked.Count == 0)
            return Result.Cancelled;          // user hit Cancel / closed grid

        // 3️⃣ ask for the new set name with a tiny WinForms dialog
        string newSetName;
        using (var dlg = new NewSetNameForm())
        {
            if (dlg.ShowDialog() != WinForms.DialogResult.OK)
                return Result.Cancelled;

            newSetName = dlg.SetName.Trim();
        }

        if (string.IsNullOrEmpty(newSetName))
        {
            message = "You must supply a name for the new view-sheet set.";
            return Result.Failed;
        }
        if (allSets.Any(s => s.Name.Equals(newSetName,
                                           StringComparison.OrdinalIgnoreCase)))
        {
            message = $"A set named “{newSetName}” already exists.";
            return Result.Failed;
        }

        // 4️⃣ build a combined ViewSet (no duplicates)
        DB.ViewSet combined = new DB.ViewSet();
        HashSet<DB.ElementId> added = new HashSet<DB.ElementId>();

        foreach (var row in picked)
        {
            string name = row["Name"].ToString();
            DB.ViewSheetSet src = allSets.First(s => s.Name == name);

            foreach (DB.View v in src.Views)
            {
                if (added.Add(v.Id))          // HashSet stops duplicates
                    combined.Insert(v);
            }
        }

        if (combined.IsEmpty)
        {
            message = "The chosen sets contain no views or sheets.";
            return Result.Failed;
        }

        // 5️⃣ create the new set in a transaction
        using (DB.Transaction tx = new DB.Transaction(doc,
                                   "Create combined view-sheet set"))
        {
            tx.Start();

            DB.PrintManager pm = doc.PrintManager;
            pm.PrintSetup.CurrentPrintSetting = pm.PrintSetup.InSession;

            DB.ViewSheetSetting vss = pm.ViewSheetSetting;
            // NOTE: ViewSheetSetting doesn't expose Views directly;
            //       we assign to its *CurrentViewSheetSet* first
            vss.CurrentViewSheetSet.Views = combined;
            vss.SaveAs(newSetName);

            tx.Commit();
        }

        return Result.Succeeded;
    }
}

// ──────────────────────────────────────────────────────────
//  Minimal WinForms prompt for the new set name
// ──────────────────────────────────────────────────────────
public class NewSetNameForm : WinForms.Form
{
    private readonly WinForms.TextBox _txt;
    public string SetName => _txt.Text;

    public NewSetNameForm()
    {
        // basic form chrome
        Text            = "New View-Sheet Set";
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        StartPosition   = WinForms.FormStartPosition.CenterScreen;
        MinimizeBox     = false;
        MaximizeBox     = false;
        Width           = 330;
        Height          = 140;

        // label
        var lbl = new WinForms.Label
        {
            Text     = "Enter a name for the new set:",
            AutoSize = true,
            Left     = 10,
            Top      = 15
        };
        Controls.Add(lbl);

        // textbox
        _txt = new WinForms.TextBox
        {
            Left  = 10,
            Top   = 40,
            Width = 300
        };
        Controls.Add(_txt);

        // buttons
        var ok = new WinForms.Button
        {
            Text         = "OK",
            DialogResult = WinForms.DialogResult.OK,
            Left         = 155,
            Width        = 70,
            Top          = 75
        };
        var cancel = new WinForms.Button
        {
            Text         = "Cancel",
            DialogResult = WinForms.DialogResult.Cancel,
            Left         = 235,
            Width        = 70,
            Top          = 75
        };

        Controls.Add(ok);
        Controls.Add(cancel);

        AcceptButton = ok;
        CancelButton = cancel;
    }
}
