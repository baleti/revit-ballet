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
public class CombineSelectionSetsIntoANewSet : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData,
                          ref string message,
                          DB.ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument    uidoc = uiapp.ActiveUIDocument;
        DB.Document   doc   = uidoc.Document;

        // 1️⃣ all stored selection sets (SelectionFilterElement)
        List<DB.SelectionFilterElement> allSets =
            new DB.FilteredElementCollector(doc)
                  .OfClass(typeof(DB.SelectionFilterElement))
                  .Cast<DB.SelectionFilterElement>()
                  .ToList();

        if (!allSets.Any())
        {
            TaskDialog.Show("Combine Selection Sets",
                            "There are no selection sets in this project.");
            return Result.Cancelled;
        }

        // 2️⃣ grid rows for CustomGUIs.DataGrid
        var rows = new List<Dictionary<string, object>>();

        foreach (DB.SelectionFilterElement set in allSets)
        {
            var elementIds = set.GetElementIds();
            int elementCount = elementIds.Count;

            rows.Add(new Dictionary<string, object>
            {
                ["Name"]          = set.Name,
                ["Description"]   = $"{elementCount} element(s)",
                ["Element Count"] = elementCount
            });
        }

        var cols = new List<string> { "Name", "Description", "Element Count" };

        IList<Dictionary<string, object>> picked =
            CustomGUIs.DataGrid(rows, cols, false);

        if (picked == null || picked.Count == 0)
            return Result.Cancelled;  // user aborted

        // 3️⃣ tiny WinForms dialog for the new-set name
        string newSetName;
        using (var dlg = new NewSetNameForm())
        {
            if (dlg.ShowDialog() != WinForms.DialogResult.OK)
                return Result.Cancelled;

            newSetName = dlg.SetName.Trim();
        }

        if (string.IsNullOrEmpty(newSetName))
        {
            message = "You must supply a name for the new selection set.";
            return Result.Failed;
        }

        // 3 b️⃣ check for name clash
        DB.SelectionFilterElement existingSet =
            allSets.FirstOrDefault(s => s.Name.Equals(newSetName,
                                                      StringComparison.OrdinalIgnoreCase));

        bool overwriteExisting = false;
        if (existingSet != null)
        {
            WinForms.DialogResult res =
                WinForms.MessageBox.Show(
                    $"A set named \"{newSetName}\" already exists.\n\n" +
                    "Do you want to overwrite it?",
                    "Overwrite Existing Set?",
                    WinForms.MessageBoxButtons.YesNo,
                    WinForms.MessageBoxIcon.Question,
                    WinForms.MessageBoxDefaultButton.Button2);

            if (res != WinForms.DialogResult.Yes)
                return Result.Cancelled;

            overwriteExisting = true;
        }

        // 4️⃣ build a combined set of ElementIds (no duplicates)
        HashSet<DB.ElementId> combinedIds = new HashSet<DB.ElementId>();

        foreach (var row in picked)
        {
            string name = row["Name"].ToString();
            DB.SelectionFilterElement src = allSets.First(s => s.Name == name);

            foreach (DB.ElementId id in src.GetElementIds())
            {
                combinedIds.Add(id);
            }
        }

        if (combinedIds.Count == 0)
        {
            message = "The chosen sets contain no elements.";
            return Result.Failed;
        }

        // 5️⃣ create or update inside ONE transaction
        using (DB.Transaction tx = new DB.Transaction(doc,
                                  overwriteExisting
                                      ? "Overwrite selection set"
                                      : "Create combined selection set"))
        {
            tx.Start();

            if (overwriteExisting)
            {
                // Update the existing selection set
                existingSet.SetElementIds(combinedIds);
            }
            else
            {
                // Create a new selection set
                DB.SelectionFilterElement newSet = 
                    DB.SelectionFilterElement.Create(doc, newSetName);
                newSet.SetElementIds(combinedIds);
            }

            tx.Commit();
        }

        return Result.Succeeded;
    }
}
