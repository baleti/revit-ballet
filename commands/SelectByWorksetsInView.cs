//
//  SelectByWorksetsInView.cs
//  Revit 2024 – C# 7.3
//
//  Selects every element that is **both**
//    • on a user-workset the user picks from a DataGrid, **and**
//    • visible in the currently–active view.
//
//  The grid shows only worksets that have at least one such element
//  and includes a live count of elements per workset.
//

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RevitWorksetVisibilityCommands
{
    [Transaction(TransactionMode.Manual)]
    public class SelectByWorksetsInView : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }

            Document doc       = uidoc.Document;
            View     activeView = doc.ActiveView;

            // ------------------------------------------------------------
            // 1. Collect all real (non-type) elements that lie in the view
            // ------------------------------------------------------------
            FilteredElementCollector viewCollector =
                new FilteredElementCollector(doc, activeView.Id)
                    .WhereElementIsNotElementType();

            // WorksetId → list of element ids in that workset & view
            Dictionary<WorksetId, List<ElementId>> worksetToElementIds =
                new Dictionary<WorksetId, List<ElementId>>();

            foreach (Element e in viewCollector)
            {
                WorksetId wsId = e.WorksetId;
                if (wsId == WorksetId.InvalidWorksetId) continue;           // no user workset
                if (!worksetToElementIds.TryGetValue(wsId, out var list))   // first hit
                {
                    list = new List<ElementId>();
                    worksetToElementIds.Add(wsId, list);
                }
                list.Add(e.Id);
            }

            if (worksetToElementIds.Count == 0)
            {
                TaskDialog.Show("Select by Worksets",
                                "No workset-based elements are visible in the active view.");
                return Result.Cancelled;
            }

            // ------------------------------------------------------------
            // 2. Build rows for the custom DataGrid
            // ------------------------------------------------------------
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            Dictionary<string, WorksetId>    nameToId = new Dictionary<string, WorksetId>();

            foreach (var pair in worksetToElementIds)
            {
                Workset ws = doc.GetWorksetTable().GetWorkset(pair.Key);
                if (ws == null) continue;       // safety

                string wsName = ws.Name;

                rows.Add(new Dictionary<string, object>
                {
                    { "Workset",  wsName },
                    { "Elements", pair.Value.Count }
                });

                nameToId[wsName] = pair.Key;
            }

            // ------------------------------------------------------------
            // 3. Ask user to pick one or more worksets
            // ------------------------------------------------------------
            List<Dictionary<string, object>> pickedRows =
                CustomGUIs.DataGrid(rows,
                                    new List<string> { "Workset", "Elements" },
                                    spanAllScreens: false);

            if (pickedRows == null || pickedRows.Count == 0)
                return Result.Cancelled;   // user cancelled grid

            // ------------------------------------------------------------
            // 4. Build the final selection set
            // ------------------------------------------------------------
            HashSet<ElementId> finalSel = new HashSet<ElementId>();

            foreach (var row in pickedRows)
            {
                if (!row.TryGetValue("Workset", out var nameObj)) continue;

                string wsName = nameObj as string;
                if (string.IsNullOrEmpty(wsName)) continue;

                if (nameToId.TryGetValue(wsName, out WorksetId wsId) &&
                    worksetToElementIds.TryGetValue(wsId, out List<ElementId> ids))
                {
                    foreach (ElementId id in ids)
                        finalSel.Add(id);
                }
            }

            if (finalSel.Count == 0)
            {
                TaskDialog.Show("Select by Worksets",
                                "Nothing matched the chosen worksets in this view.");
                return Result.Cancelled;
            }

            // ------------------------------------------------------------
            // 5. Apply the selection
            // ------------------------------------------------------------
            uidoc.SetSelectionIds(finalSel.ToList());

            return Result.Succeeded;
        }
    }
}
