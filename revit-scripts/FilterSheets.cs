#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace MyCompany.RevitCommands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class FilterSheets : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uidoc.Document;

            // --- get the user’s current selection -------------------------------------------
            IList<ElementId> selIds = uidoc.GetSelectionIds().ToList();
            if (selIds.Count == 0)
            {
                TaskDialog.Show("FilterSheets",
                    "Nothing is selected.\nPlease pick one or more sheets and run the command again.");
                return Result.Cancelled;
            }

            // --- keep only sheets ------------------------------------------------------------
            List<ViewSheet> sheets = selIds
                .Select(id => doc.GetElement(id))
                .OfType<ViewSheet>()
                .ToList();

            if (sheets.Count == 0)
            {
                TaskDialog.Show("FilterSheets",
                    "None of the selected elements are sheets.");
                return Result.Cancelled;
            }

            // --- visible column headers ------------------------------------------------------
            var columns = new List<string>
            {
                "Sheet Number",
                "Sheet Title",
                "Revision Sequence",
                "Revision Date",
                "Revision Description",
                "Issued To",
                "Issued By"
            };

            // --- build grid data + quick lookup map ------------------------------------------
            var gridData        = new List<Dictionary<string, object>>();
            var sheetNumToIdMap = new Dictionary<string, ElementId>();   // <SheetNumber, Id>

            foreach (ViewSheet vs in sheets)
            {
                IList<ElementId> revIds = vs.GetAllRevisionIds();

                var seq      = new List<string>();
                var date     = new List<string>();
                var desc     = new List<string>();
                var issuedTo = new List<string>();
                var issuedBy = new List<string>();

                foreach (ElementId rid in revIds)
                {
                    if (!(doc.GetElement(rid) is Revision rev)) continue;

                    seq.     Add(rev.SequenceNumber.ToString());
                    date.    Add(rev.RevisionDate ?? string.Empty);
                    desc.    Add(rev.Description  ?? string.Empty);
                    issuedTo.Add(rev.IssuedTo     ?? string.Empty);
                    issuedBy.Add(rev.IssuedBy     ?? string.Empty);
                }

                gridData.Add(new Dictionary<string, object>
                {
                    ["Sheet Number"]         = vs.SheetNumber,
                    ["Sheet Title"]          = vs.Name,
                    ["Revision Sequence"]    = string.Join(", ", seq),
                    ["Revision Date"]        = string.Join(", ", date),
                    ["Revision Description"] = string.Join(", ", desc),
                    ["Issued To"]            = string.Join(", ", issuedTo),
                    ["Issued By"]            = string.Join(", ", issuedBy)
                });

                // remember the id for later
                sheetNumToIdMap[vs.SheetNumber] = vs.Id;
            }

            // --- show chooser grid -----------------------------------------------------------
            List<Dictionary<string, object>> chosen =
                CustomGUIs.DataGrid(gridData, columns, spanAllScreens: false);

            if (chosen == null || chosen.Count == 0)
                return Result.Cancelled;   // user cancelled or unchecked everything

            // --- translate chosen rows back into ElementIds ----------------------------------
            var newIds = new List<ElementId>();

            foreach (Dictionary<string, object> row in chosen)
            {
                if (row.TryGetValue("Sheet Number", out object snObj) && snObj != null)
                {
                    string sheetNum = snObj.ToString();
                    if (sheetNumToIdMap.TryGetValue(sheetNum, out ElementId id))
                        newIds.Add(id);
                }
            }

            if (newIds.Count == 0)
                return Result.Cancelled;   // should not happen, but play safe

            // --- REPLACE the selection with only the sheets the user checked -----------------
            uidoc.SetSelectionIds(newIds);

            return Result.Succeeded;
        }
    }
}
