//
//  SelectByWorksetsInDocument.cs
//  Revit Ballet
//
//  Selects every element in the entire project that belongs to
//  user-worksets the user picks from a DataGrid.
//
//  The grid shows only worksets that have at least one element
//  and includes a live count of elements per workset.
//

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;

[Transaction(TransactionMode.Manual)]
public class SelectByWorksetsInDocument : IExternalCommand
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

            Document doc = uidoc.Document;

            // Check if document is workshared
            if (!doc.IsWorkshared)
            {
                TaskDialog.Show("Select by Worksets",
                                "This document is not workshared. Worksets are not available.");
                return Result.Cancelled;
            }

            // ------------------------------------------------------------
            // 1. Collect all elements in the project (not element types)
            // ------------------------------------------------------------
            FilteredElementCollector projectCollector =
                new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType();

            // WorksetId → list of element ids in that workset
            Dictionary<WorksetId, List<ElementId>> worksetToElementIds =
                new Dictionary<WorksetId, List<ElementId>>();

            foreach (Element e in projectCollector)
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
                                "No workset-based elements found in the project.");
                return Result.Cancelled;
            }

            // ------------------------------------------------------------
            // 2. Build rows for the custom DataGrid
            // ------------------------------------------------------------
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            Dictionary<string, WorksetId> nameToId = new Dictionary<string, WorksetId>();
            View activeView = doc.ActiveView;

            foreach (var pair in worksetToElementIds)
            {
                Workset ws = doc.GetWorksetTable().GetWorkset(pair.Key);
                if (ws == null) continue;       // safety

                string wsName = ws.Name;
                string wsType = GetWorksetTypeString(ws.Kind);
                string editable = ws.IsEditable ? "Yes" : "No";
                string opened = ws.IsOpen ? "Yes" : "No";
                string visibility = GetWorksetVisibilityString(activeView, ws.Id);

                rows.Add(new Dictionary<string, object>
                {
                    { "Type", wsType },
                    { "Workset",  wsName },
                    { "Elements", pair.Value.Count },
                    { "Editable", editable },
                    { "Opened", opened },
                    { "Visibility", visibility },
                    { "WorksetId", pair.Key }  // Store WorksetId for reliable lookup after edits
                });

                nameToId[wsName] = pair.Key;
            }

            // Sort by Type: User, Standard, Families, Views, then by name
            rows = rows.OrderBy(r => GetTypeSortOrder((string)r["Type"]))
                       .ThenBy(r => (string)r["Workset"])
                       .ToList();

            // ------------------------------------------------------------
            // 3. Ask user to pick one or more worksets
            // ------------------------------------------------------------
            // Set UIDocument for edit mode support
            CustomGUIs.SetCurrentUIDocument(uidoc);

            List<Dictionary<string, object>> pickedRows =
                CustomGUIs.DataGrid(rows,
                                    new List<string> { "Type", "Workset", "Elements", "Editable", "Opened", "Visibility" },
                                    false);

            // Apply any pending edits to worksets (renames, visibility changes, etc.)
            if (CustomGUIs.HasPendingEdits() && !CustomGUIs.WasCancelled())
            {
                CustomGUIs.ApplyCellEditsToEntities();
            }

            if (pickedRows == null || pickedRows.Count == 0)
                return Result.Cancelled;   // user cancelled grid

            // ------------------------------------------------------------
            // 4. Build the final selection set
            // ------------------------------------------------------------
            HashSet<ElementId> finalSel = new HashSet<ElementId>();

            foreach (var row in pickedRows)
            {
                // Use WorksetId from row data instead of looking up by name
                // This ensures that renamed worksets are still found correctly
                if (!row.TryGetValue("WorksetId", out var wsIdObj)) continue;
                if (!(wsIdObj is WorksetId wsId)) continue;

                if (worksetToElementIds.TryGetValue(wsId, out List<ElementId> ids))
                {
                    foreach (ElementId id in ids)
                        finalSel.Add(id);
                }
            }

            if (finalSel.Count == 0)
            {
                TaskDialog.Show("Select by Worksets",
                                "Nothing matched the chosen worksets in the project.");
                return Result.Cancelled;
            }

            // ------------------------------------------------------------
            // 5. Merge with current selection and apply
            // ------------------------------------------------------------
            ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();
            finalSel.UnionWith(currentSelection);

            uidoc.SetSelectionIds(finalSel.ToList());

            return Result.Succeeded;
        }

        private string GetWorksetTypeString(WorksetKind kind)
        {
            switch (kind)
            {
                case WorksetKind.UserWorkset:
                    return "User";
                case WorksetKind.StandardWorkset:
                    return "Standard";
                case WorksetKind.FamilyWorkset:
                    return "Family";
                case WorksetKind.ViewWorkset:
                    return "View";
                default:
                    return kind.ToString();
            }
        }

        private int GetTypeSortOrder(string type)
        {
            switch (type)
            {
                case "User":
                    return 1;
                case "Standard":
                    return 2;
                case "Family":
                    return 3;
                case "View":
                    return 4;
                default:
                    return 5;
            }
        }

        private string GetWorksetVisibilityString(View view, WorksetId worksetId)
        {
            if (view == null)
                return "Unknown";

            try
            {
                WorksetVisibility vis = view.GetWorksetVisibility(worksetId);
                switch (vis)
                {
                    case WorksetVisibility.Visible:
                        return "Shown";
                    case WorksetVisibility.Hidden:
                        return "Hidden";
                    case WorksetVisibility.UseGlobalSetting:
                        // Check the workset's global visibility setting
                        Workset ws = view.Document.GetWorksetTable().GetWorkset(worksetId);
                        if (ws != null)
                        {
                            return ws.IsVisibleByDefault
                                ? "Using Global Settings (Visible)"
                                : "Using Global Settings (Not Visible)";
                        }
                        return "Using Global Settings";
                    default:
                        return vis.ToString();
                }
            }
            catch
            {
                return "N/A";
            }
        }
    }
