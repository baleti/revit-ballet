//
//  SelectByWorksetsInViews.cs
//  Revit 2024 – C# 7.3
//
//  Selects every element that is **both**
//    • on a user-workset the user picks from a DataGrid, **and**
//    • visible in the target view(s).
//
//  If views are selected, uses those views; otherwise uses active view.
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

using TaskDialog = Autodesk.Revit.UI.TaskDialog;

[Transaction(TransactionMode.Manual)]
public class SelectByWorksetsInViews : IExternalCommand
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

            // ------------------------------------------------------------
            // 0. Determine target views (selected views or active view)
            // ------------------------------------------------------------
            ICollection<ElementId> currentSelection = uidoc.GetSelectionIds();
            List<View> targetViews = new List<View>();

            bool hasSelectedViews = currentSelection.Any(id =>
            {
                Element elem = doc.GetElement(id);
                return elem is View;
            });

            if (hasSelectedViews)
            {
                // Use selected views
                targetViews = currentSelection
                    .Select(id => doc.GetElement(id))
                    .OfType<View>()
                    .ToList();
            }
            else
            {
                // Use current view only
                targetViews.Add(doc.ActiveView);
            }

            // ------------------------------------------------------------
            // 1. Collect all real (non-type) elements that lie in the target views
            // ------------------------------------------------------------
            // WorksetId → list of element ids in that workset & view(s)
            Dictionary<WorksetId, List<ElementId>> worksetToElementIds =
                new Dictionary<WorksetId, List<ElementId>>();

            foreach (View view in targetViews)
            {
                FilteredElementCollector viewCollector =
                    new FilteredElementCollector(doc, view.Id)
                        .WhereElementIsNotElementType();

                foreach (Element e in viewCollector)
                {
                    WorksetId wsId = e.WorksetId;
                    if (wsId == WorksetId.InvalidWorksetId) continue;           // no user workset
                    if (!worksetToElementIds.TryGetValue(wsId, out var list))   // first hit
                    {
                        list = new List<ElementId>();
                        worksetToElementIds.Add(wsId, list);
                    }
                    if (!list.Contains(e.Id))  // Avoid duplicates across views
                        list.Add(e.Id);
                }
            }

            if (worksetToElementIds.Count == 0)
            {
                string viewText = hasSelectedViews ? "selected views" : "active view";
                TaskDialog.Show("Select by Worksets",
                                $"No workset-based elements are visible in the {viewText}.");
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
                string wsType = GetWorksetTypeString(ws.Kind);
                string editable = ws.IsEditable ? "Yes" : "No";
                string opened = ws.IsOpen ? "Yes" : "No";

                var row = new Dictionary<string, object>
                {
                    { "Type", wsType },
                    { "Workset",  wsName },
                    { "Elements", pair.Value.Count },
                    { "Editable", editable },
                    { "Opened", opened }
                };

                // Only add visibility column for single-view mode
                if (!hasSelectedViews)
                {
                    string visibility = GetWorksetVisibilityString(targetViews[0], ws.Id);
                    row.Add("Visibility", visibility);
                }

                rows.Add(row);

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

            var columns = hasSelectedViews
                ? new List<string> { "Type", "Workset", "Elements", "Editable", "Opened" }
                : new List<string> { "Type", "Workset", "Elements", "Editable", "Opened", "Visibility" };

            List<Dictionary<string, object>> pickedRows =
                CustomGUIs.DataGrid(rows, columns, false);

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
                string viewText = hasSelectedViews ? "these views" : "this view";
                TaskDialog.Show("Select by Worksets",
                                $"Nothing matched the chosen worksets in {viewText}.");
                return Result.Cancelled;
            }

            // ------------------------------------------------------------
            // 5. Apply the selection
            // ------------------------------------------------------------
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
