using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class HideWorksetsInViews : IExternalCommand
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

        try
        {
            // Get currently selected elements using SelectionModeManager
            ICollection<ElementId> selectedElementIds = uidoc.GetSelectionIds();

            // Check if any views or viewports are selected
            List<View> targetViews = new List<View>();

            foreach (ElementId id in selectedElementIds)
            {
                Element elem = doc.GetElement(id);
                if (elem == null) continue;

                if (elem is View view)
                {
                    // Don't include sheets or schedules
                    if (!(view is ViewSheet || view is ViewSchedule))
                    {
                        targetViews.Add(view);
                    }
                }
                else if (elem is Viewport viewport)
                {
                    // Get the view from the viewport
                    View viewFromViewport = doc.GetElement(viewport.ViewId) as View;
                    if (viewFromViewport != null &&
                        !(viewFromViewport is ViewSheet || viewFromViewport is ViewSchedule))
                    {
                        targetViews.Add(viewFromViewport);
                    }
                }
            }

            // If no views selected, use the active view
            if (targetViews.Count == 0)
            {
                View activeView = doc.ActiveView;
                View viewToModify = activeView;
                if (activeView.ViewTemplateId != ElementId.InvalidElementId)
                {
                    viewToModify = doc.GetElement(activeView.ViewTemplateId) as View;
                }
                targetViews.Add(viewToModify);
            }

            // Collect all user worksets
            IList<Workset> worksets = new FilteredWorksetCollector(doc)
                                        .OfKind(WorksetKind.UserWorkset)
                                        .ToWorksets()
                                        .ToList();

            if (worksets.Count == 0)
            {
                TaskDialog.Show("Info", "This document has no user worksets.");
                return Result.Cancelled;
            }

            // Prepare data for the custom UI
            // If multiple views, show worksets that are visible in ANY of the views
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            Dictionary<string, WorksetId> worksetNameToId = new Dictionary<string, WorksetId>();

            foreach (Workset ws in worksets)
            {
                // Collect visibility status from all target views
                List<string> viewVisibilityList = new List<string>();
                foreach (View view in targetViews)
                {
                    WorksetVisibility viewVisibility = view.GetWorksetVisibility(ws.Id);
                    string visibilityText;
                    if (viewVisibility == WorksetVisibility.Visible)
                        visibilityText = "Shown";
                    else if (viewVisibility == WorksetVisibility.Hidden)
                        visibilityText = "Hidden";
                    else if (viewVisibility == WorksetVisibility.UseGlobalSetting)
                        visibilityText = ws.IsVisibleByDefault ?
                            "Using Global Settings (Visible)" :
                            "Using Global Settings (Not Visible)";
                    else
                        visibilityText = "Unknown";

                    viewVisibilityList.Add(visibilityText);
                }

                // For display, show first view's visibility or a summary
                string displayVisibility = viewVisibilityList[0];
                if (targetViews.Count > 1)
                {
                    // Show a summary if multiple views have different visibility
                    if (viewVisibilityList.Distinct().Count() > 1)
                    {
                        displayVisibility = "Mixed";
                    }
                }

                Dictionary<string, object> entry = new Dictionary<string, object>
                {
                    { "Workset", ws.Name },
                    { "Visibility", displayVisibility }
                };
                entries.Add(entry);

                if (!worksetNameToId.ContainsKey(ws.Name))
                    worksetNameToId.Add(ws.Name, ws.Id);
            }

            // Allow the user to select worksets from the grid
            string gridTitle = targetViews.Count == 1
                ? $"Hide Worksets in {targetViews[0].Name}"
                : $"Hide Worksets in {targetViews.Count} Views";

            List<Dictionary<string, object>> selectedEntries =
                CustomGUIs.DataGrid(entries, new List<string> { "Workset", "Visibility" }, false);

            if (selectedEntries == null || selectedEntries.Count == 0)
                return Result.Cancelled;

            // Update the workset visibility to 'Hidden' for the selected worksets in all target views
            using (Transaction t = new Transaction(doc, "Hide Worksets in Views"))
            {
                t.Start();
                foreach (Dictionary<string, object> sel in selectedEntries)
                {
                    if (sel.TryGetValue("Workset", out object wsNameObj))
                    {
                        string wsName = wsNameObj as string;
                        if (worksetNameToId.TryGetValue(wsName, out WorksetId wsId))
                        {
                            foreach (View view in targetViews)
                            {
                                view.SetWorksetVisibility(wsId, WorksetVisibility.Hidden);
                            }
                        }
                    }
                }
                t.Commit();
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            TaskDialog.Show("Error", $"An error occurred:\n{ex.Message}");
            return Result.Failed;
        }
    }
}
