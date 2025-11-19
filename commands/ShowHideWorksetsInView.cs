using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitWorksetVisibilityCommands
{
    [Transaction(TransactionMode.Manual)]
    public class ShowWorksetsInCurrentView : IExternalCommand
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

            // Get the active view or its template if applicable
            View activeView = doc.ActiveView;
            View viewToModify = activeView;
            if (activeView.ViewTemplateId != ElementId.InvalidElementId)
            {
                viewToModify = doc.GetElement(activeView.ViewTemplateId) as View;
            }

            // Collect all user worksets
            IList<Workset> worksets = new FilteredWorksetCollector(doc)
                                        .OfKind(WorksetKind.UserWorkset)
                                        .ToWorksets()
                                        .ToList();

            // Prepare data for the custom UI
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            Dictionary<string, WorksetId> worksetNameToId = new Dictionary<string, WorksetId>();

            foreach (Workset ws in worksets)
            {
                WorksetVisibility viewVisibility = viewToModify.GetWorksetVisibility(ws.Id);
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

                Dictionary<string, object> entry = new Dictionary<string, object>
                {
                    { "Workset", ws.Name },
                    { "Visibility", visibilityText }
                };
                entries.Add(entry);

                if (!worksetNameToId.ContainsKey(ws.Name))
                    worksetNameToId.Add(ws.Name, ws.Id);
            }

            // Allow the user to select worksets from the grid
            List<Dictionary<string, object>> selectedEntries =
                CustomGUIs.DataGrid(entries, new List<string> { "Workset", "Visibility" }, spanAllScreens: false);

            if (selectedEntries == null || selectedEntries.Count == 0)
                return Result.Cancelled;

            // Update the workset visibility to 'Visible' for the selected ones
            using (Transaction t = new Transaction(doc, "Show Worksets in Current View"))
            {
                t.Start();
                foreach (Dictionary<string, object> sel in selectedEntries)
                {
                    if (sel.TryGetValue("Workset", out object wsNameObj))
                    {
                        string wsName = wsNameObj as string;
                        if (worksetNameToId.TryGetValue(wsName, out WorksetId wsId))
                            viewToModify.SetWorksetVisibility(wsId, WorksetVisibility.Visible);
                    }
                }
                t.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class HideWorksetsInCurrentView : IExternalCommand
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

            // Get the active view or its template if applicable
            View activeView = doc.ActiveView;
            View viewToModify = activeView;
            if (activeView.ViewTemplateId != ElementId.InvalidElementId)
            {
                viewToModify = doc.GetElement(activeView.ViewTemplateId) as View;
            }

            // Collect all user worksets
            IList<Workset> worksets = new FilteredWorksetCollector(doc)
                                        .OfKind(WorksetKind.UserWorkset)
                                        .ToWorksets()
                                        .ToList();

            // Prepare data for the custom UI
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            Dictionary<string, WorksetId> worksetNameToId = new Dictionary<string, WorksetId>();

            foreach (Workset ws in worksets)
            {
                WorksetVisibility viewVisibility = viewToModify.GetWorksetVisibility(ws.Id);
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

                Dictionary<string, object> entry = new Dictionary<string, object>
                {
                    { "Workset", ws.Name },
                    { "Visibility", visibilityText }
                };
                entries.Add(entry);

                if (!worksetNameToId.ContainsKey(ws.Name))
                    worksetNameToId.Add(ws.Name, ws.Id);
            }

            // Allow the user to select worksets from the grid
            List<Dictionary<string, object>> selectedEntries =
                CustomGUIs.DataGrid(entries, new List<string> { "Workset", "Visibility" }, spanAllScreens: false);

            if (selectedEntries == null || selectedEntries.Count == 0)
                return Result.Cancelled;

            // Update the workset visibility to 'Hidden' for the selected ones
            using (Transaction t = new Transaction(doc, "Hide Worksets in Current View"))
            {
                t.Start();
                foreach (Dictionary<string, object> sel in selectedEntries)
                {
                    if (sel.TryGetValue("Workset", out object wsNameObj))
                    {
                        string wsName = wsNameObj as string;
                        if (worksetNameToId.TryGetValue(wsName, out WorksetId wsId))
                            viewToModify.SetWorksetVisibility(wsId, WorksetVisibility.Hidden);
                    }
                }
                t.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class SetGlobalVisibilityToWorksetInCurrentView : IExternalCommand
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

            // Get the active view or its template if applicable
            View activeView = doc.ActiveView;
            View viewToModify = activeView;
            if (activeView.ViewTemplateId != ElementId.InvalidElementId)
            {
                viewToModify = doc.GetElement(activeView.ViewTemplateId) as View;
            }

            // Collect all user worksets
            IList<Workset> worksets = new FilteredWorksetCollector(doc)
                                        .OfKind(WorksetKind.UserWorkset)
                                        .ToWorksets()
                                        .ToList();

            // Prepare data for the custom UI
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            Dictionary<string, WorksetId> worksetNameToId = new Dictionary<string, WorksetId>();

            foreach (Workset ws in worksets)
            {
                WorksetVisibility viewVisibility = viewToModify.GetWorksetVisibility(ws.Id);
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

                Dictionary<string, object> entry = new Dictionary<string, object>
                {
                    { "Workset", ws.Name },
                    { "Visibility", visibilityText }
                };
                entries.Add(entry);

                if (!worksetNameToId.ContainsKey(ws.Name))
                    worksetNameToId.Add(ws.Name, ws.Id);
            }

            // Allow the user to select worksets from the grid
            List<Dictionary<string, object>> selectedEntries =
                CustomGUIs.DataGrid(entries, new List<string> { "Workset", "Visibility" }, spanAllScreens: false);

            if (selectedEntries == null || selectedEntries.Count == 0)
                return Result.Cancelled;

            // Update the workset visibility to 'Use Global Settings' for the selected ones
            using (Transaction t = new Transaction(doc, "Set Worksets to Use Global Settings"))
            {
                t.Start();
                foreach (Dictionary<string, object> sel in selectedEntries)
                {
                    if (sel.TryGetValue("Workset", out object wsNameObj))
                    {
                        string wsName = wsNameObj as string;
                        if (worksetNameToId.TryGetValue(wsName, out WorksetId wsId))
                            viewToModify.SetWorksetVisibility(wsId, WorksetVisibility.UseGlobalSetting);
                    }
                }
                t.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class IsolateWorksetsInCurrentView : IExternalCommand
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

            // Get the active view or its template if applicable
            View activeView = doc.ActiveView;
            View viewToModify = activeView;
            if (activeView.ViewTemplateId != ElementId.InvalidElementId)
            {
                viewToModify = doc.GetElement(activeView.ViewTemplateId) as View;
            }

            // Collect all user worksets
            IList<Workset> worksets = new FilteredWorksetCollector(doc)
                                        .OfKind(WorksetKind.UserWorkset)
                                        .ToWorksets()
                                        .ToList();

            // Prepare data for the custom UI
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            Dictionary<string, WorksetId> worksetNameToId = new Dictionary<string, WorksetId>();

            foreach (Workset ws in worksets)
            {
                WorksetVisibility viewVisibility = viewToModify.GetWorksetVisibility(ws.Id);
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

                Dictionary<string, object> entry = new Dictionary<string, object>
                {
                    { "Workset", ws.Name },
                    { "Visibility", visibilityText }
                };
                entries.Add(entry);

                if (!worksetNameToId.ContainsKey(ws.Name))
                    worksetNameToId.Add(ws.Name, ws.Id);
            }

            // Allow the user to select the worksets they want to remain visible
            List<Dictionary<string, object>> selectedEntries =
                CustomGUIs.DataGrid(entries, new List<string> { "Workset", "Visibility" }, spanAllScreens: false);

            if (selectedEntries == null || selectedEntries.Count == 0)
                return Result.Cancelled;

            // Create a lookup for the selected workset names
            HashSet<string> selectedWorksetNames = new HashSet<string>();
            foreach (var sel in selectedEntries)
            {
                if (sel.TryGetValue("Workset", out object wsNameObj))
                {
                    string wsName = wsNameObj as string;
                    if (!string.IsNullOrEmpty(wsName))
                        selectedWorksetNames.Add(wsName);
                }
            }

            // In the transaction, set the selected worksets to visible and hide all others.
            using (Transaction t = new Transaction(doc, "Isolate Worksets in Current View"))
            {
                t.Start();
                foreach (Workset ws in worksets)
                {
                    if (worksetNameToId.TryGetValue(ws.Name, out WorksetId wsId))
                    {
                        if (selectedWorksetNames.Contains(ws.Name))
                            viewToModify.SetWorksetVisibility(wsId, WorksetVisibility.Visible);
                        else
                            viewToModify.SetWorksetVisibility(wsId, WorksetVisibility.Hidden);
                    }
                }
                t.Commit();
            }
            return Result.Succeeded;
        }
    }
}
