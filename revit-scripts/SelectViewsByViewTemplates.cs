#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms; // for TaskDialog
using RevitView = Autodesk.Revit.DB.View;
using RevitViewport = Autodesk.Revit.DB.Viewport;
#endregion

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class SelectViewsByViewTemplates : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            // Get the current document.
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 1. Collect all view templates from the project.
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitView));
                List<RevitView> viewTemplates = collector
                    .Cast<RevitView>()
                    .Where(v => v.IsTemplate)
                    .ToList();

                if (viewTemplates.Count == 0)
                {
                    TaskDialog.Show("Select Views By View Templates", "No view templates were found in the project.");
                    return Result.Cancelled;
                }

                // 2. Collect all non-template views.
                List<RevitView> nonTemplateViews = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitView))
                    .Cast<RevitView>()
                    .Where(v => !v.IsTemplate)
                    .ToList();

                // 3. Build a dictionary mapping each view (by id) to the set of sheet numbers on which it appears.
                //    A view appears on a sheet if it has a viewport placed on that sheet.
                Dictionary<ElementId, HashSet<string>> viewIdToSheetNumbers = new Dictionary<ElementId, HashSet<string>>();
                FilteredElementCollector viewportCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitViewport));
                foreach (RevitViewport vp in viewportCollector.Cast<RevitViewport>())
                {
                    if (vp.SheetId != ElementId.InvalidElementId)
                    {
                        ViewSheet sheet = doc.GetElement(vp.SheetId) as ViewSheet;
                        if (sheet != null)
                        {
                            string sheetNumber = sheet.SheetNumber;
                            if (!viewIdToSheetNumbers.ContainsKey(vp.ViewId))
                            {
                                viewIdToSheetNumbers[vp.ViewId] = new HashSet<string>();
                            }
                            viewIdToSheetNumbers[vp.ViewId].Add(sheetNumber);
                        }
                    }
                }

                // 4. Compute usage counts for each view template.
                //    Also, for each view template, compute the union of sheet numbers from all views using it.
                var viewTemplateUsageCount = nonTemplateViews
                    .GroupBy(v => v.ViewTemplateId)
                    .ToDictionary(g => g.Key, g => g.Count());

                // 5. Prepare a list of dictionary entries for the DataGrid.
                //    Each entry will include the view template's Id, Name, usage Count,
                //    and a comma-delimited list of sheet numbers where views using it appear.
                List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
                foreach (RevitView vt in viewTemplates)
                {
                    // Get the usage count; if not found, default to zero.
                    int count = viewTemplateUsageCount.ContainsKey(vt.Id) ? viewTemplateUsageCount[vt.Id] : 0;

                    // For this view template, collect sheet numbers from all views that use it.
                    HashSet<string> sheetsForTemplate = new HashSet<string>();
                    foreach (RevitView view in nonTemplateViews.Where(v => v.ViewTemplateId == vt.Id))
                    {
                        if (viewIdToSheetNumbers.TryGetValue(view.Id, out var sheetNumbers))
                        {
                            sheetsForTemplate.UnionWith(sheetNumbers);
                        }
                    }
                    string sheetsList = string.Join(", ", sheetsForTemplate);

                    var dict = new Dictionary<string, object>
                    {
                        // Using the integer value of the ElementId.
                        ["Id"] = vt.Id.IntegerValue,
                        ["Name"] = vt.Name,
                        ["Count (Views)"] = count,
                        ["Sheet Numbers"] = sheetsList
                    };
                    entries.Add(dict);
                }

                // Specify the property names (columns) to display.
                List<string> propertyNames = new List<string> { "Id", "Name", "Count (Views)", "Sheet Numbers" };

                // 6. Show the DataGrid UI to the user.
                //    This custom function displays the list and returns the selected rows.
                List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);

                if (selectedEntries == null || selectedEntries.Count == 0)
                {
                    return Result.Cancelled;
                }

                // 7. Extract the selected view template IDs.
                HashSet<ElementId> selectedTemplateIds = new HashSet<ElementId>();
                foreach (var entry in selectedEntries)
                {
                    int idInt = Convert.ToInt32(entry["Id"]);
                    selectedTemplateIds.Add(new ElementId(idInt));
                }

                // 8. Find all non-template views whose ViewTemplateId matches any of the selected templates.
                List<ElementId> viewsToSelect = new List<ElementId>();
                FilteredElementCollector viewCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitView));
                foreach (RevitView view in viewCollector.Cast<RevitView>())
                {
                    // Skip the view templates themselves.
                    if (view.IsTemplate)
                        continue;

                    // Check if the view is assigned one of the selected view templates.
                    if (selectedTemplateIds.Contains(view.ViewTemplateId))
                    {
                        viewsToSelect.Add(view.Id);
                    }
                }

                if (viewsToSelect.Count == 0)
                {
                    TaskDialog.Show("Select Views By View Templates", "No views use the selected view templates.");
                    return Result.Succeeded;
                }

                // 9. Set the selection in the active UIDocument.
                uidoc.SetSelectionIds(viewsToSelect);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }
}
