using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class FilterTags : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get currently selected elements
        ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

        if (!selectedIds.Any())
        {
            TaskDialog.Show("Filter Tags", "No elements selected.");
            return Result.Cancelled;
        }

        // Filter for IndependentTag elements
        List<IndependentTag> selectedTags = new List<IndependentTag>();
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem is IndependentTag tag)
            {
                selectedTags.Add(tag);
            }
            else
            {
                // If any element is not a tag, show dialog and exit
                TaskDialog.Show("Filter Tags",
                    "Please select only tag elements (IndependentTag). Command will now exit.");
                return Result.Cancelled;
            }
        }

        // Build data to display
        // Columns required: TagText, OwnerView, SheetNumber, SheetName, ElementId
        List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();

        // Cache for view -> (sheetNumber, sheetName) lookups (optional for performance)
        Dictionary<ElementId, (string viewName, string sheetNumber, string sheetName)> viewSheetInfoCache
            = new Dictionary<ElementId, (string, string, string)>();

        foreach (var tag in selectedTags)
        {
            var dict = new Dictionary<string, object>();

            // Tag text
            dict["TagText"] = tag.TagText;

            // Retrieve the owner view, and from that the sheet info (if placed)
            ElementId ownerViewId = tag.OwnerViewId;
            string viewName = "";
            string sheetNumber = "";
            string sheetName = "";

            if (!viewSheetInfoCache.ContainsKey(ownerViewId))
            {
                var view = doc.GetElement(ownerViewId) as Autodesk.Revit.DB.View;
                if (view != null)
                {
                    viewName = view.Name;
                    
                    // Attempt to find if this view is placed on a sheet via a Viewport
                    var viewport = new FilteredElementCollector(doc)
                        .OfClass(typeof(Viewport))
                        .Cast<Viewport>()
                        .FirstOrDefault(vp => vp.ViewId == view.Id);

                    if (viewport != null)
                    {
                        ViewSheet sheet = doc.GetElement(viewport.SheetId) as ViewSheet;
                        if (sheet != null)
                        {
                            sheetNumber = sheet.SheetNumber;
                            sheetName = sheet.Name;
                        }
                    }
                }

                viewSheetInfoCache[ownerViewId] = (viewName, sheetNumber, sheetName);
            }

            // Fetch from cache
            var cached = viewSheetInfoCache[ownerViewId];
            dict["OwnerView"] = cached.viewName;
            dict["SheetNumber"] = cached.sheetNumber;
            dict["SheetName"] = cached.sheetName;

            // Finally, ElementId (at the end)
            dict["ElementId"] = tag.Id.IntegerValue;

            entries.Add(dict);
        }

        // Define the property names (columns) to display in the desired order
        // per request: "TagText", "OwnerView", "SheetNumber", "SheetName", "ElementId" (at the end)
        List<string> propertyNames = new List<string>
        {
            "TagText",
            "OwnerView",
            "SheetNumber",
            "SheetName",
            "ElementId"
        };

        // Call the provided DataGrid method (from your CustomGUIs class)
        var filteredEntries = CustomGUIs.DataGrid(
            entries,
            propertyNames,
            spanAllScreens: false,
            initialSelectionIndices: null
        );

        // If user pressed ESC or closed the form, filteredEntries may be empty
        if (filteredEntries == null || !filteredEntries.Any())
        {
            // No change to selection
            return Result.Cancelled;
        }

        // Otherwise, let's update the Revit selection with only the chosen elements
        List<ElementId> chosenIds = new List<ElementId>();
        foreach (var row in filteredEntries)
        {
            if (row.ContainsKey("ElementId")
                && int.TryParse(row["ElementId"].ToString(), out int intId))
            {
                chosenIds.Add(new ElementId(intId));
            }
        }

        using (var tx = new Transaction(doc, "Filter Tags Selection"))
        {
            tx.Start();
            uidoc.Selection.SetElementIds(chosenIds);
            tx.Commit();
        }

        return Result.Succeeded;
    }
}
