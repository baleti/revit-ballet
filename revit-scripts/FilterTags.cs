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

        // Filter for both IndependentTag and SpatialElementTag elements
        List<Element> selectedTags = new List<Element>();
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem is IndependentTag || elem is SpatialElementTag)
            {
                selectedTags.Add(elem);
            }
            else
            {
                TaskDialog.Show("Filter Tags",
                    "Please select only tag elements (IndependentTag or Area Tags). Command will now exit.");
                return Result.Cancelled;
            }
        }

        // First pass: determine the maximum number of tag text parameters
        int maxTagTextParams = 0;
        foreach (var tag in selectedTags)
        {
            string tagText = "";
            if (tag is IndependentTag independentTag)
            {
                tagText = independentTag.TagText;
            }
            else if (tag is SpatialElementTag spatialTag)
            {
                tagText = spatialTag.TagText;
            }

            // Clean and split the tag text
            if (!string.IsNullOrEmpty(tagText))
            {
                var cleanText = tagText
                    .Replace("\r\n", "\n")
                    .Replace("\r", "\n")
                    .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();

                maxTagTextParams = Math.Max(maxTagTextParams, cleanText.Count);
            }
        }

        // Build property names list with the determined number of TagText columns
        List<string> propertyNames = new List<string>();
        for (int i = 1; i <= maxTagTextParams; i++)
        {
            propertyNames.Add($"TagText{i}");
        }
        propertyNames.AddRange(new[]
        {
            "OwnerView",
            "SheetNumber",
            "SheetName",
            "ElementId"
        });

        // Build data to display
        List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();

        // Cache for view -> (sheetNumber, sheetName) lookups
        Dictionary<ElementId, (string viewName, string sheetNumber, string sheetName)> viewSheetInfoCache
            = new Dictionary<ElementId, (string viewName, string sheetNumber, string sheetName)>();

        foreach (var tag in selectedTags)
        {
            var dict = new Dictionary<string, object>();

            // Process tag text into separate columns
            string tagText = "";
            if (tag is IndependentTag independentTag)
            {
                tagText = independentTag.TagText;
            }
            else if (tag is SpatialElementTag spatialTag)
            {
                tagText = spatialTag.TagText;
            }

            // Clean and split the tag text
            var tagTextParams = new List<string>();
            if (!string.IsNullOrEmpty(tagText))
            {
                tagTextParams = tagText
                    .Replace("\r\n", "\n")
                    .Replace("\r", "\n")
                    .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();
            }

            // Add tag text parameters to separate columns
            for (int i = 1; i <= maxTagTextParams; i++)
            {
                dict[$"TagText{i}"] = i <= tagTextParams.Count ? tagTextParams[i - 1] : "";
            }

            // Get the owner view ID based on tag type
            ElementId ownerViewId;
            if (tag is IndependentTag independentTag2)
            {
                ownerViewId = independentTag2.OwnerViewId;
            }
            else // SpatialElementTag
            {
                ownerViewId = tag.OwnerViewId;
            }

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

            // ElementId
            dict["ElementId"] = tag.Id.IntegerValue;

            entries.Add(dict);
        }

        // Call the provided DataGrid method
        var filteredEntries = CustomGUIs.DataGrid(
            entries,
            propertyNames,
            spanAllScreens: false,
            initialSelectionIndices: null
        );

        // If user pressed ESC or closed the form, filteredEntries may be empty
        if (filteredEntries == null || !filteredEntries.Any())
        {
            return Result.Cancelled;
        }

        // Update the Revit selection with only the chosen elements
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
