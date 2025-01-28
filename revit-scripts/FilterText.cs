using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class FilterText : IExternalCommand
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
            TaskDialog.Show("Filter Text", "No elements selected.");
            return Result.Cancelled;
        }

        // Filter for TextNote elements
        List<Element> selectedTexts = new List<Element>();
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem is TextNote)
            {
                selectedTexts.Add(elem);
            }
            else
            {
                TaskDialog.Show("Filter Text",
                    "Please select only text elements (TextNote). Command will now exit.");
                return Result.Cancelled;
            }
        }

        // First pass: determine the maximum number of text lines
        int maxTextLines = 0;
        foreach (var text in selectedTexts)
        {
            var textNote = text as TextNote;
            string textContent = textNote.Text;

            // Clean and split the text content
            if (!string.IsNullOrEmpty(textContent))
            {
                var cleanText = textContent
                    .Replace("\r\n", "\n")
                    .Replace("\r", "\n")
                    .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();

                maxTextLines = Math.Max(maxTextLines, cleanText.Count);
            }
        }

        // Build property names list with the determined number of Text columns
        List<string> propertyNames = new List<string>();
        for (int i = 1; i <= maxTextLines; i++)
        {
            propertyNames.Add($"Text{i}");
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

        foreach (var text in selectedTexts)
        {
            var dict = new Dictionary<string, object>();
            var textNote = text as TextNote;

            // Process text content into separate columns
            string textContent = textNote.Text;

            // Clean and split the text content
            var textLines = new List<string>();
            if (!string.IsNullOrEmpty(textContent))
            {
                textLines = textContent
                    .Replace("\r\n", "\n")
                    .Replace("\r", "\n")
                    .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();
            }

            // Add text lines to separate columns
            for (int i = 1; i <= maxTextLines; i++)
            {
                dict[$"Text{i}"] = i <= textLines.Count ? textLines[i - 1] : "";
            }

            // Get the owner view ID
            ElementId ownerViewId = text.OwnerViewId;
            
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
            dict["ElementId"] = text.Id.IntegerValue;

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

        using (var tx = new Transaction(doc, "Filter Text Selection"))
        {
            tx.Start();
            uidoc.Selection.SetElementIds(chosenIds);
            tx.Commit();
        }

        return Result.Succeeded;
    }
}
