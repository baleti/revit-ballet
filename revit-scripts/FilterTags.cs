using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture; // Required for RoomTag and Room
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

        // Get currently selected elements.
        ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
        if (!selectedIds.Any())
        {
            TaskDialog.Show("Filter Tags", "No elements selected.");
            return Result.Cancelled;
        }

        // Filter only tag elements: IndependentTag and SpatialElementTag (which includes RoomTag).
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
                    "Please select only tag elements (IndependentTag or RoomTag). Command will now exit.");
                return Result.Cancelled;
            }
        }

        // Determine the maximum number of tag text lines (split by line breaks).
        int maxTagTextParams = 0;
        foreach (var tag in selectedTags)
        {
            string tagText = "";
            if (tag is IndependentTag indTag)
            {
                tagText = indTag.TagText;
            }
            else if (tag is SpatialElementTag spatTag)
            {
                tagText = spatTag.TagText;
            }
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

        // Build property names list.
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
            "TaggedFamilyAndType",
            "X",   // X coordinate
            "Y",   // Y coordinate
            "Z",   // Z coordinate
            "ElementId"
        });

        // Build data entries.
        List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();

        // Cache for view -> (sheetNumber, sheetName) lookups.
        Dictionary<ElementId, (string viewName, string sheetNumber, string sheetName)> viewSheetInfoCache =
            new Dictionary<ElementId, (string viewName, string sheetNumber, string sheetName)>();

        foreach (var tag in selectedTags)
        {
            var dict = new Dictionary<string, object>();

            // Process tag text into separate columns.
            string tagText = "";
            if (tag is IndependentTag indTag)
            {
                tagText = indTag.TagText;
            }
            else if (tag is SpatialElementTag spatTag)
            {
                tagText = spatTag.TagText;
            }
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
            for (int i = 1; i <= maxTagTextParams; i++)
            {
                dict[$"TagText{i}"] = i <= tagTextParams.Count ? tagTextParams[i - 1] : "";
            }

            // Get the tag's location.
            XYZ location = null;
            if (tag is IndependentTag indTag2)
            {
                location = indTag2.TagHeadPosition;
            }
            else if (tag is SpatialElementTag spatTag2)
            {
                LocationPoint locPoint = tag.Location as LocationPoint;
                if (locPoint != null)
                {
                    location = locPoint.Point;
                }
            }
            if (location != null)
            {
                dict["X"] = Math.Round(location.X, 2);
                dict["Y"] = Math.Round(location.Y, 2);
                dict["Z"] = Math.Round(location.Z, 2);
            }
            else
            {
                dict["X"] = "N/A";
                dict["Y"] = "N/A";
                dict["Z"] = "N/A";
            }

            // Get the owner view of the tag.
            ElementId ownerViewId = (tag is IndependentTag indTag3)
                ? indTag3.OwnerViewId
                : tag.OwnerViewId;
            string viewName = "";
            string sheetNumber = "";
            string sheetName = "";
            if (!viewSheetInfoCache.ContainsKey(ownerViewId))
            {
                var view = doc.GetElement(ownerViewId) as Autodesk.Revit.DB.View;
                if (view != null)
                {
                    viewName = view.Name;
                    // Attempt to get sheet info via a Viewport.
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
            var cached = viewSheetInfoCache[ownerViewId];
            dict["OwnerView"] = cached.viewName;
            dict["SheetNumber"] = cached.sheetNumber;
            dict["SheetName"] = cached.sheetName;

            // --- Retrieve the tagged element ---
            Element taggedElement = null;
            if (tag is IndependentTag independentTag4)
            {
                var taggedElementIds = independentTag4.GetTaggedElementIds();
                if (taggedElementIds.Any())
                {
                    LinkElementId firstId = taggedElementIds.First();
                    ElementId elementId = (firstId.LinkedElementId != ElementId.InvalidElementId)
                        ? firstId.LinkedElementId
                        : firstId.HostElementId;
                    taggedElement = doc.GetElement(elementId);
                }
            }
            else if (tag is RoomTag roomTag)
            {
                taggedElement = roomTag.Room;
            }

            // Format the tagged element information for display.
            string taggedFamilyAndType = "N/A";
            if (taggedElement != null)
            {
                if (taggedElement is FamilyInstance fi)
                {
                    var symbol = fi.Symbol;
                    if (symbol != null)
                    {
                        taggedFamilyAndType = $"{symbol.FamilyName} : {symbol.Name}";
                    }
                }
                else if (taggedElement is Room room)
                {
                    taggedFamilyAndType = $"Room {room.Number} : {room.Name}";
                }
                else
                {
                    string categoryName = taggedElement.Category != null ? taggedElement.Category.Name : "N/A";
                    taggedFamilyAndType = $"{categoryName} : {taggedElement.Name}";
                }
            }
            dict["TaggedFamilyAndType"] = taggedFamilyAndType;

            // Save the tag's own ElementId (for later updating the selection).
            dict["ElementId"] = tag.Id.IntegerValue;
            entries.Add(dict);
        }

        // Call the DataGrid display (assumed to be part of your CustomGUIs library).
        var filteredEntries = CustomGUIs.DataGrid(
            entries,
            propertyNames,
            spanAllScreens: false,
            initialSelectionIndices: null
        );
        if (filteredEntries == null || !filteredEntries.Any())
        {
            return Result.Cancelled;
        }

        // Update the selection with the chosen tags.
        List<ElementId> chosenIds = new List<ElementId>();
        foreach (var row in filteredEntries)
        {
            if (row.ContainsKey("ElementId") &&
                int.TryParse(row["ElementId"].ToString(), out int intId))
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
