#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace YourAddinNamespace
{
    /// <summary>
    /// Presents the 50 elements with the highest ElementId values
    /// (i.e. assumed “most recently created”), lets the user pick any
    /// subset in a <c>CustomGUIs.DataGrid</c>, and ADDS the chosen
    /// elements to the current Revit selection.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class SelectLast300Created : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uiDoc.Document;

            // 1. Grab every non-type element in the model.
            IEnumerable<ElementId> allIds = new FilteredElementCollector(doc)
                                            .WhereElementIsNotElementType()
                                            .ToElementIds();

            if (!allIds.Any())
            {
                message = "The project contains no model elements.";
                return Result.Failed;
            }

            // 2. Take the 50 highest Ids (or fewer, if <50 exist).
            List<ElementId> recentIds = allIds
                                        .OrderByDescending(id => id.AsLong())
                                        .Take(5000)
                                        .ToList();

            // 3. Build the rows for the DataGrid.
            var elementData = new List<Dictionary<string, object>>();

            foreach (ElementId id in recentIds)
            {
                Element el = doc.GetElement(id);
                if (el == null) continue;

                string groupName = string.Empty;
                if (el.GroupId != ElementId.InvalidElementId &&
                    el.GroupId.AsLong() != -1 &&
                    doc.GetElement(el.GroupId) is Group g)
                {
                    groupName = g.Name;
                }

                string ownerViewName = string.Empty;
                if (el.OwnerViewId != ElementId.InvalidElementId &&
                    doc.GetElement(el.OwnerViewId) is View v)
                {
                    ownerViewName = v.Name;
                }

                elementData.Add(new Dictionary<string, object>
                {
                    ["Name"]      = el.Name,
                    ["Category"]  = el.Category?.Name ?? string.Empty,
                    ["Group"]     = groupName,
                    ["OwnerView"] = ownerViewName,
                    ["Id"]        = el.Id.AsLong()
                });
            }

            if (elementData.Count == 0)
            {
                message = "Couldn’t collect any elements for the grid.";
                return Result.Failed;
            }

            // 4. Show the grid (single-monitor width assumed here).
            var propertyNames = elementData.First().Keys.ToList();
            var chosenRows    = CustomGUIs.DataGrid(
                                    elementData,
                                    propertyNames,
                                    spanAllScreens: false);

            // User pressed “Cancel” or unchecked every row.
            if (chosenRows.Count == 0)
                return Result.Cancelled;

            // 5. Convert rows back to ElementIds.
            HashSet<ElementId> chosenIds = chosenRows
                .Where(r => r.TryGetValue("Id", out var v) && v is int)
                .Select(r => ((long)r["Id"]).ToElementId())
                .ToHashSet();

            // 6. Add to the user’s existing selection (don’t replace).
            ICollection<ElementId> currentSelection = uiDoc.GetSelectionIds();
            uiDoc.SetSelectionIds(
                currentSelection.Union(chosenIds).ToList());

            return Result.Succeeded;
        }
    }
}

#endif
