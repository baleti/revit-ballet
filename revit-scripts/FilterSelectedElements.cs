using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public static class ElementDataHelper
{
    public static List<Dictionary<string, object>> GetElementData(UIDocument uiDoc, bool selectedOnly = false, bool includeParameters = false)
    {
        Document doc = uiDoc.Document;
        IEnumerable<ElementId> elementIds;
        if (selectedOnly)
        {
            elementIds = uiDoc.Selection.GetElementIds();
            if (!elementIds.Any())
                throw new InvalidOperationException("No elements are selected.");
        }
        else
        {
            elementIds = new FilteredElementCollector(doc, doc.ActiveView.Id).ToElementIds();
        }

        var elementData = new List<Dictionary<string, object>>();
        foreach (var id in elementIds)
        {
            Element element = doc.GetElement(id);

            string groupName = string.Empty;
            if (element.GroupId != null && element.GroupId != ElementId.InvalidElementId && element.GroupId.IntegerValue != -1)
            {
                if (doc.GetElement(element.GroupId) is Group g)
                    groupName = g.Name;
            }

            string ownerViewName = string.Empty;
            if (element.OwnerViewId != null && element.OwnerViewId != ElementId.InvalidElementId)
            {
                if (doc.GetElement(element.OwnerViewId) is View v)
                    ownerViewName = v.Name;
            }

            var data = new Dictionary<string, object>
            {
                ["Name"]      = element.Name,
                ["Category"]  = element.Category?.Name ?? string.Empty,
                ["Group"]     = groupName,
                ["OwnerView"] = ownerViewName,
                ["Id"]        = element.Id.IntegerValue,
            };

            if (includeParameters)
            {
                foreach (Parameter p in element.Parameters)
                {
                    string pName  = p.Definition.Name;
                    string pValue = p.AsValueString() ?? p.AsString() ?? "None";
                    data[pName] = pValue;
                }
            }

            elementData.Add(data);
        }
        return elementData;
    }
}

/// <summary>
/// Base class for commands that display Revit elements in a custom data‑grid for filtering and re‑selection.
/// </summary>
public abstract class FilterElementsBase : IExternalCommand
{
    public abstract bool SpanAllScreens      { get; }
    public abstract bool UseSelectedElements { get; }
    public abstract bool IncludeParameters   { get; }

    public Result Execute(ExternalCommandData cData, ref string message, ElementSet elements)
    {
        try
        {
            var uiDoc = cData.Application.ActiveUIDocument;
            var elementData = ElementDataHelper.GetElementData(uiDoc, UseSelectedElements, IncludeParameters);

            var propertyNames = elementData.Any()
                ? elementData.First().Keys.ToHashSet()
                : new HashSet<string>();

            var chosenRows = CustomGUIs.DataGrid(elementData, propertyNames.ToList(), SpanAllScreens);
            if (chosenRows.Count == 0)
                return Result.Cancelled;

            var ids = chosenRows
                .Where(d => d.TryGetValue("Id", out var v) && v is int)
                .Select(d => new ElementId((int)d["Id"]))
                .ToList();

            uiDoc.Selection.SetElementIds(ids);
            return Result.Succeeded;
        }
        catch (InvalidOperationException ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}

#region Concrete commands

[Transaction(TransactionMode.Manual)]
public class FilterSelectedElements : FilterElementsBase
{
    public override bool SpanAllScreens      => false;
    public override bool UseSelectedElements => true;
    public override bool IncludeParameters   => true;
}

[Transaction(TransactionMode.Manual)]
public class FilterSelectedElementsSpanAllScreens : FilterElementsBase
{
    public override bool SpanAllScreens      => true;
    public override bool UseSelectedElements => true;
    public override bool IncludeParameters   => true;
}

#endregion
