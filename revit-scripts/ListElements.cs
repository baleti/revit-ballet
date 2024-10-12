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
            // Check if there are any elements selected
            if (!elementIds.Any())
            {
                throw new InvalidOperationException("No elements are selected.");
            }
        }
        else
        {
            // Get all visible elements in the active view
            var collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            elementIds = collector.ToElementIds();
        }


        List<Dictionary<string, object>> elementData = new List<Dictionary<string, object>>();

        foreach (var id in elementIds)
        {
            Element element = doc.GetElement(id);
            ElementId groupId = element.GroupId;
            string groupName = "";
            if (groupId != null && groupId != ElementId.InvalidElementId && groupId.IntegerValue != -1)
            {
                Group group = doc.GetElement(groupId) as Group;
                groupName = group != null ? group.Name : "Unnamed Group";
            }

            ElementId ownerViewId = element.OwnerViewId;
            string ownerViewName = "";
            if (ownerViewId != null && ownerViewId != ElementId.InvalidElementId)
            {
                View view = doc.GetElement(ownerViewId) as View;
                ownerViewName = view != null ? view.Name : "Unnamed View";
            }

            var customElementData = new Dictionary<string, object>
            {
                ["Name"] = element.Name,
                ["Category"] = element.Category?.Name ?? "",
                ["Group"] = groupName,
                ["OwnerView"] = ownerViewName,
                ["Id"] = element.Id.IntegerValue,
            };

            if (includeParameters)
            {
                foreach (Parameter param in element.Parameters)
                {
                    string paramName = param.Definition.Name;
                    string paramValue = param.AsValueString() ?? param.AsString() ?? "None";
                    customElementData[paramName] = paramValue;
                }
            }

            elementData.Add(customElementData); // Move this line outside of the if block
        }
        return elementData;
    }
}

public abstract class ListElementsBase : IExternalCommand
{
    public abstract bool spanAllScreens { get; }
    public abstract bool useSelectedElements { get; }
    public abstract bool includeParameters { get; }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;

            List<Dictionary<string, object>> elementData = ElementDataHelper.GetElementData(uiDoc, useSelectedElements, includeParameters);
            var propertyNames = new HashSet<string>();

            if (elementData.Any())
            {
                foreach (var key in elementData.First().Keys)
                {
                    propertyNames.Add(key);
                }
            }

            var selectedCustomElements = CustomGUIs.DataGrid(elementData, propertyNames.ToList(), spanAllScreens);

            if (selectedCustomElements.Count == 0)
                return Result.Cancelled;

            List<ElementId> elementIdsToSelect = new List<ElementId>();
            foreach (var elementDict in selectedCustomElements)
            {
                if (elementDict.TryGetValue("Id", out object idValue) && idValue is int id)
                {
                    elementIdsToSelect.Add(new ElementId(id));
                }
            }

            uiDoc.Selection.SetElementIds(elementIdsToSelect);

            return Result.Succeeded;
        }
        catch (InvalidOperationException ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}

[Transaction(TransactionMode.Manual)]
public class ListSelectedElements : ListElementsBase
{
    public override bool spanAllScreens => false;
    public override bool useSelectedElements => true;
    public override bool includeParameters => false;
}
[Transaction(TransactionMode.Manual)]
public class ListSelectedElementsWithParameters : ListElementsBase
{
    public override bool spanAllScreens => false;
    public override bool useSelectedElements => true;
    public override bool includeParameters => true;
}

[Transaction(TransactionMode.Manual)]
public class ListSelectedElementsWithParametersSpanAllScreens : ListElementsBase
{
    public override bool spanAllScreens => true;
    public override bool useSelectedElements => true;
    public override bool includeParameters => true;
}
[Transaction(TransactionMode.Manual)]
public class ListAllElementsInView : ListElementsBase
{
    public override bool spanAllScreens => false;
    public override bool useSelectedElements => false;
    public override bool includeParameters => false;
}
[Transaction(TransactionMode.Manual)]
public class ListAllElementsInViewWithParameters : ListElementsBase
{
    public override bool spanAllScreens => false;
    public override bool useSelectedElements => false;
    public override bool includeParameters => true;
}
