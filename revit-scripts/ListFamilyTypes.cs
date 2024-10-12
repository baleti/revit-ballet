using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public static class FamilyTypeDataHelper
{
    public static List<Dictionary<string, object>> GetFamilyTypeData(UIDocument uiDoc, bool selectedOnly = false)
    {
        Document doc = uiDoc.Document;
        IEnumerable<ElementId> elementIds;
        if (selectedOnly)
        {
            elementIds = uiDoc.Selection.GetElementIds();
            if (!elementIds.Any())
            {
                throw new InvalidOperationException("No elements are selected.");
            }
        }
        else
        {
            var collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            elementIds = collector.ToElementIds();
        }

        List<Dictionary<string, object>> familyTypeData = new List<Dictionary<string, object>>();

        foreach (var id in elementIds)
        {
            Element element = doc.GetElement(id);
            FamilySymbol familySymbol = element is FamilyInstance familyInstance ? familyInstance.Symbol : null;

            if (familySymbol != null)
            {
                var familyTypeDict = new Dictionary<string, object>
                {
                    ["FamilyName"] = familySymbol.FamilyName,
                    ["TypeName"] = familySymbol.Name,
                    ["Category"] = familySymbol.Category?.Name ?? "",
                    ["Id"] = familySymbol.Id.IntegerValue,
                };

                foreach (Parameter param in familySymbol.Parameters)
                {
                    string paramName = param.Definition.Name;
                    string paramValue = param.AsValueString() ?? param.AsString() ?? "None";
                    familyTypeDict[paramName] = paramValue;
                }

                familyTypeData.Add(familyTypeDict);
            }
        }
        return familyTypeData;
    }
}

public abstract class ListFamilyTypesBase : IExternalCommand
{
    public abstract bool spanAllScreens { get; }
    public abstract bool useSelectedElements { get; }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;

            List<Dictionary<string, object>> familyTypeData = FamilyTypeDataHelper.GetFamilyTypeData(uiDoc, useSelectedElements);
            var propertyNames = new HashSet<string>();

            if (familyTypeData.Any())
            {
                foreach (var key in familyTypeData.First().Keys)
                {
                    propertyNames.Add(key);
                }
            }

            var selectedCustomElements = CustomGUIs.DataGrid(familyTypeData, propertyNames.ToList(), spanAllScreens);

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
public class ListSelectedFamilyTypes : ListFamilyTypesBase
{
    public override bool spanAllScreens => false;
    public override bool useSelectedElements => true;
}

[Transaction(TransactionMode.Manual)]
public class ListAllFamilyTypesInView : ListFamilyTypesBase
{
    public override bool spanAllScreens => false;
    public override bool useSelectedElements => false;
}
[Transaction(TransactionMode.Manual)]
public class ListAllFamilyTypesSpanAllScreens : ListFamilyTypesBase
{
    public override bool spanAllScreens => true;
    public override bool useSelectedElements => true;
}
