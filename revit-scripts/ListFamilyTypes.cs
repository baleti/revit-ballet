using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public static class FamilyTypeParameterHelper
{
    public static (List<Dictionary<string, object>> data, List<string> properties) GetFamilyTypeParameterData(UIDocument uiDoc, bool selectedOnly = false)
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

        List<Dictionary<string, object>> elementData = new List<Dictionary<string, object>>();
        HashSet<string> propertyNames = new HashSet<string>
        {
            "Element Id",
            "Family Name",
            "Type Name",
            "Category",
            "Type Id"
        };

        foreach (var id in elementIds)
        {
            Element element = doc.GetElement(id);
            FamilySymbol familySymbol = element is FamilyInstance familyInstance ? familyInstance.Symbol : null;

            if (familySymbol != null)
            {
                var parameterDict = new Dictionary<string, object>
                {
                    ["Element Id"] = element.Id.IntegerValue,
                    ["Family Name"] = familySymbol.FamilyName,
                    ["Type Name"] = familySymbol.Name,
                    ["Category"] = familySymbol.Category?.Name ?? "",
                    ["Type Id"] = familySymbol.Id.IntegerValue
                };

                // Add all family type parameters
                foreach (Parameter param in familySymbol.Parameters)
                {
                    string paramName = param.Definition.Name;
                    string paramValue = param.AsValueString() ?? param.AsString() ?? "None";
                    parameterDict[paramName] = paramValue;
                    propertyNames.Add(paramName);
                }

                // Add all instance parameters
                if (element is FamilyInstance instance)
                {
                    foreach (Parameter param in instance.Parameters)
                    {
                        string paramName = $"Instance: {param.Definition.Name}";
                        string paramValue = param.AsValueString() ?? param.AsString() ?? "None";
                        parameterDict[paramName] = paramValue;
                        propertyNames.Add(paramName);
                    }
                }

                elementData.Add(parameterDict);
            }
        }
        
        return (elementData, propertyNames.OrderBy(p => p).ToList());
    }
}

public abstract class ListElementsFamilyTypeParametersBase : IExternalCommand
{
    public abstract bool SpanAllScreens { get; }
    public abstract bool UseSelectedElements { get; }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;

            // Get the family type data and property names
            var (elementData, propertyNames) = 
                FamilyTypeParameterHelper.GetFamilyTypeParameterData(uiDoc, UseSelectedElements);

            if (!elementData.Any())
            {
                message = UseSelectedElements ? 
                    "No family instances found in selection." : 
                    "No family instances found in active view.";
                return Result.Failed;
            }

            // Use the CustomGUIs.DataGrid implementation
            var selectedElements = CustomGUIs.DataGrid(
                entries: elementData,
                propertyNames: propertyNames,
                spanAllScreens: SpanAllScreens,
                initialSelectionIndices: null
            );

            if (selectedElements?.Count == 0)
                return Result.Cancelled;

            // Select the elements in Revit
            List<ElementId> elementIdsToSelect = new List<ElementId>();
            foreach (var elementDict in selectedElements)
            {
                if (elementDict.TryGetValue("Element Id", out object idValue) && idValue is int id)
                {
                    elementIdsToSelect.Add(new ElementId(id));
                }
            }

            uiDoc.Selection.SetElementIds(elementIdsToSelect);
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}

[Transaction(TransactionMode.Manual)]
public class ListSelectedElementsFamilyTypeParameters : ListElementsFamilyTypeParametersBase
{
    public override bool SpanAllScreens => false;
    public override bool UseSelectedElements => true;
}

[Transaction(TransactionMode.Manual)]
public class ListAllElementsFamilyTypeParameters : ListElementsFamilyTypeParametersBase
{
    public override bool SpanAllScreens => false;
    public override bool UseSelectedElements => false;
}

[Transaction(TransactionMode.Manual)]
public class ListAllElementsFamilyTypeParametersSpanScreens : ListElementsFamilyTypeParametersBase
{
    public override bool SpanAllScreens => true;
    public override bool UseSelectedElements => false;
}
