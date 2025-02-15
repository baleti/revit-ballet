using System;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Text;

public class GetFilterRules : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get all ParameterFilterElement instances
        FilteredElementCollector collector = new FilteredElementCollector(doc)
            .OfClass(typeof(ParameterFilterElement));

        StringBuilder log = new StringBuilder();

        foreach (Element e in collector)
        {
            ParameterFilterElement filter = e as ParameterFilterElement;
            if (filter == null) continue;

            log.AppendLine($"Filter: {filter.Name}");

            // Get the ElementFilter from ParameterFilterElement
            ElementFilter elementFilter = filter.GetElementFilter();
            ProcessElementFilter(elementFilter, doc, log);
        }

        // Show results
        TaskDialog.Show("Filter Rules", log.ToString());

        return Result.Succeeded;
    }

    private void ProcessElementFilter(ElementFilter elementFilter, Document doc, StringBuilder log)
    {
        if (elementFilter is LogicalAndFilter logicalAndFilter)
        {
            // LogicalAndFilter contains multiple ElementFilters
            FieldInfo filtersField = typeof(LogicalAndFilter).GetField("m_filters", BindingFlags.NonPublic | BindingFlags.Instance);
            if (filtersField != null)
            {
                IList<ElementFilter> filters = filtersField.GetValue(logicalAndFilter) as IList<ElementFilter>;
                if (filters != null)
                {
                    foreach (ElementFilter subFilter in filters)
                    {
                        ProcessElementFilter(subFilter, doc, log);
                    }
                }
            }
        }
        else if (elementFilter is ElementParameterFilter paramFilter)
        {
            // Use reflection to get rules from ElementParameterFilter
            MethodInfo getRulesMethod = typeof(ElementParameterFilter).GetMethod("GetRules", BindingFlags.NonPublic | BindingFlags.Instance);
            if (getRulesMethod != null)
            {
                IList<FilterRule> rules = getRulesMethod.Invoke(paramFilter, null) as IList<FilterRule>;
                if (rules != null && rules.Count > 0)
                {
                    foreach (FilterRule rule in rules)
                    {
                        string ruleInfo = GetRuleDescription(rule, doc);
                        log.AppendLine($"  {ruleInfo}");
                    }
                }
                else
                {
                    log.AppendLine("  No rules found.");
                }
            }
            else
            {
                log.AppendLine("  Failed to access GetRules() method.");
            }
        }
    }

    private string GetRuleDescription(FilterRule rule, Document doc)
    {
        int paramId = GetParameterIdFromRule(rule);
        string paramName = GetParameterNameFromId(paramId, doc);

        if (rule is FilterStringRule stringRule)
        {
            return $"String Rule: Parameter='{paramName}', Rule={stringRule.ToString()}";
        }
        else if (rule is FilterDoubleRule doubleRule)
        {
            return $"Double Rule: Parameter='{paramName}', Rule={doubleRule.ToString()}";
        }
        else if (rule is FilterIntegerRule integerRule)
        {
            return $"Integer Rule: Parameter='{paramName}', Rule={integerRule.ToString()}";
        }
        else if (rule is FilterElementIdRule idRule)
        {
            return $"Element ID Rule: Parameter='{paramName}', Rule={idRule.ToString()}";
        }
        return $"Unknown rule type for Parameter='{paramName}'";
    }

    private int GetParameterIdFromRule(FilterRule rule)
    {
        // Extract the internal provider field
        FieldInfo providerField = rule.GetType().GetField("m_provider", BindingFlags.NonPublic | BindingFlags.Instance);
        if (providerField != null)
        {
            ParameterValueProvider provider = providerField.GetValue(rule) as ParameterValueProvider;
            if (provider != null)
            {
                // Extract the internal parameter ID field from ParameterValueProvider
                FieldInfo paramIdField = typeof(ParameterValueProvider).GetField("m_paramId", BindingFlags.NonPublic | BindingFlags.Instance);
                if (paramIdField != null)
                {
                    ElementId paramElementId = paramIdField.GetValue(provider) as ElementId;
                    if (paramElementId != null)
                    {
                        return paramElementId.IntegerValue;
                    }
                }
            }
        }
        return -1; // Unknown
    }

    private string GetParameterNameFromId(int paramId, Document doc)
    {
        if (paramId == -1)
            return "Unknown Parameter";

        try
        {
            BuiltInParameter bip = (BuiltInParameter)paramId;
            return LabelUtils.GetLabelFor(bip);
        }
        catch
        {
            return $"Unknown Parameter (ID: {paramId})";
        }
    }
}
