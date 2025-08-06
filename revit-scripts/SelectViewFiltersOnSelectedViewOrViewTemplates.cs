using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

namespace MyRevitCommands
{
   /// <summary>
   /// Enhanced helper methods for filter commands with better information display
   /// </summary>
   public static class EnhancedFilterCommandHelper
   {
      /// <summary>
      /// Gets the active view or uses selected views, with fallback to active view's template
      /// </summary>
      public static List<View> GetViewsWithActiveViewFallback(UIDocument uiDoc, Document doc)
      {
         // First try to get selected views
         List<View> selectedViews = FilterCommandHelper.GetSelectedViews(uiDoc, doc);
         
         // If no views selected, use active view
         if (selectedViews.Count == 0)
         {
            View activeView = doc.ActiveView;
            if (activeView != null && !(activeView is ViewSchedule) && activeView.ViewType != ViewType.DrawingSheet)
            {
               selectedViews.Add(activeView);
            }
         }
         
         return selectedViews;
      }

      /// <summary>
      /// Builds enhanced grid entries with filter information and override details
      /// </summary>
      public static List<Dictionary<string, object>> BuildEnhancedFilterGridEntries(
         Document doc, 
         IEnumerable<ParameterFilterElement> filters,
         HashSet<ElementId> targetViewIds)
      {
         List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
         
         // Get first target view for override information
         View targetView = null;
         if (targetViewIds.Count > 0)
         {
            targetView = doc.GetElement(targetViewIds.First()) as View;
         }
         
         // Track maximum number of rules to know how many columns we need
         int maxRuleCount = 0;
         
         foreach (ParameterFilterElement filter in filters)
         {
            var entry = new Dictionary<string, object>();
            entry["Name"] = filter.Name;
            entry["Id"] = filter.Id.IntegerValue;
            
            // Get filter categories - show all of them
            ICollection<ElementId> categoryIds = filter.GetCategories();
            List<string> categoryNames = new List<string>();
            foreach (ElementId catId in categoryIds)
            {
               Category cat = Category.GetCategory(doc, catId);
               if (cat != null)
                  categoryNames.Add(cat.Name);
            }
            string categoriesStr = categoryNames.Count > 0 ? string.Join(", ", categoryNames) : "No categories";
            entry["Categories"] = categoriesStr;
            
            // Get filter rule information
            string ruleInfo = GetFilterRuleInfo(doc, filter);
            entry["Rule Info"] = ruleInfo;
            
            // Extract individual filter rules
            List<string> filterRules = ExtractFilterRules(doc, filter);
            for (int i = 0; i < filterRules.Count; i++)
            {
               entry[$"Filter Rule {i + 1}"] = filterRules[i];
            }
            maxRuleCount = Math.Max(maxRuleCount, filterRules.Count);
            
            // Initialize all override columns with default values
            entry["Enabled"] = "";     // Is filter enabled (checked)?
            entry["Visibility"] = "";  // Visibility setting
            entry["Halftone"] = "";
            entry["Proj Line Color"] = "";
            entry["Proj Weight"] = "";
            entry["Cut Line Color"] = "";
            entry["Cut Weight"] = "";
            entry["Transparency"] = "";
            entry["Surface Pattern"] = "";
            entry["Surface Color"] = "";
            entry["Cut Pattern"] = "";
            entry["Cut Color"] = "";
            
            // Get override information if filter is applied to target view
            if (targetView != null)
            {
               ICollection<ElementId> viewFilters = targetView.GetFilters();
               
               if (viewFilters.Contains(filter.Id))
               {
                  // Check if filter is enabled (this is the checkbox in the UI)
                  // Use GetIsFilterEnabled to check if filter is enabled
                  bool isEnabled = targetView.GetIsFilterEnabled(filter.Id);
                  entry["Enabled"] = isEnabled ? "Yes" : "No";
                  
                  // Check visibility using GetFilterVisibility
                  bool isVisible = targetView.GetFilterVisibility(filter.Id);
                  entry["Visibility"] = isVisible ? "Yes" : "No";
                  
                  // Get detailed overrides only if filter is enabled
                  if (isEnabled)
                  {
                     GetDetailedFilterOverrides(targetView, filter.Id, entry);
                  }
               }
            }
            
            entries.Add(entry);
         }
         
         // Ensure all entries have the same rule columns (even if empty)
         foreach (var entry in entries)
         {
            for (int i = 1; i <= maxRuleCount; i++)
            {
               if (!entry.ContainsKey($"Filter Rule {i}"))
               {
                  entry[$"Filter Rule {i}"] = "";
               }
            }
         }
         
         return entries;
      }

      /// <summary>
      /// Populates detailed override information in separate columns
      /// </summary>
      private static void GetDetailedFilterOverrides(View view, ElementId filterId, Dictionary<string, object> entry)
      {
         try
         {
            OverrideGraphicSettings ogs = view.GetFilterOverrides(filterId);
            
            // Halftone
            if (ogs.Halftone)
            {
               entry["Halftone"] = "Yes";
            }
            
            // Projection line color
            Color projLineColor = ogs.ProjectionLineColor;
            if (projLineColor != null && projLineColor.IsValid)
            {
               entry["Proj Line Color"] = projLineColor; // Store Color object for background coloring
            }
            
            // Cut line color
            Color cutLineColor = ogs.CutLineColor;
            if (cutLineColor != null && cutLineColor.IsValid)
            {
               entry["Cut Line Color"] = cutLineColor; // Store Color object for background coloring
            }
            
            // Transparency
            int transp = ogs.Transparency;
            if (transp > 0 && transp < 100)
            {
               entry["Transparency"] = $"{transp}%";
            }
            
            // Line weights
            int projWeight = ogs.ProjectionLineWeight;
            if (projWeight > 0 && projWeight != -1)
            {
               entry["Proj Weight"] = projWeight.ToString();
            }
            
            int cutWeight = ogs.CutLineWeight;
            if (cutWeight > 0 && cutWeight != -1)
            {
               entry["Cut Weight"] = cutWeight.ToString();
            }
            
            // Surface pattern and color
            try
            {
               ElementId projPatternId = ogs.SurfaceForegroundPatternId;
               if (projPatternId != null && projPatternId != ElementId.InvalidElementId)
               {
                  FillPatternElement pattern = view.Document.GetElement(projPatternId) as FillPatternElement;
                  if (pattern != null)
                  {
                     entry["Surface Pattern"] = pattern.Name;
                  }
               }
               
               Color surfaceColor = ogs.SurfaceForegroundPatternColor;
               if (surfaceColor != null && surfaceColor.IsValid)
               {
                  entry["Surface Color"] = surfaceColor; // Store Color object
               }
            }
            catch { }
            
            // Cut pattern and color
            try
            {
               ElementId cutPatternId = ogs.CutForegroundPatternId;
               if (cutPatternId != null && cutPatternId != ElementId.InvalidElementId)
               {
                  FillPatternElement pattern = view.Document.GetElement(cutPatternId) as FillPatternElement;
                  if (pattern != null)
                  {
                     entry["Cut Pattern"] = pattern.Name;
                  }
               }
               
               Color cutColor = ogs.CutForegroundPatternColor;
               if (cutColor != null && cutColor.IsValid)
               {
                  entry["Cut Color"] = cutColor; // Store Color object
               }
            }
            catch { }
         }
         catch (Exception ex)
         {
            // Log error but don't fail
         }
      }

      /// <summary>
      /// Gets filter rule information as a readable string
      /// </summary>
      private static string GetFilterRuleInfo(Document doc, ParameterFilterElement filter)
      {
         try
         {
            // Try to get the logical filter
            ElementFilter elementFilter = filter.GetElementFilter();
            if (elementFilter == null)
               return "No rules defined";
            
            // Build a description of the filter
            StringBuilder sb = new StringBuilder();
            
            // Try to extract some information from the filter
            // Note: The exact structure depends on Revit API version
            Type filterType = elementFilter.GetType();
            string filterTypeName = filterType.Name;
            
            // Provide basic type information (without category count)
            if (filterTypeName.Contains("And"))
               sb.Append("AND logic");
            else if (filterTypeName.Contains("Or"))
               sb.Append("OR logic");
            else
               sb.Append(filterTypeName.Replace("Filter", ""));
            
            return sb.ToString();
         }
         catch
         {
            return "Rule info unavailable";
         }
      }
      
      /// <summary>
      /// Extracts individual filter rules from a ParameterFilterElement
      /// </summary>
      private static List<string> ExtractFilterRules(Document doc, ParameterFilterElement filter)
      {
         List<string> rules = new List<string>();
         
         try
         {
            ElementFilter elementFilter = filter.GetElementFilter();
            if (elementFilter == null)
               return rules;
            
            // Extract rules from the ElementFilter hierarchy
            ExtractRulesFromElementFilter(doc, elementFilter, rules);
         }
         catch (Exception ex)
         {
            rules.Add($"Error: {ex.Message}");
         }
         
         // If no rules found, add a message
         if (rules.Count == 0)
         {
            rules.Add("No rules found");
         }
         
         return rules;
      }
      
      /// <summary>
      /// Recursively extracts rules from an ElementFilter
      /// </summary>
      private static void ExtractRulesFromElementFilter(Document doc, ElementFilter filter, List<string> rules)
      {
         if (filter == null) return;
         
         Type filterType = filter.GetType();
         string typeName = filterType.Name;
         
         // Check if it's a LogicalAndFilter or LogicalOrFilter
         if (filter is LogicalAndFilter andFilter)
         {
            // Get the filters it contains
            IList<ElementFilter> filters = andFilter.GetFilters();
            foreach (ElementFilter subFilter in filters)
            {
               ExtractRulesFromElementFilter(doc, subFilter, rules);
            }
         }
         else if (filter is LogicalOrFilter orFilter)
         {
            // Get the filters it contains
            IList<ElementFilter> filters = orFilter.GetFilters();
            foreach (ElementFilter subFilter in filters)
            {
               ExtractRulesFromElementFilter(doc, subFilter, rules);
            }
         }
         else if (filter is ElementParameterFilter paramFilter)
         {
            // This is where the actual rules are!
            try
            {
               IList<FilterRule> filterRules = paramFilter.GetRules();
               foreach (FilterRule rule in filterRules)
               {
                  string ruleString = ConvertFilterRuleToString(doc, rule);
                  if (!string.IsNullOrEmpty(ruleString))
                     rules.Add(ruleString);
               }
            }
            catch (Exception ex)
            {
               rules.Add($"Rule error: {ex.Message}");
            }
         }
         else
         {
            // Some other type of filter - try to get info about it
            rules.Add($"Filter type: {typeName}");
         }
      }
      
      /// <summary>
      /// Converts a FilterRule to a readable string
      /// </summary>
      private static string ConvertFilterRuleToString(Document doc, FilterRule rule)
      {
         if (rule == null)
            return "";
         
         try
         {
            StringBuilder sb = new StringBuilder();
            Type ruleType = rule.GetType();
            
            // Handle different types of filter rules
            if (rule is FilterStringRule stringRule)
            {
               // Get parameter
               ElementId paramId = stringRule.GetRuleParameter();
               string paramName = GetParameterName(doc, paramId);
               sb.Append(paramName);
               
               // Get evaluator (operator)
               FilterStringRuleEvaluator evaluator = stringRule.GetEvaluator();
               string evalType = evaluator.GetType().Name;
               
               if (evalType.Contains("Equals"))
                  sb.Append(" equals ");
               else if (evalType.Contains("Contains"))
                  sb.Append(" contains ");
               else if (evalType.Contains("BeginsWith"))
                  sb.Append(" begins with ");
               else if (evalType.Contains("EndsWith"))
                  sb.Append(" ends with ");
               else if (evalType.Contains("Greater"))
                  sb.Append(" greater than ");
               else if (evalType.Contains("Less"))
                  sb.Append(" less than ");
               else
                  sb.Append(" " + evalType + " ");
               
               // Get value
               string value = stringRule.RuleString;
               sb.Append($"\"{value}\"");
            }
            else if (rule is FilterDoubleRule doubleRule)
            {
               // Get parameter
               ElementId paramId = doubleRule.GetRuleParameter();
               string paramName = GetParameterName(doc, paramId);
               sb.Append(paramName);
               
               // Get evaluator
               FilterNumericRuleEvaluator evaluator = doubleRule.GetEvaluator();
               string evalType = evaluator.GetType().Name;
               
               if (evalType.Contains("Equals"))
                  sb.Append(" = ");
               else if (evalType.Contains("Greater"))
                  sb.Append(" > ");
               else if (evalType.Contains("Less"))
                  sb.Append(" < ");
               else if (evalType.Contains("GreaterOrEqual"))
                  sb.Append(" ≥ ");
               else if (evalType.Contains("LessOrEqual"))
                  sb.Append(" ≤ ");
               else
                  sb.Append(" " + evalType + " ");
               
               // Get value
               double value = doubleRule.RuleValue;
               // Try to format nicely (could convert units here if needed)
               sb.Append(value.ToString("0.###"));
            }
            else if (rule is FilterIntegerRule intRule)
            {
               // Get parameter
               ElementId paramId = intRule.GetRuleParameter();
               string paramName = GetParameterName(doc, paramId);
               sb.Append(paramName);
               
               // Get evaluator
               FilterNumericRuleEvaluator evaluator = intRule.GetEvaluator();
               string evalType = evaluator.GetType().Name;
               
               if (evalType.Contains("Equals"))
                  sb.Append(" = ");
               else if (evalType.Contains("Greater"))
                  sb.Append(" > ");
               else if (evalType.Contains("Less"))
                  sb.Append(" < ");
               else if (evalType.Contains("GreaterOrEqual"))
                  sb.Append(" ≥ ");
               else if (evalType.Contains("LessOrEqual"))
                  sb.Append(" ≤ ");
               else
                  sb.Append(" " + evalType + " ");
               
               // Get value
               int value = intRule.RuleValue;
               sb.Append(value);
            }
            else if (rule is FilterElementIdRule idRule)
            {
               // Get parameter
               ElementId paramId = idRule.GetRuleParameter();
               string paramName = GetParameterName(doc, paramId);
               sb.Append(paramName);
               
               // Get evaluator
               FilterNumericRuleEvaluator evaluator = idRule.GetEvaluator();
               string evalType = evaluator.GetType().Name;
               
               if (evalType.Contains("Equals"))
                  sb.Append(" equals ");
               else if (evalType.Contains("NotEquals"))
                  sb.Append(" not equals ");
               else if (evalType.Contains("Greater"))
                  sb.Append(" > ");
               else if (evalType.Contains("Less"))
                  sb.Append(" < ");
               else
                  sb.Append(" " + evalType + " ");
               
               // Get value
               ElementId value = idRule.RuleValue;
               Element elem = doc.GetElement(value);
               if (elem != null)
                  sb.Append(elem.Name);
               else if (value == ElementId.InvalidElementId)
                  sb.Append("<none>");
               else
                  sb.Append($"Id:{value.IntegerValue}");
            }
            else
            {
               // Unknown rule type - try generic approach
               sb.Append(rule.GetType().Name);
            }
            
            return sb.ToString();
         }
         catch (Exception ex)
         {
            return $"Error parsing rule: {ex.Message}";
         }
      }
      
      /// <summary>
      /// Gets the name of a parameter from its ElementId
      /// </summary>
      private static string GetParameterName(Document doc, ElementId paramId)
      {
         if (paramId == null)
            return "Unknown";
         
         // Check if it's a built-in parameter
         if (paramId.IntegerValue < 0)
         {
            try
            {
               BuiltInParameter bip = (BuiltInParameter)paramId.IntegerValue;
               return LabelUtils.GetLabelFor(bip);
            }
            catch
            {
               return $"BuiltIn({paramId.IntegerValue})";
            }
         }
         
         // Try to get shared parameter
         ParameterElement paramElem = doc.GetElement(paramId) as ParameterElement;
         if (paramElem != null)
            return paramElem.Name;
         
         return $"Param({paramId.IntegerValue})";
      }

      /// <summary>
      /// Custom DataGrid with color cell backgrounds
      /// </summary>
      public static List<Dictionary<string, object>> ColoredDataGrid(
         List<Dictionary<string, object>> entries,
         List<string> propertyNames)
      {
         // This would need to be implemented in CustomGUIs to support colored cells
         // For now, convert Color objects to simple RGB strings for display
         foreach (var entry in entries)
         {
            foreach (string propName in new[] { "Proj Line Color", "Cut Line Color", "Surface Color", "Cut Color" })
            {
               if (entry.ContainsKey(propName) && entry[propName] is Color color)
               {
                  entry[propName] = $"{color.Red},{color.Green},{color.Blue}";
               }
            }
         }
         
         return CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);
      }
   }

   /// <summary>
   /// Command to select view filters from selected views or view templates
   /// Falls back to active view if no selection
   /// </summary>
   [Transaction(TransactionMode.Manual)]
   public class SelectViewFiltersOnSelectedViewOrViewTemplates : IExternalCommand
   {
      public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
      {
         UIDocument uiDoc = commandData.Application.ActiveUIDocument;
         Document doc = uiDoc.Document;
         
         // Get views with fallback to active view
         List<View> selectedViews = EnhancedFilterCommandHelper.GetViewsWithActiveViewFallback(uiDoc, doc);
         if (selectedViews.Count == 0)
         {
            TaskDialog.Show("Error", "No valid view available (no selection and active view is not valid).");
            return Result.Failed;
         }
         
         // Get target view IDs (view templates if assigned, otherwise the views themselves)
         HashSet<ElementId> targetViewIds = FilterCommandHelper.GetTargetViewIds(selectedViews);
         
         // Get filters applied to the target views
         List<ParameterFilterElement> appliedFilters = FilterCommandHelper.GetAppliedFilterElements(doc, targetViewIds);
         if (appliedFilters.Count == 0)
         {
            TaskDialog.Show("Info", "No filters are applied to the selected view(s) or their view templates.");
            return Result.Cancelled;
         }
         
         // Build enhanced grid entries with override information
         List<Dictionary<string, object>> gridEntries = EnhancedFilterCommandHelper.BuildEnhancedFilterGridEntries(
            doc, appliedFilters, targetViewIds);
         
         // Determine how many rule columns we need
         int maxRuleColumns = 0;
         foreach (var entry in gridEntries)
         {
            int ruleCount = 0;
            for (int i = 1; i <= 20; i++) // Check up to 20 possible rule columns
            {
               if (entry.ContainsKey($"Filter Rule {i}") && !string.IsNullOrEmpty(entry[$"Filter Rule {i}"].ToString()))
                  ruleCount = i;
            }
            maxRuleColumns = Math.Max(maxRuleColumns, ruleCount);
         }
         
         // Build property names list dynamically based on number of rules
         List<string> propertyNames = new List<string> { 
            "Name", 
            "Id", 
            "Categories",
            "Enabled",
            "Visibility",
            "Halftone",
            "Proj Line Color",
            "Proj Weight",
            "Cut Line Color", 
            "Cut Weight",
            "Transparency",
            "Surface Pattern",
            "Surface Color",
            "Cut Pattern",
            "Cut Color",
            "Rule Info"
         };
         
         // Add rule columns at the end
         for (int i = 1; i <= maxRuleColumns; i++)
         {
            propertyNames.Add($"Filter Rule {i}");
         }
         
         // Show the data grid for selection
         List<Dictionary<string, object>> selectedEntries = EnhancedFilterCommandHelper.ColoredDataGrid(
            gridEntries, propertyNames);
         
         if (selectedEntries == null || selectedEntries.Count == 0)
         {
            return Result.Cancelled;
         }
         
         // Extract selected filter IDs
         List<ElementId> selectedFilterIds = new List<ElementId>();
         foreach (var entry in selectedEntries)
         {
            if (entry.TryGetValue("Id", out object idObj) && int.TryParse(idObj.ToString(), out int intId))
            {
               selectedFilterIds.Add(new ElementId(intId));
            }
         }
         
         // Add selected filters to current selection
         try
         {
            // Get current selection
            ICollection<ElementId> currentSelection = uiDoc.Selection.GetElementIds();
            
            // Create new selection set
            List<ElementId> newSelection = new List<ElementId>(currentSelection);
            
            // Add filter IDs that aren't already selected
            foreach (ElementId filterId in selectedFilterIds)
            {
               if (!newSelection.Contains(filterId))
               {
                  newSelection.Add(filterId);
               }
            }
            
            // Update selection
            uiDoc.Selection.SetElementIds(newSelection);
            
            return Result.Succeeded;
         }
         catch (Exception ex)
         {
            TaskDialog.Show("Error", $"Failed to update selection: {ex.Message}");
            return Result.Failed;
         }
      }
   }
}
