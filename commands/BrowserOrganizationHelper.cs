using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitBallet.Commands
{
    /// <summary>
    /// Helper class for working with Revit Project Browser organization settings.
    /// Extracts grouping and sorting parameters from browser organization to use as DataGrid columns.
    /// </summary>
    public static class BrowserOrganizationHelper
    {
        /// <summary>
        /// Represents a browser organization column with its name and value extractor.
        /// </summary>
        public class BrowserColumn
        {
            public string Name { get; set; }
            public Func<View, Document, string> GetValue { get; set; }
        }

        /// <summary>
        /// Gets browser organization columns for BOTH sheets and views.
        /// Returns columns in the order: sheet browser org columns first, then view browser org columns.
        /// This allows proper sorting that matches the Revit project browser hierarchy.
        /// </summary>
        /// <param name="doc">The Revit document</param>
        /// <param name="sampleElements">Optional sample elements (not used, kept for compatibility)</param>
        /// <returns>List of browser columns with value extractors</returns>
        public static List<BrowserColumn> GetBrowserColumnsForViews(Document doc, IEnumerable<View> sampleElements = null)
        {
            List<BrowserColumn> columns = new List<BrowserColumn>();
            List<string> diagnosticLines = new List<string>();

            try
            {
                diagnosticLines.Add("=== GetBrowserColumnsForViews Detailed Diagnostic ===");
                diagnosticLines.Add($"Document: {doc.Title}");
                diagnosticLines.Add($"Time: {System.DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
                diagnosticLines.Add("\nGetting BOTH sheet and view browser organizations");

                // ============================================
                // PART 1: Get Sheet Browser Organization
                // ============================================
                diagnosticLines.Add("\n--- SHEET BROWSER ORGANIZATION ---");
                BrowserOrganization sheetOrg = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc);
                diagnosticLines.Add($"Sheet BrowserOrganization retrieved: {sheetOrg != null}");

                if (sheetOrg != null)
                {
                    diagnosticLines.Add($"Sheet Browser Organization Name: {sheetOrg.Name}");

                    // Get a sample sheet
                    ViewSheet sampleSheet = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .FirstOrDefault();

                    if (sampleSheet != null)
                    {
                        diagnosticLines.Add($"Sample Sheet: {sampleSheet.SheetNumber} - {sampleSheet.Name}");

                        IList<FolderItemInfo> sheetFolderItems = sheetOrg.GetFolderItems(sampleSheet.Id);
                        diagnosticLines.Add($"Sheet Folder Items: {sheetFolderItems?.Count ?? 0}");

                        if (sheetFolderItems != null && sheetFolderItems.Count > 0)
                        {
                            ProcessFolderItems(doc, sheetFolderItems, columns, diagnosticLines, "SHEET");
                        }
                    }
                    else
                    {
                        diagnosticLines.Add("No sheets found in document");
                    }
                }

                // ============================================
                // PART 2: Get View Browser Organization
                // ============================================
                diagnosticLines.Add("\n--- VIEW BROWSER ORGANIZATION ---");
                BrowserOrganization viewOrg = BrowserOrganization.GetCurrentBrowserOrganizationForViews(doc);
                diagnosticLines.Add($"View BrowserOrganization retrieved: {viewOrg != null}");

                if (viewOrg != null)
                {
                    diagnosticLines.Add($"View Browser Organization Name: {viewOrg.Name}");

                    // Get a sample view (non-sheet)
                    View sampleView = new FilteredElementCollector(doc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .FirstOrDefault(v => !v.IsTemplate && !(v is ViewSheet));

                    if (sampleView != null)
                    {
                        diagnosticLines.Add($"Sample View: {sampleView.Name}");

                        IList<FolderItemInfo> viewFolderItems = viewOrg.GetFolderItems(sampleView.Id);
                        diagnosticLines.Add($"View Folder Items: {viewFolderItems?.Count ?? 0}");

                        if (viewFolderItems != null && viewFolderItems.Count > 0)
                        {
                            ProcessFolderItems(doc, viewFolderItems, columns, diagnosticLines, "VIEW");
                        }
                    }
                    else
                    {
                        diagnosticLines.Add("No non-sheet views found in document");
                    }
                }

                diagnosticLines.Add($"\n=== FINAL COLUMN COUNT: {columns.Count} ===");
                // WriteDiagnostic(diagnosticLines); // Disabled diagnostic logging
            }
            catch (Exception ex)
            {
                // If browser organization isn't available or fails, return empty list
                diagnosticLines.Add($"\nEXCEPTION: {ex.GetType().Name}: {ex.Message}");
                diagnosticLines.Add($"Stack trace: {ex.StackTrace}");
                // WriteDiagnostic(diagnosticLines); // Disabled diagnostic logging
                return columns;
            }

            return columns;
        }

        /// <summary>
        /// Processes folder items from a browser organization and adds them to the columns list.
        /// </summary>
        private static void ProcessFolderItems(
            Document doc,
            IList<FolderItemInfo> folderItems,
            List<BrowserColumn> columns,
            List<string> diagnosticLines,
            string prefix)
        {
            diagnosticLines.Add($"\nProcessing {prefix} folder items:");
            int itemIndex = 0;
            foreach (FolderItemInfo folderItem in folderItems)
            {
                diagnosticLines.Add($"\n  {prefix} FolderItem [{itemIndex}]:");
                diagnosticLines.Add($"    Folder Name: '{folderItem.Name}'");

#if REVIT2026
                    // Use new API for Revit 2026+
                    IList<ElementId> parameterIdPath = folderItem.GetGroupingParameterIdPath();
                    diagnosticLines.Add($"    Parameter ID Path Count: {parameterIdPath?.Count ?? 0}");

                    if (parameterIdPath == null || parameterIdPath.Count == 0)
                    {
                        diagnosticLines.Add($"    SKIPPED: No parameter ID path");
                        itemIndex++;
                        continue;
                    }

                    // Use the first (and typically only) parameter in the path
                    ElementId parameterId = parameterIdPath[0];
                    diagnosticLines.Add($"    Parameter ID from path: {parameterId}");
#else
                    // Use deprecated API for earlier versions
                    ElementId parameterId = folderItem.ElementId;
                    diagnosticLines.Add($"    Parameter ID (ElementId): {parameterId}");
#endif

                    if (parameterId == null || parameterId == ElementId.InvalidElementId)
                    {
                        diagnosticLines.Add($"    SKIPPED: Invalid parameter ID");
                        itemIndex++;
                        continue;
                    }

                    // Check if it's a ParameterElement
                    ParameterElement paramElem = doc.GetElement(parameterId) as ParameterElement;
                    if (paramElem != null)
                    {
                        diagnosticLines.Add($"    Found ParameterElement: '{paramElem.Name}'");
                    }
                    else
                    {
#if REVIT2024 || REVIT2025 || REVIT2026
                        long paramIdValue = parameterId.Value;
#else
                        int paramIdInt = parameterId.IntegerValue;
                        long paramIdValue = paramIdInt;
#endif
                        // Check if it's a valid BuiltInParameter by trying to cast
                        try
                        {
                            BuiltInParameter builtInParam = (BuiltInParameter)paramIdValue;
                            if (System.Enum.IsDefined(typeof(BuiltInParameter), builtInParam))
                            {
                                diagnosticLines.Add($"    Found BuiltInParameter: {builtInParam}");
                            }
                            else
                            {
                                diagnosticLines.Add($"    Parameter ID {parameterId} ({paramIdValue}) is not a valid BuiltInParameter enum value");
                            }
                        }
                        catch (Exception ex)
                        {
                            diagnosticLines.Add($"    Parameter ID {parameterId} ({paramIdValue}) is neither ParameterElement nor BuiltInParameter: {ex.Message}");
                        }
                    }

                    // Try to get parameter name and create value extractor
                    BrowserColumn column = CreateColumnForParameter(doc, parameterId);

                    if (column != null)
                    {
                        diagnosticLines.Add($"    ✓ SUCCESS: Created column '{column.Name}'");
                        columns.Add(column);
                    }
                    else
                    {
                        diagnosticLines.Add($"    ✗ FAILED: Could not create column for parameter ID {parameterId}");
                    }

                itemIndex++;
            }
        }

        /// <summary>
        /// Writes diagnostic information to a file in runtime/diagnostics/
        /// </summary>
        private static void WriteDiagnostic(List<string> diagnosticLines)
        {
            try
            {
                string runtimeDir = System.IO.Path.Combine(
                    System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
                    "revit-ballet",
                    "runtime");

                string diagnosticPath = System.IO.Path.Combine(
                    runtimeDir,
                    "diagnostics",
                    $"BrowserOrgHelper-{System.DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");

                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
                System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);
            }
            catch
            {
                // Silently fail if diagnostic writing fails
            }
        }

        /// <summary>
        /// Creates a browser column for a given parameter ID.
        /// </summary>
        private static BrowserColumn CreateColumnForParameter(Document doc, ElementId parameterId)
        {
            try
            {
                // Try to get as ParameterElement (shared or project parameter)
                ParameterElement paramElem = doc.GetElement(parameterId) as ParameterElement;

                if (paramElem != null)
                {
                    // Shared or project parameter
                    string paramName = paramElem.Name;

                    return new BrowserColumn
                    {
                        Name = paramName,
                        GetValue = (view, d) => GetParameterValue(view, parameterId, paramName)
                    };
                }
                else
                {
                    // Might be a built-in parameter
#if REVIT2024 || REVIT2025 || REVIT2026
                    long paramIdValue = parameterId.Value;
#else
                    int paramIdInt = parameterId.IntegerValue;
                    long paramIdValue = paramIdInt;
#endif

                    // Try to cast to BuiltInParameter
                    try
                    {
                        BuiltInParameter builtInParam = (BuiltInParameter)paramIdValue;

                        // Check if it's a valid built-in parameter
                        if (Enum.IsDefined(typeof(BuiltInParameter), builtInParam))
                        {
                            string paramName = GetBuiltInParameterDisplayName(builtInParam);

                            return new BrowserColumn
                            {
                                Name = paramName,
                                GetValue = (view, d) => GetBuiltInParameterValue(view, builtInParam)
                            };
                        }
                    }
                    catch
                    {
                        // Not a valid BuiltInParameter
                    }
                }
            }
            catch (Exception)
            {
                // Parameter not accessible, skip it
            }

            return null;
        }

        /// <summary>
        /// Gets a parameter value from a view by parameter ID.
        /// </summary>
        private static string GetParameterValue(View view, ElementId parameterId, string fallbackName)
        {
            try
            {
                // Find parameter by iterating through all parameters and matching ID
                foreach (Parameter param in view.Parameters)
                {
                    if (param.Id == parameterId)
                    {
                        return GetParameterValueAsString(param);
                    }
                }
            }
            catch (Exception)
            {
                // Parameter not available on this view
            }

            return "";
        }

        /// <summary>
        /// Gets a built-in parameter value from a view.
        /// </summary>
        private static string GetBuiltInParameterValue(View view, BuiltInParameter builtInParam)
        {
            try
            {
                Parameter param = view.get_Parameter(builtInParam);

                if (param != null)
                {
                    return GetParameterValueAsString(param);
                }
            }
            catch (Exception)
            {
                // Parameter not available on this view
            }

            return "";
        }

        /// <summary>
        /// Converts a parameter value to a string.
        /// </summary>
        private static string GetParameterValueAsString(Parameter param)
        {
            if (param == null || !param.HasValue)
                return "";

            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString() ?? "";

                case StorageType.Integer:
                    return param.AsInteger().ToString();

                case StorageType.Double:
                    return param.AsValueString() ?? param.AsDouble().ToString();

                case StorageType.ElementId:
                    ElementId elemId = param.AsElementId();
                    if (elemId != null && elemId != ElementId.InvalidElementId)
                    {
                        // Try to get element name
                        Element elem = param.Element.Document.GetElement(elemId);
                        return elem != null ? elem.Name : elemId.ToString();
                    }
                    return "";

                default:
                    return "";
            }
        }

        /// <summary>
        /// Gets a user-friendly display name for a built-in parameter.
        /// </summary>
        private static string GetBuiltInParameterDisplayName(BuiltInParameter builtInParam)
        {
            // Map common built-in parameters to friendly names
            switch (builtInParam)
            {
                case BuiltInParameter.VIEW_TYPE:
                    return "View Type";
                case BuiltInParameter.VIEW_FAMILY:
                    return "Family";
                case BuiltInParameter.VIEW_DISCIPLINE:
                    return "Discipline";
                case BuiltInParameter.ELEM_FAMILY_PARAM:
                    return "Family";
                case BuiltInParameter.ELEM_TYPE_PARAM:
                    return "Type";
                case BuiltInParameter.VIEW_PHASE:
                    return "Phase";
                case BuiltInParameter.VIEW_DETAIL_LEVEL:
                    return "Detail Level";
                default:
                    // Use the enum name as fallback
                    return builtInParam.ToString().Replace("_", " ");
            }
        }

        /// <summary>
        /// Adds browser organization columns to a view dictionary.
        /// </summary>
        /// <param name="viewDict">The dictionary to populate</param>
        /// <param name="view">The view to extract values from</param>
        /// <param name="doc">The document</param>
        /// <param name="browserColumns">The browser columns to add</param>
        public static void AddBrowserColumnsToDict(
            Dictionary<string, object> viewDict,
            View view,
            Document doc,
            List<BrowserColumn> browserColumns)
        {
            foreach (var column in browserColumns)
            {
                string value = column.GetValue(view, doc);
                viewDict[column.Name] = value;
            }
        }

        /// <summary>
        /// Sorts view data by browser organization columns in order.
        /// Uses natural sorting to handle numeric prefixes correctly (e.g., "100 -" before "500 -").
        /// Falls back to sorting by "Name" column when browser columns are empty or identical.
        /// </summary>
        /// <param name="viewData">The list of view dictionaries to sort</param>
        /// <param name="browserColumns">The browser columns to sort by</param>
        /// <returns>Sorted list</returns>
        public static List<Dictionary<string, object>> SortByBrowserColumns(
            List<Dictionary<string, object>> viewData,
            List<BrowserColumn> browserColumns)
        {
            if (browserColumns == null || browserColumns.Count == 0)
                return viewData;

            // Create ordered enumerable with natural sorting
            IOrderedEnumerable<Dictionary<string, object>> sorted = null;

            for (int i = 0; i < browserColumns.Count; i++)
            {
                string columnName = browserColumns[i].Name;

                if (i == 0)
                {
                    // First column - use OrderBy with natural sorting
                    sorted = viewData.OrderBy(v =>
                    {
                        string value = v.ContainsKey(columnName) ? v[columnName]?.ToString() ?? "" : "";
                        return value;
                    }, new NaturalStringComparer());
                }
                else
                {
                    // Subsequent columns - use ThenBy with natural sorting
                    sorted = sorted.ThenBy(v =>
                    {
                        string value = v.ContainsKey(columnName) ? v[columnName]?.ToString() ?? "" : "";
                        return value;
                    }, new NaturalStringComparer());
                }
            }

            // Always add final sort by "Name" column as tiebreaker
            // This handles cases where browser organization parameters are empty or identical
            if (sorted != null)
            {
                sorted = sorted.ThenBy(v =>
                {
                    string value = v.ContainsKey("Name") ? v["Name"]?.ToString() ?? "" : "";
                    return value;
                }, new NaturalStringComparer());
            }

            return sorted != null ? sorted.ToList() : viewData;
        }

        /// <summary>
        /// Natural string comparer that handles numeric prefixes correctly.
        /// Empty strings sort last, then numeric-aware comparison.
        /// </summary>
        private class NaturalStringComparer : System.Collections.Generic.IComparer<string>
        {
            public int Compare(string x, string y)
            {
                if (x == y) return 0;
                if (string.IsNullOrEmpty(x)) return 1;
                if (string.IsNullOrEmpty(y)) return -1;

                // Extract leading numbers if present
                int xNum = 0, yNum = 0;
                bool xHasNum = TryExtractLeadingNumber(x, out xNum);
                bool yHasNum = TryExtractLeadingNumber(y, out yNum);

                // If both have numbers, compare numerically
                if (xHasNum && yHasNum)
                {
                    int numCompare = xNum.CompareTo(yNum);
                    if (numCompare != 0) return numCompare;
                    // Numbers are equal, fall through to string comparison
                }
                // If only one has a number, it comes first
                else if (xHasNum) return -1;
                else if (yHasNum) return 1;

                // Fall back to normal string comparison
                return string.Compare(x, y, StringComparison.OrdinalIgnoreCase);
            }

            private bool TryExtractLeadingNumber(string s, out int number)
            {
                number = 0;
                if (string.IsNullOrEmpty(s)) return false;

                int i = 0;
                while (i < s.Length && char.IsDigit(s[i]))
                {
                    i++;
                }

                if (i == 0) return false;

                return int.TryParse(s.Substring(0, i), out number);
            }
        }
    }
}
