using RevitBallet.Core;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RevitBallet.Commands;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Performance Caching Fields
    // ──────────────────────────────────────────────────────────────

    // Virtual mode caching
    private static List<Dictionary<string, object>> _cachedOriginalData;
    private static List<Dictionary<string, object>> _cachedFilteredData;
    private static DataGridView _currentGrid;

    // Search index cache
    private static Dictionary<string, Dictionary<int, string>> _searchIndexByColumn;
    private static Dictionary<int, string> _searchIndexAllColumns;

    // Column visibility cache
    private static HashSet<string> _lastVisibleColumns = new HashSet<string>();
    private static string _lastColumnVisibilityFilter = "";

    // Column ordering cache
    private static string _lastColumnOrderingFilter = "";

    // Selection set cache (name -> ElementId set)
    private static Dictionary<string, HashSet<long>> _selectionSetCache = new Dictionary<string, HashSet<long>>();

    // ──────────────────────────────────────────────────────────────
    //  Internal ID tracking for stable edit tracking
    // ──────────────────────────────────────────────────────────────

    private static long _nextInternalId = 1;
    private const string INTERNAL_ID_KEY = "__DATAGRID_INTERNAL_ID__";

    /// <summary>
    /// Assigns unique internal IDs to all entries that don't already have one.
    /// This provides a stable identifier for edit tracking that is independent of:
    /// - Row position (which changes with filtering/sorting)
    /// - Data content (which can be edited)
    /// - Specific columns (like Handle or DocumentPath which may not exist)
    /// </summary>
    private static void AssignInternalIdsToEntries(List<Dictionary<string, object>> entries)
    {
        foreach (var entry in entries)
        {
            if (!entry.ContainsKey(INTERNAL_ID_KEY))
            {
                entry[INTERNAL_ID_KEY] = _nextInternalId++;
            }
        }
    }

    /// <summary>
    /// AUTOMATIC SYSTEM: Ensure ElementIdObject exists in all rows for edit support
    /// If missing but Id exists, try to reconstruct ElementId from long Id
    /// </summary>
    private static void EnsureElementIdObjectInRows(List<Dictionary<string, object>> entries)
    {
        // Skip if running without Revit API access (e.g., in standalone network launcher)
        if (!_hasRevitApiAccess)
            return;

        foreach (var entry in entries)
        {
            // If ElementIdObject already exists, nothing to do
            if (entry.ContainsKey("ElementIdObject"))
                continue;

            // Try to reconstruct from Id field
            if (entry.ContainsKey("Id"))
            {
                try
                {
                    object idValue = entry["Id"];

                    // Handle different ID types
                    if (idValue is long longId)
                    {
                        entry["ElementIdObject"] = longId.ToElementId();
                    }
                    else if (idValue is int intId)
                    {
                        entry["ElementIdObject"] = intId.ToElementId();
                    }
                    else if (idValue is string strId && long.TryParse(strId, out long parsedId))
                    {
                        entry["ElementIdObject"] = parsedId.ToElementId();
                    }
                }
                catch
                {
                    // If reconstruction fails, entry won't be editable (no ElementId to look up)
                    // This is OK - not all grid entries need to be editable
                }
            }
        }
    }

    /// <summary>
    /// Gets the internal ID for an entry, or creates one if it doesn't exist
    /// </summary>
    public static long GetInternalId(Dictionary<string, object> entry)
    {
        if (entry.ContainsKey(INTERNAL_ID_KEY))
        {
            return System.Convert.ToInt64(entry[INTERNAL_ID_KEY]);
        }

        // Assign a new ID if one doesn't exist
        long newId = _nextInternalId++;
        entry[INTERNAL_ID_KEY] = newId;
        return newId;
    }

    // ──────────────────────────────────────────────────────────────
    //  Helper types
    // ──────────────────────────────────────────────────────────────

    // Filter model types (FilterGroup, ColumnValueFilter, ...) and the
    // ListStringComparer moved to RevitBallet.Core.

    private class SortCriteria
    {
        public string ColumnName { get; set; }
        public ListSortDirection Direction { get; set; }
    }

    // ──────────────────────────────────────────────────────────────
    //  Utility Methods
    // ──────────────────────────────────────────────────────────────

    /// <summary>Get element IDs from a selection set by exact name</summary>
    private static HashSet<long> GetSelectionSetElementIds(string selectionSetName)
    {
        // Check cache first
        if (_selectionSetCache.ContainsKey(selectionSetName))
            return _selectionSetCache[selectionSetName];

        // Return empty set if no UIDocument is available
        var currentUIDocTyped = CurrentUIDoc as Autodesk.Revit.UI.UIDocument;
        if (currentUIDocTyped == null || currentUIDocTyped.Document == null)
            return new HashSet<long>();

        var doc = currentUIDocTyped.Document;
        var result = new HashSet<long>();

        try
        {
            // Find selection set by exact name
            var collector = new Autodesk.Revit.DB.FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.SelectionFilterElement));

            var selectionSet = collector
                .Cast<Autodesk.Revit.DB.SelectionFilterElement>()
                .FirstOrDefault(s => s.Name.Equals(selectionSetName, System.StringComparison.Ordinal));

            if (selectionSet != null)
            {
                // Get element IDs and convert to long values (compatible with all Revit versions)
                foreach (var id in selectionSet.GetElementIds())
                {
                    result.Add(id.AsLong());
                }
            }

            // Cache the result (even if empty, to avoid repeated queries)
            _selectionSetCache[selectionSetName] = result;
        }
        catch
        {
            // If any error occurs, return empty set
            // Cache empty result to avoid repeated failures
            _selectionSetCache[selectionSetName] = result;
        }

        return result;
    }

    /// <summary>Build search index for fast filtering</summary>
    private static void BuildSearchIndex(List<Dictionary<string, object>> data, List<string> propertyNames)
    {
        _searchIndexByColumn = new Dictionary<string, Dictionary<int, string>>();
        _searchIndexAllColumns = new Dictionary<int, string>();

        // Initialize column indices
        foreach (string prop in propertyNames)
        {
            _searchIndexByColumn[prop] = new Dictionary<int, string>();
        }

        // Build indices
        for (int i = 0; i < data.Count; i++)
        {
            var entry = data[i];
            var allValuesBuilder = new System.Text.StringBuilder();

            foreach (string prop in propertyNames)
            {
                object value;
                if (entry.TryGetValue(prop, out value) && value != null)
                {
                    string strVal = value.ToString().ToLowerInvariant();
                    _searchIndexByColumn[prop][i] = strVal;
                    allValuesBuilder.Append(strVal).Append(" ");
                }
            }

            _searchIndexAllColumns[i] = allValuesBuilder.ToString();
        }
    }

    // ──────────────────────────────────────────────────────────────
    //  Object-to-Dictionary Conversion Helpers
    //  (For migrating from old DataGrid1 generic implementation)
    // ──────────────────────────────────────────────────────────────

    private const string ORIGINAL_OBJECT_KEY = "__OriginalObject";

    /// <summary>
    /// Converts a list of objects to DataGrid-compatible dictionaries using reflection.
    /// Stores original object reference for later retrieval.
    /// </summary>
    /// <typeparam name="T">Type of objects to convert</typeparam>
    /// <param name="objects">List of objects to convert</param>
    /// <param name="propertyNames">Property names to extract</param>
    /// <returns>List of dictionaries suitable for DataGrid</returns>
    public static List<Dictionary<string, object>> ConvertToDataGridFormat<T>(
        List<T> objects,
        List<string> propertyNames)
    {
        var result = new List<Dictionary<string, object>>();

        foreach (var obj in objects)
        {
            var dict = new Dictionary<string, object>();

            // Add requested properties using reflection
            foreach (var propName in propertyNames)
            {
                var prop = typeof(T).GetProperty(propName);
                if (prop != null)
                {
                    dict[propName] = prop.GetValue(obj, null);
                }
            }

            // Store original object for later retrieval
            dict[ORIGINAL_OBJECT_KEY] = obj;

            result.Add(dict);
        }

        return result;
    }

    /// <summary>
    /// Extracts original objects from DataGrid result dictionaries.
    /// </summary>
    /// <typeparam name="T">Type of objects to extract</typeparam>
    /// <param name="dictionaries">Result from DataGrid</param>
    /// <returns>List of original objects</returns>
    public static List<T> ExtractOriginalObjects<T>(List<Dictionary<string, object>> dictionaries)
    {
        return dictionaries
            .Where(d => d.ContainsKey(ORIGINAL_OBJECT_KEY))
            .Select(d => (T)d[ORIGINAL_OBJECT_KEY])
            .ToList();
    }
}
