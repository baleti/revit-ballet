using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;

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

    // ──────────────────────────────────────────────────────────────
    //  Helper types
    // ──────────────────────────────────────────────────────────────

    private struct ColumnValueFilter
    {
        public List<string> ColumnParts;   // column-header fragments to match
        public string       Value;         // value to look for in the cell
        public bool         IsExclusion;   // true ⇒ "must NOT contain"
    }

    private enum ComparisonOperator
    {
        GreaterThan,
        LessThan
    }

    private struct ComparisonFilter
    {
        public List<string> ColumnParts;       // column-header fragments to match (null = all columns)
        public ComparisonOperator Operator;    // > or 
        public double Value;                   // numeric value to compare against
        public bool IsExclusion;               // true ⇒ "must NOT match comparison"
    }

    /// <summary>
    /// Represents a group of filters that use AND logic internally.
    /// Multiple FilterGroups are combined with OR logic.
    /// </summary>
    private class FilterGroup
    {
        public List<List<string>> ColVisibilityFilters { get; set; }
        public List<ColumnValueFilter> ColValueFilters { get; set; }
        public List<string> GeneralFilters { get; set; }
        public List<ComparisonFilter> ComparisonFilters { get; set; }

        public FilterGroup()
        {
            ColVisibilityFilters = new List<List<string>>();
            ColValueFilters = new List<ColumnValueFilter>();
            GeneralFilters = new List<string>();
            ComparisonFilters = new List<ComparisonFilter>();
        }
    }

    /// <summary>
    /// Comparer for List<string> to use in HashSet
    /// </summary>
    private class ListStringComparer : IEqualityComparer<List<string>>
    {
        public bool Equals(List<string> x, List<string> y)
        {
            if (x == null && y == null) return true;
            if (x == null || y == null) return false;
            if (x.Count != y.Count) return false;

            return x.SequenceEqual(y);
        }

        public int GetHashCode(List<string> obj)
        {
            if (obj == null) return 0;

            int hash = 17;
            foreach (string s in obj)
            {
                hash = hash * 31 + (s != null ? s.GetHashCode() : 0);
            }
            return hash;
        }
    }

    private class SortCriteria
    {
        public string ColumnName { get; set; }
        public ListSortDirection Direction { get; set; }
    }

    // ──────────────────────────────────────────────────────────────
    //  Utility Methods
    // ──────────────────────────────────────────────────────────────

    private static string StripQuotes(string s)
    {
        return s.StartsWith("\"") && s.EndsWith("\"") && s.Length > 1
            ? s.Substring(1, s.Length - 2)
            : s;
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
}
