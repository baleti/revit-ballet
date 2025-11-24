using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
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

    // Column ordering cache
    private static string _lastColumnOrderingFilter = "";

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

    private struct ColumnValueFilter
    {
        public List<string> ColumnParts;   // column-header fragments to match
        public string       Value;         // value to look for in the cell
        public bool         IsExclusion;   // true ⇒ "must NOT contain"
        public bool         IsGlobPattern; // true if value contains wildcards
        public bool         IsExactMatch;  // true if value should match exactly
        public bool         IsColumnExactMatch; // true if column should match exactly
    }

    private enum ComparisonOperator
    {
        GreaterThan,
        LessThan
    }

    private struct ComparisonFilter
    {
        public List<string> ColumnParts;       // column-header fragments to match (null = all columns)
        public ComparisonOperator Operator;    // > or <
        public double Value;                   // numeric value to compare against
        public bool IsExclusion;               // true ⇒ "must NOT match comparison"
    }

    /// <summary>
    /// Represents column ordering information
    /// </summary>
    private struct ColumnOrderInfo
    {
        public List<string> ColumnParts;   // column-header fragments to match
        public int Position;               // desired position (1-based)
        public bool IsExactMatch;          // true if column should match exactly
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
        public List<string> GeneralGlobPatterns { get; set; } // New field for glob patterns
        public List<ColumnOrderInfo> ColumnOrdering { get; set; } // New field for column ordering
        public List<bool> ColVisibilityExactMatch { get; set; } // Track exact match for visibility
        public List<string> GeneralExactFilters { get; set; } // Exact match general filters

        public FilterGroup()
        {
            ColVisibilityFilters = new List<List<string>>();
            ColValueFilters = new List<ColumnValueFilter>();
            GeneralFilters = new List<string>();
            ComparisonFilters = new List<ComparisonFilter>();
            GeneralGlobPatterns = new List<string>(); // Initialize new field
            ColumnOrdering = new List<ColumnOrderInfo>(); // Initialize column ordering
            ColVisibilityExactMatch = new List<bool>(); // Initialize exact match tracking
            GeneralExactFilters = new List<string>(); // Initialize exact match filters
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

    /// <summary>Check if a string contains glob wildcards</summary>
    private static bool ContainsGlobWildcards(string pattern)
    {
        return pattern != null && pattern.Contains("*");
    }

    /// <summary>Convert glob pattern to regex pattern</summary>
    private static string GlobToRegexPattern(string globPattern)
    {
        // Escape special regex characters except *
        string escaped = Regex.Escape(globPattern).Replace("\\*", ".*");
        return "^" + escaped + "$";
    }

    /// <summary>Check if a value matches a glob pattern</summary>
    private static bool MatchesGlobPattern(string value, string pattern)
    {
        if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(pattern))
            return false;

        // Convert to lowercase for case-insensitive matching
        value = value.ToLowerInvariant();
        pattern = pattern.ToLowerInvariant();

        // If no wildcards, use simple contains (backward compatibility)
        if (!pattern.Contains("*"))
            return value.Contains(pattern);

        // Convert glob to regex and match
        string regexPattern = GlobToRegexPattern(pattern);
        return Regex.IsMatch(value, regexPattern);
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
