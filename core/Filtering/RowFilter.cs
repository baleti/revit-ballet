using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitBallet.Core
{
    /// <summary>
    /// Environment the row matcher runs against. Everything is optional:
    /// without search indexes the matcher falls back to direct row lookups,
    /// without a selection-set lookup #set filters match nothing, and the
    /// foreign-id converter lets a host map opaque id objects (e.g. Revit
    /// ElementId) to longs.
    /// </summary>
    public class RowMatchContext
    {
        public IReadOnlyList<string> PropertyNames { get; set; }

        /// <summary>column name -> row index -> lowercased cell text</summary>
        public Dictionary<string, Dictionary<int, string>> SearchIndexByColumn { get; set; }

        /// <summary>row index -> all lowercased cell text joined with spaces</summary>
        public Dictionary<int, string> SearchIndexAllColumns { get; set; }

        /// <summary>Resolves a selection-set name to member element ids.</summary>
        public Func<string, HashSet<long>> SelectionSetLookup { get; set; }

        /// <summary>Converts host-specific id objects to long (return 0 when not convertible).</summary>
        public Func<object, long> ForeignIdConverter { get; set; }
    }

    /// <summary>
    /// Applies parsed filter groups to rows. OR across groups, AND within a group.
    /// </summary>
    public static class RowFilter
    {
        public static List<Dictionary<string, object>> Filter(
            List<Dictionary<string, object>> entries,
            List<FilterGroup> filterGroups,
            RowMatchContext ctx)
        {
            var matchingIndices = new List<int>();

            for (int i = 0; i < entries.Count; i++)
            {
                // Check if entry matches ANY of the OR groups
                foreach (FilterGroup group in filterGroups)
                {
                    if (Matches(i, entries[i], group, ctx))
                    {
                        matchingIndices.Add(i);
                        break; // Entry matches this OR group, no need to check others
                    }
                }
            }

            var result = new List<Dictionary<string, object>>(matchingIndices.Count);
            foreach (int idx in matchingIndices)
            {
                result.Add(entries[idx]);
            }

            return result;
        }

        public static bool Matches(
            int entryIndex,
            Dictionary<string, object> entry,
            FilterGroup group,
            RowMatchContext ctx)
        {
            IReadOnlyList<string> propertyNames = ctx.PropertyNames ?? new List<string>();

            // Check column-qualified value filters
            foreach (ColumnValueFilter f in group.ColValueFilters)
            {
                List<string> matchCols;

                if (f.IsColumnExactMatch)
                {
                    // For exact column match, join parts with space and compare
                    string exactPattern = string.Join(" ", f.ColumnParts);
                    matchCols = propertyNames.Where(p => p.ToLowerInvariant() == exactPattern).ToList();
                }
                else
                {
                    // Original behavior: all parts must be contained
                    matchCols = propertyNames
                        .Where(p => f.ColumnParts.All(part =>
                                    p.ToLowerInvariant().Contains(part)))
                        .ToList();
                }

                if (matchCols.Count == 0) continue;

                bool valuePresent = matchCols.Any(c =>
                {
                    string cellValue = null;

                    if (ctx.SearchIndexByColumn != null &&
                        ctx.SearchIndexByColumn.ContainsKey(c) &&
                        ctx.SearchIndexByColumn[c].ContainsKey(entryIndex))
                    {
                        cellValue = ctx.SearchIndexByColumn[c][entryIndex];
                    }
                    else
                    {
                        // Fallback to direct lookup if index not available
                        object v;
                        if (entry.TryGetValue(c, out v) && v != null)
                        {
                            cellValue = v.ToString().ToLowerInvariant();
                        }
                    }

                    if (cellValue == null) return false;

                    // Check value match based on exact match flag
                    if (f.IsExactMatch)
                    {
                        return cellValue == f.Value;
                    }
                    else if (f.IsGlobPattern)
                    {
                        return TextMatch.MatchesGlobPattern(cellValue, f.Value);
                    }
                    else
                    {
                        return cellValue.Contains(f.Value);
                    }
                });

                if (!f.IsExclusion && !valuePresent) return false;
                if (f.IsExclusion && valuePresent) return false;
            }

            // Check comparison filters
            foreach (ComparisonFilter f in group.ComparisonFilters)
            {
                bool matchFound = false;

                if (f.ColumnParts == null)
                {
                    // Check all columns
                    foreach (var kvp in entry)
                    {
                        if (kvp.Value != null && TextMatch.TryParseDouble(kvp.Value.ToString(), out double val))
                        {
                            if (f.Operator == ComparisonOperator.GreaterThan && val > f.Value)
                                matchFound = true;
                            else if (f.Operator == ComparisonOperator.LessThan && val < f.Value)
                                matchFound = true;

                            if (matchFound) break;
                        }
                    }
                }
                else
                {
                    // Check specific columns
                    List<string> matchCols = propertyNames
                        .Where(p => f.ColumnParts.All(part =>
                                    p.ToLowerInvariant().Contains(part)))
                        .ToList();

                    foreach (string col in matchCols)
                    {
                        object v;
                        if (entry.TryGetValue(col, out v) && v != null &&
                            TextMatch.TryParseDouble(v.ToString(), out double val))
                        {
                            if (f.Operator == ComparisonOperator.GreaterThan && val > f.Value)
                                matchFound = true;
                            else if (f.Operator == ComparisonOperator.LessThan && val < f.Value)
                                matchFound = true;

                            if (matchFound) break;
                        }
                    }
                }

                if (!f.IsExclusion && !matchFound) return false;
                if (f.IsExclusion && matchFound) return false;
            }

            // Check selection set filters
            foreach (SelectionSetFilter f in group.SelectionSetFilters)
            {
                long elementIdLong = GetRowId(entry, "ElementID", ctx);
                if (elementIdLong == 0)
                    elementIdLong = GetRowId(entry, "Id", ctx);

                // If no valid element ID found, treat as not in set
                if (elementIdLong == 0)
                {
                    if (!f.IsExclusion) return false; // Required to be in set, but no ID found
                    continue; // Excluded from set, no ID means it's not in set (which is what we want)
                }

                HashSet<long> selectionSetIds = ctx.SelectionSetLookup != null
                    ? ctx.SelectionSetLookup(f.SelectionSetName)
                    : new HashSet<long>();
                bool isInSet = selectionSetIds.Contains(elementIdLong);

                // Also check ViewportID field (if present) - allows view rows to match when viewport is in selection set
                if (!isInSet)
                {
                    long viewportIdLong = GetRowId(entry, "ViewportID", ctx);
                    if (viewportIdLong != 0)
                    {
                        isInSet = selectionSetIds.Contains(viewportIdLong);
                    }
                }

                if (!f.IsExclusion && !isInSet) return false; // Required to be in set, but it's not
                if (f.IsExclusion && isInSet) return false;   // Required NOT to be in set, but it is
            }

            // Check general include/exclude filters using index
            if (group.GeneralFilters.Count > 0 || group.GeneralGlobPatterns.Count > 0 || group.GeneralExactFilters.Count > 0)
            {
                // Check exact filters separately for each value
                if (group.GeneralExactFilters.Count > 0)
                {
                    bool anyInc = group.GeneralExactFilters.Any(g => !g.StartsWith("!"));
                    bool anyExc = group.GeneralExactFilters.Any(g => g.StartsWith("!"));

                    // For exact matches, check each cell value individually
                    bool hasExactMatch = false;
                    foreach (var kvp in entry)
                    {
                        if (kvp.Value != null)
                        {
                            string val = kvp.Value.ToString().ToLowerInvariant();

                            // Check inclusion filters
                            if (anyInc)
                            {
                                foreach (string filter in group.GeneralExactFilters.Where(g => !g.StartsWith("!")))
                                {
                                    if (val == filter)
                                    {
                                        hasExactMatch = true;
                                        break;
                                    }
                                }
                            }

                            // Check exclusion filters
                            if (anyExc)
                            {
                                foreach (string filter in group.GeneralExactFilters.Where(g => g.StartsWith("!")))
                                {
                                    string cleanFilter = filter.Substring(1);
                                    if (val == cleanFilter)
                                        return false; // Excluded value found
                                }
                            }
                        }
                    }

                    if (anyInc && !hasExactMatch)
                        return false;
                }

                // Check substring filters (original behavior)
                if (group.GeneralFilters.Count > 0 || group.GeneralGlobPatterns.Count > 0)
                {
                    string allValues = ctx.SearchIndexAllColumns != null && ctx.SearchIndexAllColumns.ContainsKey(entryIndex)
                        ? ctx.SearchIndexAllColumns[entryIndex]
                        : string.Join(" ", entry.Values.Where(v => v != null)
                                                .Select(v => v.ToString().ToLowerInvariant()));

                    // Check regular filters (substring match)
                    bool anyInc = group.GeneralFilters.Any(g => !g.StartsWith("!"));
                    bool anyExc = group.GeneralFilters.Any(g => g.StartsWith("!"));

                    if (anyInc &&
                        !group.GeneralFilters.Where(g => !g.StartsWith("!"))
                                       .All(inc => allValues.Contains(inc)))
                        return false;

                    if (anyExc &&
                        group.GeneralFilters.Where(g => g.StartsWith("!"))
                                       .Select(ex => ex.Substring(1))
                                       .Any(ex => allValues.Contains(ex)))
                        return false;

                    // Check glob patterns
                    foreach (string globPattern in group.GeneralGlobPatterns)
                    {
                        bool isExclusion = globPattern.StartsWith("!");
                        string pattern = isExclusion ? globPattern.Substring(1) : globPattern;

                        // For general glob patterns, check each value individually
                        bool matchFound = false;
                        foreach (var kvp in entry)
                        {
                            if (kvp.Value != null)
                            {
                                string val = kvp.Value.ToString().ToLowerInvariant();
                                if (TextMatch.MatchesGlobPattern(val, pattern))
                                {
                                    matchFound = true;
                                    break;
                                }
                            }
                        }

                        if (!isExclusion && !matchFound) return false;
                        if (isExclusion && matchFound) return false;
                    }
                }
            }

            return true;
        }

        /// <summary>Extracts a numeric id from a row column, using the host converter for foreign id types.</summary>
        private static long GetRowId(Dictionary<string, object> entry, string columnName, RowMatchContext ctx)
        {
            object idObj;
            if (!entry.TryGetValue(columnName, out idObj) || idObj == null)
                return 0;

            if (idObj is long l) return l;
            if (idObj is int i) return i;

            if (ctx.ForeignIdConverter != null)
            {
                long converted = ctx.ForeignIdConverter(idObj);
                if (converted != 0) return converted;
            }

            if (long.TryParse(idObj.ToString(), out long parsed))
                return parsed;

            return 0;
        }
    }
}
