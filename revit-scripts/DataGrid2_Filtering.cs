using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Filtering Logic (Optimized)
    // ──────────────────────────────────────────────────────────────

    /// <summary>Applies all filters to the data and returns filtered result</summary>
    private static List<Dictionary<string, object>> ApplyFilters(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        string searchText,
        DataGridView grid)
    {
        // Quick return for empty filter
        if (string.IsNullOrWhiteSpace(searchText))
        {
            UpdateColumnVisibilityOptimized(grid, new HashSet<List<string>>(new ListStringComparer()), new List<ColumnOrderInfo>(), new List<bool>());
            return entries;
        }

        // Parse filter groups once
        List<string> orGroups = SplitByOrOperator(searchText);
        List<FilterGroup> filterGroups = new List<FilterGroup>();

        foreach (string orGroup in orGroups)
        {
            FilterGroup group = ParseFilterGroup(orGroup.Trim());
            filterGroups.Add(group);
        }

        // Update column visibility and ordering (optimized)
        HashSet<List<string>> allColVisibilityFilters = new HashSet<List<string>>(
            filterGroups.SelectMany(g => g.ColVisibilityFilters),
            new ListStringComparer());

        // Collect all column ordering from all groups
        List<ColumnOrderInfo> allColumnOrdering = filterGroups.SelectMany(g => g.ColumnOrdering).ToList();
        
        // Collect exact match flags for visibility
        List<bool> allColVisibilityExactMatch = filterGroups.SelectMany(g => g.ColVisibilityExactMatch).ToList();

        UpdateColumnVisibilityOptimized(grid, allColVisibilityFilters, allColumnOrdering, allColVisibilityExactMatch);

        // Use optimized row filtering
        List<Dictionary<string, object>> filtered = FilterRowsOptimized(entries, filterGroups, propertyNames);

        return filtered;
    }

    /// <summary>Optimized column visibility update - only updates when changed</summary>
    private static void UpdateColumnVisibilityOptimized(DataGridView grid, HashSet<List<string>> filters, List<ColumnOrderInfo> columnOrdering, List<bool> exactMatchFlags)
    {
        // Create a key for the current filter state (including ordering)
        string filterKey = string.Join("|", filters.Select(f => string.Join(",", f)));
        string orderingKey = string.Join("|", columnOrdering.Select(o => o.Position + ":" + string.Join(",", o.ColumnParts) + ":" + o.IsExactMatch));
        string exactMatchKey = string.Join("|", exactMatchFlags.Select(e => e ? "1" : "0"));
        string combinedKey = filterKey + "||" + orderingKey + "||" + exactMatchKey;

        // Skip if nothing changed
        if (combinedKey == _lastColumnVisibilityFilter + "||" + _lastColumnOrderingFilter)
            return;

        _lastColumnVisibilityFilter = filterKey + "||" + exactMatchKey;
        _lastColumnOrderingFilter = orderingKey;
        
        var newVisible = new HashSet<string>();

        // Update visibility
        if (filters.Count > 0)
        {
            var filterList = filters.ToList();
            for (int i = 0; i < filterList.Count; i++)
            {
                var parts = filterList[i];
                bool isExactMatch = i < exactMatchFlags.Count && exactMatchFlags[i];
                
                foreach (DataGridViewColumn col in grid.Columns)
                {
                    string colName = col.HeaderText.ToLowerInvariant();
                    bool show = false;
                    
                    if (isExactMatch)
                    {
                        // For exact match, join parts with space and compare
                        string exactPattern = string.Join(" ", parts);
                        show = colName == exactPattern;
                    }
                    else
                    {
                        // Original behavior: all parts must be contained
                        show = parts.All(p => colName.Contains(p));
                    }
                    
                    if (show) newVisible.Add(col.Name);
                }
            }
        }
        else
        {
            // No filters, show all columns
            foreach (DataGridViewColumn col in grid.Columns)
            {
                newVisible.Add(col.Name);
            }
        }

        // Apply column ordering if any
        if (columnOrdering.Count > 0)
        {
            grid.SuspendLayout(); // Prevent flicker

            // First, update visibility
            foreach (DataGridViewColumn col in grid.Columns)
            {
                col.Visible = newVisible.Contains(col.Name);
            }

            // Create a list to track new display indices
            var columnPositions = new List<Tuple<DataGridViewColumn, int>>();

            // Process columns with explicit ordering
            foreach (var orderInfo in columnOrdering.OrderBy(o => o.Position))
            {
                foreach (DataGridViewColumn col in grid.Columns)
                {
                    if (!col.Visible) continue;

                    string colName = col.HeaderText.ToLowerInvariant();
                    bool matches = false;
                    
                    if (orderInfo.IsExactMatch)
                    {
                        // For exact match, join parts with space and compare
                        string exactPattern = string.Join(" ", orderInfo.ColumnParts);
                        matches = colName == exactPattern;
                    }
                    else
                    {
                        // Original behavior: all parts must be contained
                        matches = orderInfo.ColumnParts.All(part => colName.Contains(part));
                    }
                    
                    if (matches)
                    {
                        // Check if this column is already in the list
                        if (!columnPositions.Any(cp => cp.Item1 == col))
                        {
                            columnPositions.Add(Tuple.Create(col, orderInfo.Position));
                        }
                    }
                }
            }

            // Add remaining visible columns that don't have explicit ordering
            int nextPosition = columnPositions.Count > 0 ? columnPositions.Max(cp => cp.Item2) + 1 : 1;
            foreach (DataGridViewColumn col in grid.Columns)
            {
                if (col.Visible && !columnPositions.Any(cp => cp.Item1 == col))
                {
                    columnPositions.Add(Tuple.Create(col, nextPosition++));
                }
            }

            // Apply the new display order
            int displayIndex = 0;
            foreach (var colPos in columnPositions.OrderBy(cp => cp.Item2))
            {
                colPos.Item1.DisplayIndex = displayIndex++;
            }

            grid.ResumeLayout();
            _lastVisibleColumns = newVisible;
        }
        else if (!_lastVisibleColumns.SetEquals(newVisible))
        {
            // Only update visibility if no ordering is specified and visibility changed
            grid.SuspendLayout(); // Prevent flicker
            foreach (DataGridViewColumn col in grid.Columns)
            {
                col.Visible = newVisible.Contains(col.Name);
            }
            grid.ResumeLayout();
            _lastVisibleColumns = newVisible;
        }
    }

    /// <summary>Optimized row filtering using search index</summary>
    private static List<Dictionary<string, object>> FilterRowsOptimized(
        List<Dictionary<string, object>> entries,
        List<FilterGroup> filterGroups,
        List<string> propertyNames)
    {
        // Use indices for filtering
        var matchingIndices = new List<int>();

        for (int i = 0; i < entries.Count; i++)
        {
            // Check if entry matches ANY of the OR groups
            foreach (FilterGroup group in filterGroups)
            {
                if (EntryMatchesFilterGroupOptimized(i, entries[i], group, propertyNames))
                {
                    matchingIndices.Add(i);
                    break; // Entry matches this OR group, no need to check others
                }
            }
        }

        // Build result list
        var result = new List<Dictionary<string, object>>(matchingIndices.Count);
        foreach (int idx in matchingIndices)
        {
            result.Add(entries[idx]);
        }

        return result;
    }

    /// <summary>Optimized entry matching using search index</summary>
    private static bool EntryMatchesFilterGroupOptimized(
        int entryIndex,
        Dictionary<string, object> entry,
        FilterGroup group,
        List<string> propertyNames)
    {
        // Check column-qualified value filters
        foreach (ColumnValueFilter f in group.ColValueFilters)
        {
            List<string> matchCols = new List<string>();
            
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

                if (_searchIndexByColumn != null &&
                    _searchIndexByColumn.ContainsKey(c) &&
                    _searchIndexByColumn[c].ContainsKey(entryIndex))
                {
                    cellValue = _searchIndexByColumn[c][entryIndex];
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
                    return MatchesGlobPattern(cellValue, f.Value);
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
                    if (kvp.Value != null && TryParseDouble(kvp.Value.ToString(), out double val))
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
                        TryParseDouble(v.ToString(), out double val))
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
                string allValues = _searchIndexAllColumns != null && _searchIndexAllColumns.ContainsKey(entryIndex)
                    ? _searchIndexAllColumns[entryIndex]
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
                            if (MatchesGlobPattern(val, pattern))
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

    /// <summary>Split search text by || operator, respecting quotes</summary>
    private static List<string> SplitByOrOperator(string searchText)
    {
        List<string> groups = new List<string>();
        StringBuilder current = new StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < searchText.Length; i++)
        {
            char c = searchText[i];

            if (c == '"' && (i == 0 || searchText[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
                current.Append(c);
            }
            else if (!inQuotes && i < searchText.Length - 1 &&
                     c == '|' && searchText[i + 1] == '|')
            {
                // Found || outside quotes
                groups.Add(current.ToString());
                current.Clear();
                i++; // Skip the second |
            }
            else
            {
                current.Append(c);
            }
        }

        // Add the last group
        if (current.Length > 0)
        {
            groups.Add(current.ToString());
        }

        // If no || found, return the whole text as a single group
        if (groups.Count == 0)
        {
            groups.Add(searchText);
        }

        return groups;
    }

    /// <summary>Parse a single filter group (AND logic within the group)</summary>
    private static FilterGroup ParseFilterGroup(string groupText)
    {
        FilterGroup group = new FilterGroup();

        // Split into tokens - updated regex to handle 'e' prefix for exact matching
        List<string> tokens = Regex.Matches(
                groupText,
                @"(!\d+\$e""[^""]+?""::e""[^""]+?""|!\d+\$e""[^""]+?""::[^ ]+|!\d+\$[^ ]+?::e""[^""]+?""|!\d+\$e""[^""]+?""::""[^""]+?""|!\d+\$""[^""]+?""::e""[^""]+?""|!\d+\$e""[^""]+?""|!\d+\$""[^""]+?""::""[^""]+?""|!\d+\$""[^""]+?""\:\:[^ ]+|!\d+\$[^ ]+?::""[^""]+?""|!\d+\$[^ ]+?::[^ ]+|!\d+\$""[^""]+?""\:[^ ]+|!\d+\$[^ ]+?:[^ ]+|!\d+\$""[^""]+?""|!\d+\$[^ ]+|\d+\$e""[^""]+?""::e""[^""]+?""|\d+\$e""[^""]+?""::[^ ]+|\d+\$[^ ]+?::e""[^""]+?""|\d+\$e""[^""]+?""::""[^""]+?""|\d+\$""[^""]+?""::e""[^""]+?""|\d+\$e""[^""]+?""|\d+\$""[^""]+?""::""[^""]+?""|\d+\$""[^""]+?""\:\:[^ ]+|\d+\$[^ ]+?::""[^""]+?""|\d+\$[^ ]+?::[^ ]+|\d+\$""[^""]+?""\:[^ ]+|\d+\$[^ ]+?:[^ ]+|\d+\$""[^""]+?""|\d+\$[^ ]+|!\$e""[^""]+?""::e""[^""]+?""|!\$e""[^""]+?""::[^ ]+|!\$[^ ]+?::e""[^""]+?""|!\$e""[^""]+?""::""[^""]+?""|!\$""[^""]+?""::e""[^""]+?""|!\$e""[^""]+?""|!\$""[^""]+?""::""[^""]+?""|!\$""[^""]+?""\:\:[^ ]+|!\$[^ ]+?::""[^""]+?""|!\$[^ ]+?::[^ ]+|!\$""[^""]+?""\:[^ ]+|!\$[^ ]+?:[^ ]+|!\$""[^""]+?""|!\$[^ ]+|\$e""[^""]+?""::e""[^""]+?""|\$e""[^""]+?""::[^ ]+|\$[^ ]+?::e""[^""]+?""|\$e""[^""]+?""::""[^""]+?""|\$""[^""]+?""::e""[^""]+?""|\$e""[^""]+?""|\$""[^""]+?""::""[^""]+?""|\$""[^""]+?""\:\:[^ ]+|\$[^ ]+?::""[^""]+?""|\$[^ ]+?::[^ ]+|\$""[^""]+?""\:[^ ]+|\$[^ ]+?:[^ ]+|\$""[^""]+?""|\$[^ ]+|[<>]\d+\.?\d*|e""[^""]+?""|""[^""]+""|\S+)")
            .Cast<Match>()
            .Select(m => m.Value.Trim())
            .Where(t => t.Length > 0)
            .ToList();

        // Parse each token
        foreach (string rawToken in tokens)
        {
            bool isExcl = rawToken.StartsWith("!");
            string token = isExcl ? rawToken.Substring(1) : rawToken;

            // Check for standalone comparison operators (>50, <50)
            var compMatch = Regex.Match(token, @"^([<>])(\d+\.?\d*)$");
            if (compMatch.Success)
            {
                group.ComparisonFilters.Add(new ComparisonFilter
                {
                    Operator = compMatch.Groups[1].Value == ">" ? ComparisonOperator.GreaterThan : ComparisonOperator.LessThan,
                    Value = double.Parse(compMatch.Groups[2].Value),
                    ColumnParts = null, // null means check all columns
                    IsExclusion = isExcl
                });
                continue;
            }

            // Check for numeric prefix before $ (e.g., 1$col, 2$"column name")
            var numericPrefixMatch = Regex.Match(token, @"^(\d+)\$(.+)$");
            int? columnPosition = null;
            string tokenWithoutPosition = token;

            if (numericPrefixMatch.Success)
            {
                columnPosition = int.Parse(numericPrefixMatch.Groups[1].Value);
                tokenWithoutPosition = "$" + numericPrefixMatch.Groups[2].Value;
                token = tokenWithoutPosition; // Process the rest as normal $column syntax
            }

            // Check for exact match prefix on general tokens
            bool isGeneralExactMatch = false;
            if (!token.StartsWith("$") && token.StartsWith("e\"") && token.EndsWith("\"") && token.Length > 3)
            {
                isGeneralExactMatch = true;
                token = token.Substring(1); // Remove 'e' prefix
            }

            // plain (general) token
            if (!token.StartsWith("$"))
            {
                string cleanToken = StripQuotes(token).ToLowerInvariant();

                if (isGeneralExactMatch)
                {
                    // Exact match filter
                    group.GeneralExactFilters.Add(isExcl ? "!" + cleanToken : cleanToken);
                }
                else if (ContainsGlobWildcards(cleanToken))
                {
                    // Glob pattern
                    group.GeneralGlobPatterns.Add(isExcl ? "!" + cleanToken : cleanToken);
                }
                else
                {
                    // Regular substring filter
                    group.GeneralFilters.Add(isExcl ? "!" + cleanToken : cleanToken);
                }
                continue;
            }

            // token begins with '$' → column-qualified
            string body = token.Substring(1); // drop '$'
            
            // Check for exact match on column
            bool isColumnExactMatch = false;
            if (body.StartsWith("e\"") && body.Contains("\""))
            {
                isColumnExactMatch = true;
                body = body.Substring(1); // Remove 'e' prefix
            }
            
            int dblColonPos = body.IndexOf("::", StringComparison.Ordinal);
            int colonPos = dblColonPos >= 0 ? dblColonPos : body.IndexOf(':');

            string colPart = colonPos > 0 ? body.Substring(0, colonPos) : body;
            string valPart = "";
            if (colonPos > 0)
            {
                int start = colonPos + (dblColonPos >= 0 ? 2 : 1);
                valPart = body.Substring(start);
            }

            // Check for exact match on value
            bool isValueExactMatch = false;
            if (!string.IsNullOrWhiteSpace(valPart) && valPart.StartsWith("e\"") && valPart.EndsWith("\"") && valPart.Length > 3)
            {
                isValueExactMatch = true;
                valPart = valPart.Substring(1); // Remove 'e' prefix
            }

            bool quotedCol = colPart.StartsWith("\"") && colPart.EndsWith("\"");
            string cleanCol = StripQuotes(colPart).ToLowerInvariant();
            List<string> colPieces = quotedCol
                ? cleanCol.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList()
                : new List<string> { cleanCol };

            if (colPieces.Count == 0) continue;

            // Column visibility with exact match tracking
            group.ColVisibilityFilters.Add(colPieces);
            group.ColVisibilityExactMatch.Add(isColumnExactMatch);

            // If we have a column position, add it to ordering
            if (columnPosition.HasValue && !isExcl)
            {
                group.ColumnOrdering.Add(new ColumnOrderInfo
                {
                    ColumnParts = colPieces,
                    Position = columnPosition.Value,
                    IsExactMatch = isColumnExactMatch
                });
            }

            // Check if value part has comparison operator
            if (!string.IsNullOrWhiteSpace(valPart))
            {
                var valCompMatch = Regex.Match(valPart, @"^([<>])(\d+\.?\d*)$");
                if (valCompMatch.Success)
                {
                    group.ComparisonFilters.Add(new ComparisonFilter
                    {
                        Operator = valCompMatch.Groups[1].Value == ">" ? ComparisonOperator.GreaterThan : ComparisonOperator.LessThan,
                        Value = double.Parse(valCompMatch.Groups[2].Value),
                        ColumnParts = colPieces,
                        IsExclusion = isExcl
                    });
                }
                else
                {
                    // Regular value filter
                    string cleanValue = StripQuotes(valPart).ToLowerInvariant();
                    ColumnValueFilter f = new ColumnValueFilter
                    {
                        ColumnParts = colPieces,
                        Value = cleanValue,
                        IsExclusion = isExcl,
                        IsGlobPattern = ContainsGlobWildcards(cleanValue),
                        IsExactMatch = isValueExactMatch,
                        IsColumnExactMatch = isColumnExactMatch
                    };
                    group.ColValueFilters.Add(f);
                }
            }
        }

        return group;
    }

    /// <summary>Try to parse a string as a double, handling common formats</summary>
    private static bool TryParseDouble(string s, out double result)
    {
        result = 0;
        if (string.IsNullOrWhiteSpace(s)) return false;

        // Remove common formatting characters
        s = s.Replace(",", "").Replace("$", "").Trim();

        // Check if it's a percentage
        if (s.EndsWith("%"))
        {
            s = s.TrimEnd('%');
            if (double.TryParse(s, out result))
            {
                result /= 100; // Convert percentage to decimal
                return true;
            }
        }

        return double.TryParse(s, out result);
    }
}
