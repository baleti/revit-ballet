using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Filtering Logic
    // ──────────────────────────────────────────────────────────────

    /// <summary>Applies all filters to the data and returns filtered result</summary>
    private static List<Dictionary<string, object>> ApplyFilters(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        string searchText,
        DataGridView grid)
    {
        // First, split by || to get OR groups
        List<string> orGroups = SplitByOrOperator(searchText);
        
        // Process each OR group independently
        List<FilterGroup> filterGroups = new List<FilterGroup>();
        
        foreach (string orGroup in orGroups)
        {
            FilterGroup group = ParseFilterGroup(orGroup.Trim());
            filterGroups.Add(group);
        }
        
        // Apply column visibility based on ALL groups (union of all column filters)
        HashSet<List<string>> allColVisibilityFilters = new HashSet<List<string>>(
            filterGroups.SelectMany(g => g.ColVisibilityFilters),
            new ListStringComparer());
        
        foreach (DataGridViewColumn col in grid.Columns)
        {
            string colName = col.HeaderText.ToLowerInvariant();
            bool show = allColVisibilityFilters.Count == 0 ||
                        allColVisibilityFilters.Any(parts =>
                            parts.All(p => colName.Contains(p)));
            col.Visible = show;
        }

        // Filter rows - a row passes if it matches ANY of the OR groups
        List<Dictionary<string, object>> filtered = entries.Where(entry =>
        {
            // Check if entry matches ANY of the OR groups
            foreach (FilterGroup group in filterGroups)
            {
                if (EntryMatchesFilterGroup(entry, group, propertyNames))
                {
                    return true; // Entry matches this OR group
                }
            }
            return false; // Entry doesn't match any OR group
        }).ToList();

        return filtered;
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
        
        // Split into tokens - updated regex to handle comparison operators
        List<string> tokens = Regex.Matches(
                groupText,
                @"(\$""[^""]+?""::""[^""]+?""|\$""[^""]+?""\:\:[^ ]+|\$[^ ]+?::""[^""]+?""|\$[^ ]+?::[^ ]+|\$""[^""]+?""\:[^ ]+|\$[^ ]+?:[^ ]+|\$""[^""]+?""|\$[^ ]+|[<>]\d+\.?\d*|""[^""]+""|\S+)")
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

            // plain (general) token
            if (!token.StartsWith("$"))
            {
                group.GeneralFilters.Add(isExcl ? "!" + StripQuotes(token).ToLowerInvariant() : StripQuotes(token).ToLowerInvariant());
                continue;
            }

            // token begins with '$' → column-qualified
            string body = token.Substring(1); // drop '$'
            int dblColonPos = body.IndexOf("::", StringComparison.Ordinal);
            int colonPos = dblColonPos >= 0 ? dblColonPos : body.IndexOf(':');

            string colPart = colonPos > 0 ? body.Substring(0, colonPos) : body;
            string valPart = "";
            if (colonPos > 0)
            {
                int start = colonPos + (dblColonPos >= 0 ? 2 : 1);
                valPart = body.Substring(start);
            }

            bool quotedCol = colPart.StartsWith("\"") && colPart.EndsWith("\"");
            string cleanCol = StripQuotes(colPart).ToLowerInvariant();
            List<string> colPieces = quotedCol
                ? cleanCol.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList()
                : new List<string> { cleanCol };

            if (colPieces.Count == 0) continue;

            // Column visibility
            group.ColVisibilityFilters.Add(colPieces);

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
                    ColumnValueFilter f = new ColumnValueFilter
                    {
                        ColumnParts = colPieces,
                        Value = StripQuotes(valPart).ToLowerInvariant(),
                        IsExclusion = isExcl
                    };
                    group.ColValueFilters.Add(f);
                }
            }
        }
        
        return group;
    }

    /// <summary>Check if an entry matches all filters in a group (AND logic)</summary>
    private static bool EntryMatchesFilterGroup(
        Dictionary<string, object> entry,
        FilterGroup group,
        List<string> propertyNames)
    {
        // Check column-qualified value filters
        foreach (ColumnValueFilter f in group.ColValueFilters)
        {
            List<string> matchCols = propertyNames
                .Where(p => f.ColumnParts.All(part =>
                            p.ToLowerInvariant().Contains(part)))
                .ToList();

            if (matchCols.Count == 0) continue;

            bool valuePresent = matchCols.Any(c =>
            {
                object v;
                return entry.TryGetValue(c, out v) &&
                       v != null &&
                       v.ToString().ToLowerInvariant().Contains(f.Value);
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

        // Check general include/exclude filters
        if (group.GeneralFilters.Count > 0)
        {
            string allValues = string.Join(" ",
                entry.Values.Where(v => v != null)
                            .Select(v => v.ToString().ToLowerInvariant()));

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
        }

        return true;
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
