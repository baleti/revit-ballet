using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RevitBallet.Core
{
    /// <summary>
    /// Parses the DataGrid search-box query language into filter groups.
    ///
    /// Syntax summary:
    ///   term term          - AND of substring matches across all columns
    ///   a || b             - OR groups
    ///   !term              - exclusion
    ///   "multi word"       - quoted term
    ///   e"exact"           - exact cell match
    ///   *glob*             - glob pattern
    ///   &gt;50  &lt;50           - numeric comparison across all columns
    ///   $col:value         - value filter on columns whose header contains col
    ///   $col::value        - same, :: separator for values containing ':'
    ///   $"col name":value  - quoted column
    ///   $e"col":value      - exact column match
    ///   2$col              - show column at position 2
    ///   #name  #"set name" - element is member of a Revit selection set
    /// </summary>
    public static class FilterQueryParser
    {
        /// <summary>Parses a full search text into OR-combined filter groups.</summary>
        public static List<FilterGroup> Parse(string searchText)
        {
            List<string> orGroups = SplitByOrOperator(searchText);
            List<FilterGroup> filterGroups = new List<FilterGroup>();

            foreach (string orGroup in orGroups)
            {
                filterGroups.Add(ParseGroup(orGroup.Trim()));
            }

            return filterGroups;
        }

        /// <summary>Split search text by || operator, respecting quotes</summary>
        public static List<string> SplitByOrOperator(string searchText)
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
        public static FilterGroup ParseGroup(string groupText)
        {
            FilterGroup group = new FilterGroup();

            // Split into tokens - updated regex to handle 'e' prefix for exact matching and #"selection set" syntax (both quoted and unquoted)
            List<string> tokens = Regex.Matches(
                    groupText,
                    @"(!?#""[^""]+?""|!?#[^ ]+|!\d+\$e""[^""]+?""::e""[^""]+?""|!\d+\$e""[^""]+?""::[^ ]+|!\d+\$[^ ]+?::e""[^""]+?""|!\d+\$e""[^""]+?""::""[^""]+?""|!\d+\$""[^""]+?""::e""[^""]+?""|!\d+\$e""[^""]+?""|!\d+\$""[^""]+?""::""[^""]+?""|!\d+\$""[^""]+?""\:\:[^ ]+|!\d+\$[^ ]+?::""[^""]+?""|!\d+\$[^ ]+?::[^ ]+|!\d+\$""[^""]+?""\:[^ ]+|!\d+\$[^ ]+?:[^ ]+|!\d+\$""[^""]+?""|!\d+\$[^ ]+|\d+\$e""[^""]+?""::e""[^""]+?""|\d+\$e""[^""]+?""::[^ ]+|\d+\$[^ ]+?::e""[^""]+?""|\d+\$e""[^""]+?""::""[^""]+?""|\d+\$""[^""]+?""::e""[^""]+?""|\d+\$e""[^""]+?""|\d+\$""[^""]+?""::""[^""]+?""|\d+\$""[^""]+?""\:\:[^ ]+|\d+\$[^ ]+?::""[^""]+?""|\d+\$[^ ]+?::[^ ]+|\d+\$""[^""]+?""\:[^ ]+|\d+\$[^ ]+?:[^ ]+|\d+\$""[^""]+?""|\d+\$[^ ]+|!\$e""[^""]+?""::e""[^""]+?""|!\$e""[^""]+?""::[^ ]+|!\$[^ ]+?::e""[^""]+?""|!\$e""[^""]+?""::""[^""]+?""|!\$""[^""]+?""::e""[^""]+?""|!\$e""[^""]+?""|!\$""[^""]+?""::""[^""]+?""|!\$""[^""]+?""\:\:[^ ]+|!\$[^ ]+?::""[^""]+?""|!\$[^ ]+?::[^ ]+|!\$""[^""]+?""\:[^ ]+|!\$[^ ]+?:[^ ]+|!\$""[^""]+?""|!\$[^ ]+|\$e""[^""]+?""::e""[^""]+?""|\$e""[^""]+?""::[^ ]+|\$[^ ]+?::e""[^""]+?""|\$e""[^""]+?""::""[^""]+?""|\$""[^""]+?""::e""[^""]+?""|\$e""[^""]+?""|\$""[^""]+?""::""[^""]+?""|\$""[^""]+?""\:\:[^ ]+|\$[^ ]+?::""[^""]+?""|\$[^ ]+?::[^ ]+|\$""[^""]+?""\:[^ ]+|\$[^ ]+?:[^ ]+|\$""[^""]+?""|\$[^ ]+|[<>]\d+\.?\d*|e""[^""]+?""|""[^""]+""|\S+)")
                .Cast<Match>()
                .Select(m => m.Value.Trim())
                .Where(t => t.Length > 0)
                .ToList();

            // Parse each token
            foreach (string rawToken in tokens)
            {
                bool isExcl = rawToken.StartsWith("!");
                string token = isExcl ? rawToken.Substring(1) : rawToken;

                // Check for selection set filter (#"selection set name" or #name)
                if (token.StartsWith("#"))
                {
                    string selectionSetName;
                    if (token.StartsWith("#\"") && token.EndsWith("\"") && token.Length > 3)
                    {
                        // Quoted selection set name: #"my selection set"
                        selectionSetName = TextMatch.StripQuotes(token.Substring(1)); // Remove # and quotes
                    }
                    else if (token.Length > 1)
                    {
                        // Unquoted selection set name: #myselectionset
                        selectionSetName = token.Substring(1); // Remove # only
                    }
                    else
                    {
                        // Just "#" with nothing after it - skip
                        continue;
                    }

                    group.SelectionSetFilters.Add(new SelectionSetFilter
                    {
                        SelectionSetName = selectionSetName,
                        IsExclusion = isExcl
                    });
                    continue;
                }

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

                if (numericPrefixMatch.Success)
                {
                    columnPosition = int.Parse(numericPrefixMatch.Groups[1].Value);
                    token = "$" + numericPrefixMatch.Groups[2].Value; // Process the rest as normal $column syntax
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
                    string cleanToken = TextMatch.StripQuotes(token).ToLowerInvariant();

                    if (isGeneralExactMatch)
                    {
                        // Exact match filter
                        group.GeneralExactFilters.Add(isExcl ? "!" + cleanToken : cleanToken);
                    }
                    else if (TextMatch.ContainsGlobWildcards(cleanToken))
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

                // token begins with '$' -> column-qualified
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
                string cleanCol = TextMatch.StripQuotes(colPart).ToLowerInvariant();
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
                        string cleanValue = TextMatch.StripQuotes(valPart).ToLowerInvariant();
                        ColumnValueFilter f = new ColumnValueFilter
                        {
                            ColumnParts = colPieces,
                            Value = cleanValue,
                            IsExclusion = isExcl,
                            IsGlobPattern = TextMatch.ContainsGlobWildcards(cleanValue),
                            IsExactMatch = isValueExactMatch,
                            IsColumnExactMatch = isColumnExactMatch
                        };
                        group.ColValueFilters.Add(f);
                    }
                }
            }

            return group;
        }
    }
}
