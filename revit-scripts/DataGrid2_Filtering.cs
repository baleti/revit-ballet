using System;
using System.Collections.Generic;
using System.Linq;
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
        // (1) split into tokens - updated regex to handle comparison operators
        List<string> tokens = Regex.Matches(
                searchText,
                @"(\$""[^""]+?""::""[^""]+?""|\$""[^""]+?""\:\:[^ ]+|\$[^ ]+?::""[^""]+?""|\$[^ ]+?::[^ ]+|\$""[^""]+?""\:[^ ]+|\$[^ ]+?:[^ ]+|\$""[^""]+?""|\$[^ ]+|[<>]\d+\.?\d*|""[^""]+""|\S+)")
            .Cast<Match>()
            .Select(m => m.Value.Trim())
            .Where(t => t.Length > 0)
            .ToList();

        // buckets for the parsed filters
        List<List<string>> colVisibilityFilters = new List<List<string>>();
        List<ColumnValueFilter> colValueFilters = new List<ColumnValueFilter>();
        List<string> generalFilters = new List<string>();
        List<ComparisonFilter> comparisonFilters = new List<ComparisonFilter>();

        // (2) parse
        foreach (string rawToken in tokens)
        {
            bool isExcl = rawToken.StartsWith("!");
            string token = isExcl ? rawToken.Substring(1) : rawToken;

            // Check for standalone comparison operators (>50, <50)
            var compMatch = Regex.Match(token, @"^([<>])(\d+\.?\d*)$");
            if (compMatch.Success)
            {
                comparisonFilters.Add(new ComparisonFilter
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
                generalFilters.Add(isExcl ? "!" + StripQuotes(token).ToLowerInvariant() : StripQuotes(token).ToLowerInvariant());
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

            // (2a) column visibility
            colVisibilityFilters.Add(colPieces);

            // (2b) check if value part has comparison operator
            if (!string.IsNullOrWhiteSpace(valPart))
            {
                var valCompMatch = Regex.Match(valPart, @"^([<>])(\d+\.?\d*)$");
                if (valCompMatch.Success)
                {
                    comparisonFilters.Add(new ComparisonFilter
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
                    colValueFilters.Add(f);
                }
            }
        }

        // (3) hide / show columns
        foreach (DataGridViewColumn col in grid.Columns)
        {
            string colName = col.HeaderText.ToLowerInvariant();
            bool show = colVisibilityFilters.Count == 0 ||
                        colVisibilityFilters.Any(parts =>
                            parts.All(p => colName.Contains(p)));
            col.Visible = show;
        }

        // (4) filter the rows
        List<Dictionary<string, object>> filtered = entries.Where(entry =>
        {
            // 4-a column-qualified value filters
            foreach (ColumnValueFilter f in colValueFilters)
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

            // 4-b comparison filters
            foreach (ComparisonFilter f in comparisonFilters)
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

            // 4-c general include / exclude
            if (generalFilters.Count > 0)
            {
                string allValues = string.Join(" ",
                    entry.Values.Where(v => v != null)
                                .Select(v => v.ToString().ToLowerInvariant()));

                bool anyInc = generalFilters.Any(g => !g.StartsWith("!"));
                bool anyExc = generalFilters.Any(g => g.StartsWith("!"));

                if (anyInc &&
                    !generalFilters.Where(g => !g.StartsWith("!"))
                                   .All(inc => allValues.Contains(inc)))
                    return false;

                if (anyExc &&
                    generalFilters.Where(g => g.StartsWith("!"))
                                   .Select(ex => ex.Substring(1))
                                   .Any(ex => allValues.Contains(ex)))
                    return false;
            }

            return true;
        }).ToList();

        return filtered;
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
