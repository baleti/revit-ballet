using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Sorting Logic
    // ──────────────────────────────────────────────────────────────
    
    private static readonly NaturalComparer naturalComparer = new NaturalComparer();

    /// <summary>A string comparer that sorts "A2" before "A10" and handles mixed numeric/text data.</summary>
    private sealed class NaturalComparer : IComparer<object>
    {
        public int Compare(object x, object y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            string s1 = x.ToString();
            string s2 = y.ToString();

            // Handle special non-numeric values that should be treated as text
            bool s1IsNonNumeric = IsNonNumericValue(s1);
            bool s2IsNonNumeric = IsNonNumericValue(s2);
            
            // If both are non-numeric, compare as strings
            if (s1IsNonNumeric && s2IsNonNumeric)
            {
                return string.Compare(s1, s2, StringComparison.OrdinalIgnoreCase);
            }
            
            // If one is non-numeric and one is numeric, non-numeric comes last
            if (s1IsNonNumeric && !s2IsNonNumeric) return 1;
            if (!s1IsNonNumeric && s2IsNonNumeric) return -1;

            // Try to parse as numbers
            double numA, numB;
            bool aIsNum = double.TryParse(s1, out numA);
            bool bIsNum = double.TryParse(s2, out numB);
            
            if (aIsNum && bIsNum) return numA.CompareTo(numB);

            // Fall back to natural string comparison
            return CompareNatural(s1, s2);
        }

        /// <summary>Checks if a value should be treated as non-numeric text (like "-", "N/A", etc.)</summary>
        private static bool IsNonNumericValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return true;
            
            // Single dash or common placeholder values
            if (value == "-" || 
                value.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("NULL", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("NONE", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("--", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            
            // If it can't be parsed as a number, treat as non-numeric
            double dummy;
            return !double.TryParse(value, out dummy);
        }

        private static int CompareNatural(string a, string b)
        {
            int i = 0, j = 0;
            while (i < a.Length && j < b.Length)
            {
                if (char.IsDigit(a[i]) && char.IsDigit(b[j]))
                {
                    int startI = i;
                    while (i < a.Length && char.IsDigit(a[i])) i++;
                    int startJ = j;
                    while (j < b.Length && char.IsDigit(b[j])) j++;

                    string numA = a.Substring(startI, i - startI).TrimStart('0');
                    string numB = b.Substring(startJ, j - startJ).TrimStart('0');
                    if (numA.Length == 0) numA = "0";
                    if (numB.Length == 0) numB = "0";

                    int cmp = numA.Length.CompareTo(numB.Length);
                    if (cmp != 0) return cmp;

                    cmp = string.Compare(numA, numB, StringComparison.Ordinal);
                    if (cmp != 0) return cmp;
                }
                else
                {
                    int cmp = a[i].CompareTo(b[j]);
                    if (cmp != 0) return cmp;
                    i++;
                    j++;
                }
            }
            return a.Length.CompareTo(b.Length);
        }
    }

    /// <summary>Applies multi-column sorting to the data</summary>
    private static List<Dictionary<string, object>> ApplySorting(
        List<Dictionary<string, object>> data,
        List<SortCriteria> sortCriteria)
    {
        if (sortCriteria.Count == 0) return data;

        IOrderedEnumerable<Dictionary<string, object>> ordered = null;
        foreach (SortCriteria sc in sortCriteria)
        {
            Func<Dictionary<string, object>, object> key =
                x => x.ContainsKey(sc.ColumnName) ? x[sc.ColumnName] : null;

            if (ordered == null)
            {
                ordered = (sc.Direction == ListSortDirection.Ascending)
                    ? data.OrderBy(key, naturalComparer)
                    : data.OrderByDescending(key, naturalComparer);
            }
            else
            {
                ordered = (sc.Direction == ListSortDirection.Ascending)
                    ? ordered.ThenBy(key, naturalComparer)
                    : ordered.ThenByDescending(key, naturalComparer);
            }
        }
        return ordered.ToList();
    }
}
