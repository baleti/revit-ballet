using System;
using System.Collections.Generic;

namespace RevitBallet.Core
{
    /// <summary>A string comparer that sorts "A2" before "A10" and handles mixed numeric/text data.</summary>
    public sealed class NaturalComparer : IComparer<object>
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

            // If both are non-numeric, compare naturally so embedded numbers
            // sort numerically ("A9" before "A10")
            if (s1IsNonNumeric && s2IsNonNumeric)
            {
                return CompareNatural(s1, s2);
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
                    // Case-insensitive, preserving the previous comparer's behavior
                    int cmp = char.ToLowerInvariant(a[i]).CompareTo(char.ToLowerInvariant(b[j]));
                    if (cmp != 0) return cmp;
                    i++;
                    j++;
                }
            }
            return a.Length.CompareTo(b.Length);
        }
    }
}
