using System;
using System.Text.RegularExpressions;

namespace RevitBallet.Core.Query
{
    /// <summary>
    /// The one value comparator behind `/sort`, comparison filters (`//col >4`) and any column-header sort a host offers.
    /// Numbers compare numerically, age buckets ("30s" "5m" "3h" "2d") by their seconds, everything else naturally
    /// (so "A2" sorts before "A10"). Ages carry the direction trap: ascending means oldest first.
    /// </summary>
    public static class FieldCompare
    {
        private static readonly Regex AgeRx = new Regex(@"^\d+[smhd]$", RegexOptions.Compiled);
        private static readonly NaturalComparer Natural = new NaturalComparer();

        public static bool TryAgeSeconds(string s, out double seconds)
        {
            seconds = 0;
            if (s == null) return false;
            s = s.Trim();
            if (!AgeRx.IsMatch(s)) return false;
            double n = double.Parse(s.Substring(0, s.Length - 1), System.Globalization.CultureInfo.InvariantCulture);
            switch (s[s.Length - 1])
            {
                case 's': seconds = n; break;
                case 'm': seconds = n * 60; break;
                case 'h': seconds = n * 3600; break;
                default: seconds = n * 86400; break;
            }
            return true;
        }

        public static bool TryNumber(string s, out double n)
        {
            return TextMatch.TryParseDouble(s, out n);
        }

        /// <summary>Compares two cell values for ordering; `descending` already folded in, the age trap applied.</summary>
        public static int Compare(string a, string b, bool descending)
        {
            double sa, sb;
            int c;
            bool ageA = TryAgeSeconds(a, out sa), ageB = TryAgeSeconds(b, out sb);
            if (ageA && ageB)
            {
                c = -sa.CompareTo(sb);                 // trap: the stored value is an age, so "ascending" (oldest first) is larger first
            }
            else
            {
                c = Natural.Compare(a, b);
            }
            return descending ? -c : c;
        }

        /// <summary>Evaluates `value OP operand` for a comparison filter. Only two numbers, or two age buckets, are comparable;
        /// anything else never matches (it is not read as literal text). For ages the operator compares the raw stored age.</summary>
        public static bool TryCompareFilter(string cell, string op, string operand, out bool matches)
        {
            matches = false;
            double ca, co;
            bool ageCell = TryAgeSeconds(cell, out ca), ageOp = TryAgeSeconds(operand, out co);
            if (!(ageCell && ageOp))
            {
                double na, no;
                if (!TextMatch.TryParseDouble(cell, out na) || !TextMatch.TryParseDouble(operand, out no)) return false;
                ca = na; co = no;
            }
            switch (op)
            {
                case ">": matches = ca > co; break;
                case "<": matches = ca < co; break;
                case ">=": matches = ca >= co; break;
                case "<=": matches = ca <= co; break;
                default: return false;
            }
            return true;
        }
    }
}
