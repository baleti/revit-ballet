using System.Text.RegularExpressions;

namespace RevitBallet.Core
{
    /// <summary>
    /// String matching helpers shared by the filter parser and row matcher.
    /// </summary>
    public static class TextMatch
    {
        public static string StripQuotes(string s)
        {
            return s.StartsWith("\"") && s.EndsWith("\"") && s.Length > 1
                ? s.Substring(1, s.Length - 2)
                : s;
        }

        /// <summary>Check if a string contains glob wildcards</summary>
        public static bool ContainsGlobWildcards(string pattern)
        {
            return pattern != null && pattern.Contains("*");
        }

        /// <summary>Convert glob pattern to regex pattern</summary>
        public static string GlobToRegexPattern(string globPattern)
        {
            // Escape special regex characters except *
            string escaped = Regex.Escape(globPattern).Replace("\\*", ".*");
            return "^" + escaped + "$";
        }

        /// <summary>Check if a value matches a glob pattern</summary>
        public static bool MatchesGlobPattern(string value, string pattern)
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

        /// <summary>Try to parse a string as a double, handling common formats</summary>
        public static bool TryParseDouble(string s, out double result)
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
}
