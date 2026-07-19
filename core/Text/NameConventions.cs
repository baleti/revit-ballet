using System.Text;
using System.Text.RegularExpressions;

namespace RevitBallet.Core
{
    /// <summary>
    /// Naming and header-formatting conventions shared across the UI.
    /// </summary>
    public static class NameConventions
    {
        /// <summary>
        /// Converts PascalCase to kebab-case (e.g., "SwitchView" -> "switch-view")
        /// </summary>
        public static string ConvertToKebabCase(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return input;

            // Insert hyphen before uppercase letters (except first character)
            var kebabCase = Regex.Replace(input, "(?<!^)([A-Z])", "-$1");

            // Convert to lowercase
            return kebabCase.ToLower();
        }

        /// <summary>
        /// Formats column headers: replaces underscores with spaces, converts PascalCase to lowercase with spaces
        /// </summary>
        public static string FormatColumnHeader(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
                return columnName;

            var result = new StringBuilder();

            for (int i = 0; i < columnName.Length; i++)
            {
                char c = columnName[i];

                // Replace underscores with spaces
                if (c == '_')
                {
                    result.Append(' ');
                }
                // Add space before uppercase letters (except at start)
                else if (i > 0 && char.IsUpper(c) && !char.IsUpper(columnName[i - 1]))
                {
                    result.Append(' ');
                    result.Append(char.ToLower(c));
                }
                // Convert to lowercase
                else
                {
                    result.Append(char.ToLower(c));
                }
            }

            return result.ToString();
        }
    }
}
