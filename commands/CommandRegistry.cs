using System.Collections.Generic;

namespace RevitBallet.Commands
{
    /// <summary>
    /// Central registry for command display names
    /// Maps class names to user-friendly display names
    /// </summary>
    public static class CommandRegistry
    {
        private static readonly Dictionary<string, string> _commandNames = new Dictionary<string, string>
        {
            // Add friendly names for commands here
            ["OpenRvtFilesInNewSessions"] = "Open Recent Revit Files",
            // Add more mappings as needed
        };

        /// <summary>
        /// Get friendly display name for a command class
        /// </summary>
        public static string GetDisplayName(string className)
        {
            if (_commandNames.TryGetValue(className, out string displayName))
            {
                return displayName;
            }

            // Fall back to class name with spaces inserted before capital letters
            return InsertSpacesBeforeCapitals(className);
        }

        /// <summary>
        /// Get class name from display name
        /// </summary>
        public static string GetClassName(string displayName)
        {
            foreach (var kvp in _commandNames)
            {
                if (kvp.Value == displayName)
                {
                    return kvp.Key;
                }
            }

            // Fall back to removing spaces
            return displayName.Replace(" ", "");
        }

        /// <summary>
        /// Check if a class should be shown in network mode
        /// </summary>
        public static bool IsNetworkCommand(string className)
        {
            // Commands that work without Revit API context
            var networkCommands = new HashSet<string>
            {
                "OpenRvtFilesInNewSessions"
            };

            return networkCommands.Contains(className);
        }

        private static string InsertSpacesBeforeCapitals(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            var result = new System.Text.StringBuilder();
            result.Append(text[0]);

            for (int i = 1; i < text.Length; i++)
            {
                if (char.IsUpper(text[i]) && !char.IsUpper(text[i - 1]))
                {
                    result.Append(' ');
                }
                result.Append(text[i]);
            }

            return result.ToString();
        }
    }
}
