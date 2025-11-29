using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Search Query History Management
    // ──────────────────────────────────────────────────────────────

    /// <summary>Manages persistence of search query history for DataGrid</summary>
    public static class SearchQueryHistory
    {
        private static string GetHistoryDirectory()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string historyDir = Path.Combine(appData, "revit-ballet", "runtime", "SearchboxQueries");

            if (!Directory.Exists(historyDir))
            {
                Directory.CreateDirectory(historyDir);
            }

            return historyDir;
        }

        private static string GetHistoryFilePath(string commandName)
        {
            if (string.IsNullOrWhiteSpace(commandName))
            {
                commandName = "default";
            }

            // Sanitize command name for file system
            string safeCommandName = string.Join("_", commandName.Split(Path.GetInvalidFileNameChars()));

            return Path.Combine(GetHistoryDirectory(), safeCommandName);
        }

        /// <summary>Records a search query for a specific command (only if non-empty and not duplicate of last entry)</summary>
        public static void RecordQuery(string commandName, string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return; // Don't record empty queries
            }

            query = query.Trim();
            string filePath = GetHistoryFilePath(commandName);

            try
            {
                // Check if this is a duplicate of the last query
                string lastQuery = null;
                if (File.Exists(filePath))
                {
                    var lines = File.ReadAllLines(filePath);
                    if (lines.Length > 0)
                    {
                        lastQuery = lines[lines.Length - 1].Trim();
                    }
                }

                // Don't record if it's the same as the last query
                if (!string.IsNullOrEmpty(lastQuery) && string.Equals(lastQuery, query, StringComparison.Ordinal))
                {
                    return;
                }

                // Append query to file
                File.AppendAllText(filePath, query + Environment.NewLine);
            }
            catch (Exception)
            {
                // Silently ignore errors - search history is non-critical
            }
        }

        /// <summary>Retrieves all search queries for a specific command (most recent last)</summary>
        public static List<string> GetQueryHistory(string commandName)
        {
            string filePath = GetHistoryFilePath(commandName);
            var history = new List<string>();

            try
            {
                if (File.Exists(filePath))
                {
                    var lines = File.ReadAllLines(filePath);

                    // Remove duplicates while preserving order (keep most recent occurrence)
                    var uniqueQueries = new Dictionary<string, int>();
                    for (int i = 0; i < lines.Length; i++)
                    {
                        string query = lines[i].Trim();
                        if (!string.IsNullOrWhiteSpace(query))
                        {
                            uniqueQueries[query] = i; // Update to latest occurrence
                        }
                    }

                    // Sort by original position and return
                    history = uniqueQueries
                        .OrderBy(kvp => kvp.Value)
                        .Select(kvp => kvp.Key)
                        .ToList();
                }
            }
            catch (Exception)
            {
                // Silently ignore errors - return empty list
            }

            return history;
        }

        /// <summary>Clears all search history for a specific command</summary>
        public static void ClearHistory(string commandName)
        {
            string filePath = GetHistoryFilePath(commandName);

            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch (Exception)
            {
                // Silently ignore errors
            }
        }
    }
}
