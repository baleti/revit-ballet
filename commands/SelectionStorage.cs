using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;

#if REVIT2025 || REVIT2026
using Microsoft.Data.Sqlite;
#else
using System.Data.SQLite;
#endif

namespace RevitBallet.Commands
{
    /// <summary>
    /// Represents a selected element with document and unique identifier information.
    /// Used for cross-document selection (Session and Network scopes).
    /// </summary>
    public class SelectionItem
    {
        /// <summary>Document title (e.g., "Project1.rvt")</summary>
        public string DocumentTitle { get; set; }

        /// <summary>Document path (full path when available)</summary>
        public string DocumentPath { get; set; }

        /// <summary>Element's UniqueId (stable across sessions)</summary>
        public string UniqueId { get; set; }

        /// <summary>Optional: Element's current ElementId (for convenience, but not stable)</summary>
        public int ElementIdValue { get; set; }

        /// <summary>Session ID that created this selection</summary>
        public string SessionId { get; set; }
    }

    /// <summary>
    /// SQLite-based selection storage for cross-document operations.
    /// Native Revit Selection API only works within the active document, so we use
    /// external storage to support Session and Network scopes.
    /// </summary>
    public static class SelectionStorage
    {
        private static readonly string DatabasePath = Path.Combine(
            PathHelper.RuntimeDirectory,
            "selection.sqlite"
        );

        private static string GetSessionId()
        {
            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            return $"revit_{processId}_{DateTime.Now:yyyyMMdd_HHmmss}";
        }

        /// <summary>
        /// Initialize the database schema. Called automatically on first use.
        /// </summary>
        private static void InitializeDatabase()
        {
#if REVIT2025 || REVIT2026
            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
#else
            using (var connection = new SQLiteConnection($"Data Source={DatabasePath};Version=3;"))
#endif
            {
                connection.Open();

                string createTableSql = @"
                    CREATE TABLE IF NOT EXISTS selection (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        document_title TEXT NOT NULL,
                        document_path TEXT,
                        unique_ids TEXT NOT NULL,
                        session_id TEXT NOT NULL,
                        created_at TEXT NOT NULL
                    );

                    CREATE INDEX IF NOT EXISTS idx_document_title ON selection(document_title);
                    CREATE INDEX IF NOT EXISTS idx_session_id ON selection(session_id);
                ";

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = createTableSql;
                    command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Save selection items to database. Replaces existing selection.
        /// Groups elements by document to minimize storage.
        /// </summary>
        public static void SaveSelection(List<SelectionItem> items)
        {
            // Ensure database exists
            Directory.CreateDirectory(PathHelper.RuntimeDirectory);
            InitializeDatabase();

            string sessionId = GetSessionId();

#if REVIT2025 || REVIT2026
            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
#else
            using (var connection = new SQLiteConnection($"Data Source={DatabasePath};Version=3;"))
#endif
            {
                connection.Open();

                // Clear existing selection
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "DELETE FROM selection";
                    command.ExecuteNonQuery();
                }

                // Insert new selection grouped by document
                if (items != null && items.Count > 0)
                {
                    // Group items by document
                    var groupedByDocument = items.GroupBy(item => new
                    {
                        DocumentTitle = item.DocumentTitle ?? "",
                        DocumentPath = item.DocumentPath ?? ""
                    });

                    using (var transaction = connection.BeginTransaction())
                    {
                        foreach (var docGroup in groupedByDocument)
                        {
                            // Collect all UniqueIds for this document
                            var uniqueIds = docGroup.Select(item => item.UniqueId ?? "").Where(id => !string.IsNullOrEmpty(id));
                            string uniqueIdsCommaSeparated = string.Join(",", uniqueIds);

                            if (string.IsNullOrEmpty(uniqueIdsCommaSeparated))
                                continue;

                            using (var command = connection.CreateCommand())
                            {
                                command.CommandText = @"
                                    INSERT INTO selection (document_title, document_path, unique_ids, session_id, created_at)
                                    VALUES (@documentTitle, @documentPath, @uniqueIds, @sessionId, @createdAt)
                                ";

                                AddParameter(command, "@documentTitle", docGroup.Key.DocumentTitle);
                                AddParameter(command, "@documentPath", docGroup.Key.DocumentPath);
                                AddParameter(command, "@uniqueIds", uniqueIdsCommaSeparated);
                                AddParameter(command, "@sessionId", sessionId);
                                AddParameter(command, "@createdAt", DateTime.Now.ToString("o"));

                                command.ExecuteNonQuery();
                            }
                        }
                        transaction.Commit();
                    }
                }
            }
        }

        /// <summary>
        /// Add selection items to existing selection (append mode).
        /// </summary>
        public static void AddToSelection(List<SelectionItem> items)
        {
            if (items == null || items.Count == 0)
                return;

            // Load existing selection, merge with new items, and save
            var existingSelection = LoadSelection();
            var existingUniqueIds = new HashSet<string>(existingSelection.Select(s => $"{s.DocumentTitle}|{s.UniqueId}"));

            // Add new items that don't already exist
            foreach (var item in items)
            {
                string key = $"{item.DocumentTitle}|{item.UniqueId}";
                if (!existingUniqueIds.Contains(key))
                {
                    existingSelection.Add(item);
                }
            }

            // Save merged selection
            SaveSelection(existingSelection);
        }

        /// <summary>
        /// Load all selection items from database.
        /// Expands comma-separated UniqueIds into individual SelectionItems.
        /// </summary>
        public static List<SelectionItem> LoadSelection()
        {
            var items = new List<SelectionItem>();

            if (!File.Exists(DatabasePath))
                return items;

#if REVIT2025 || REVIT2026
            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
#else
            using (var connection = new SQLiteConnection($"Data Source={DatabasePath};Version=3;"))
#endif
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = @"
                        SELECT document_title, document_path, unique_ids, session_id
                        FROM selection
                        ORDER BY created_at
                    ";

                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string documentTitle = reader.GetString(0);
                            string documentPath = reader.IsDBNull(1) ? "" : reader.GetString(1);
                            string uniqueIdsCommaSeparated = reader.GetString(2);
                            string sessionId = reader.GetString(3);

                            // Split comma-separated UniqueIds and create SelectionItem for each
                            var uniqueIds = uniqueIdsCommaSeparated.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string uniqueId in uniqueIds)
                            {
                                items.Add(new SelectionItem
                                {
                                    DocumentTitle = documentTitle,
                                    DocumentPath = documentPath,
                                    UniqueId = uniqueId.Trim(),
                                    ElementIdValue = 0, // Not stored in optimized schema
                                    SessionId = sessionId
                                });
                            }
                        }
                    }
                }
            }

            return items;
        }

        /// <summary>
        /// Load selection items only for currently open documents.
        /// </summary>
        public static List<SelectionItem> LoadSelectionForOpenDocuments(Autodesk.Revit.ApplicationServices.Application app)
        {
            var allItems = LoadSelection();
            var openDocumentTitles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Get all open document titles
            foreach (Document doc in app.Documents)
            {
                openDocumentTitles.Add(doc.Title);
            }

            // Filter to only open documents
            var filteredItems = new List<SelectionItem>();
            foreach (var item in allItems)
            {
                if (openDocumentTitles.Contains(item.DocumentTitle))
                {
                    filteredItems.Add(item);
                }
            }

            return filteredItems;
        }

        /// <summary>
        /// Clear all selection items from database.
        /// </summary>
        public static void ClearSelection()
        {
            if (!File.Exists(DatabasePath))
                return;

#if REVIT2025 || REVIT2026
            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
#else
            using (var connection = new SQLiteConnection($"Data Source={DatabasePath};Version=3;"))
#endif
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "DELETE FROM selection";
                    command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Helper to add parameter with correct type based on SQLite library.
        /// </summary>
#if REVIT2025 || REVIT2026
        private static void AddParameter(SqliteCommand command, string name, object value)
        {
            var parameter = command.CreateParameter();
            parameter.ParameterName = name;
            parameter.Value = value ?? DBNull.Value;
            command.Parameters.Add(parameter);
        }
#else
        private static void AddParameter(SQLiteCommand command, string name, object value)
        {
            var parameter = command.CreateParameter();
            parameter.ParameterName = name;
            parameter.Value = value ?? DBNull.Value;
            command.Parameters.Add(parameter);
        }
#endif
    }
}
