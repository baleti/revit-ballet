using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Autodesk.Revit.DB;

#if REVIT2025 || REVIT2026
using Microsoft.Data.Sqlite;
using SQLitePCL;
#else
using System.Data.SQLite;
using SqliteConnection = System.Data.SQLite.SQLiteConnection;
using SqliteCommand = System.Data.SQLite.SQLiteCommand;
using SqliteDataReader = System.Data.SQLite.SQLiteDataReader;
#endif

namespace RevitBallet.Commands
{
    /// <summary>
    /// SQLite database helper for managing view history across sessions and documents.
    /// </summary>
    public static class LogViewChangesDatabase
    {
        private static string DatabasePath => PathHelper.GetRuntimeFilePath("LogViewChanges.sqlite");

        /// <summary>
        /// Initializes the database and creates tables if they don't exist.
        /// Migrates TEXT timestamps to INTEGER if needed.
        /// </summary>
        public static void InitializeDatabase()
        {
#if REVIT2025 || REVIT2026
            // Initialize SQLitePCL for Microsoft.Data.Sqlite
            SQLitePCL.Batteries.Init();
#endif

            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
            {
                connection.Open();

                // Check if table exists and has TEXT timestamp (migration needed)
                var checkCommand = connection.CreateCommand();
                checkCommand.CommandText = @"
                    SELECT sql FROM sqlite_master
                    WHERE type='table' AND name='ViewHistory' AND sql LIKE '%Timestamp TEXT%'
                ";
                bool needsMigration = checkCommand.ExecuteScalar() != null;

                if (needsMigration)
                {
                    // Migrate from TEXT to INTEGER timestamps
                    var migrateCommand = connection.CreateCommand();
                    migrateCommand.CommandText = @"
                        -- Create new table with INTEGER timestamp
                        CREATE TABLE ViewHistory_New (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            SessionId TEXT NOT NULL,
                            DocumentSessionId TEXT NOT NULL,
                            DocumentTitle TEXT NOT NULL,
                            DocumentPath TEXT,
                            ViewId INTEGER NOT NULL,
                            ViewTitle TEXT NOT NULL,
                            ViewType TEXT,
                            Timestamp INTEGER NOT NULL
                        );

                        -- Copy data, converting TEXT timestamps to Unix milliseconds
                        INSERT INTO ViewHistory_New
                        SELECT
                            Id, SessionId, DocumentSessionId, DocumentTitle, DocumentPath,
                            ViewId, ViewTitle, ViewType,
                            CAST((julianday(Timestamp) - 2440587.5) * 86400000 AS INTEGER) as Timestamp
                        FROM ViewHistory;

                        -- Drop old table and rename new one
                        DROP TABLE ViewHistory;
                        ALTER TABLE ViewHistory_New RENAME TO ViewHistory;

                        -- Recreate indexes
                        CREATE INDEX idx_session ON ViewHistory(SessionId);
                        CREATE INDEX idx_document_session ON ViewHistory(DocumentSessionId);
                        CREATE INDEX idx_timestamp ON ViewHistory(Timestamp);
                    ";
                    migrateCommand.ExecuteNonQuery();
                }
                else
                {
                    // Create table if it doesn't exist (new installation or already migrated)
                    var createCommand = connection.CreateCommand();
                    createCommand.CommandText = @"
                        CREATE TABLE IF NOT EXISTS ViewHistory (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            SessionId TEXT NOT NULL,
                            DocumentSessionId TEXT NOT NULL,
                            DocumentTitle TEXT NOT NULL,
                            DocumentPath TEXT,
                            ViewId INTEGER NOT NULL,
                            ViewTitle TEXT NOT NULL,
                            ViewType TEXT,
                            Timestamp INTEGER NOT NULL
                        );

                        CREATE INDEX IF NOT EXISTS idx_session ON ViewHistory(SessionId);
                        CREATE INDEX IF NOT EXISTS idx_document_session ON ViewHistory(DocumentSessionId);
                        CREATE INDEX IF NOT EXISTS idx_timestamp ON ViewHistory(Timestamp);
                    ";
                    createCommand.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Logs a view activation event.
        /// Updates the timestamp if the view was already activated in this session, otherwise inserts a new entry.
        /// </summary>
        public static void LogViewActivation(
            string sessionId,
            string documentSessionId,
            string documentTitle,
            string documentPath,
            ElementId viewId,
            string viewTitle,
            string viewType,
            DateTime timestamp)
        {
            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
            {
                connection.Open();

                // First, try to UPDATE existing entry
                // Composite key: (SessionId, DocumentTitle, ViewId)
                // SessionId = ProcessId, DocumentTitle distinguishes documents in same process
                var updateCommand = connection.CreateCommand();
                updateCommand.CommandText = @"
                    UPDATE ViewHistory
                    SET Timestamp = @Timestamp,
                        DocumentPath = @DocumentPath,
                        ViewTitle = @ViewTitle,
                        ViewType = @ViewType
                    WHERE SessionId = @SessionId
                      AND DocumentTitle = @DocumentTitle
                      AND ViewId = @ViewId
                ";

                updateCommand.Parameters.AddWithValue("@SessionId", sessionId);
                updateCommand.Parameters.AddWithValue("@DocumentTitle", documentTitle);
                updateCommand.Parameters.AddWithValue("@ViewId", viewId.AsLong());
                updateCommand.Parameters.AddWithValue("@DocumentPath", documentPath ?? "");
                updateCommand.Parameters.AddWithValue("@ViewTitle", viewTitle);
                updateCommand.Parameters.AddWithValue("@ViewType", viewType ?? "");
                updateCommand.Parameters.AddWithValue("@Timestamp", new DateTimeOffset(timestamp).ToUnixTimeMilliseconds());

                int rowsAffected = updateCommand.ExecuteNonQuery();

                // If no rows were updated, INSERT a new entry
                if (rowsAffected == 0)
                {
                    var insertCommand = connection.CreateCommand();
                    insertCommand.CommandText = @"
                        INSERT INTO ViewHistory
                        (SessionId, DocumentSessionId, DocumentTitle, DocumentPath, ViewId, ViewTitle, ViewType, Timestamp)
                        VALUES
                        (@SessionId, @DocumentSessionId, @DocumentTitle, @DocumentPath, @ViewId, @ViewTitle, @ViewType, @Timestamp)
                    ";

                    insertCommand.Parameters.AddWithValue("@SessionId", sessionId);
                    insertCommand.Parameters.AddWithValue("@DocumentSessionId", documentSessionId);
                    insertCommand.Parameters.AddWithValue("@DocumentTitle", documentTitle);
                    insertCommand.Parameters.AddWithValue("@DocumentPath", documentPath ?? "");
                    insertCommand.Parameters.AddWithValue("@ViewId", viewId.AsLong());
                    insertCommand.Parameters.AddWithValue("@ViewTitle", viewTitle);
                    insertCommand.Parameters.AddWithValue("@ViewType", viewType ?? "");
                    insertCommand.Parameters.AddWithValue("@Timestamp", new DateTimeOffset(timestamp).ToUnixTimeMilliseconds());

                    insertCommand.ExecuteNonQuery();
                }
            }
        }


        /// <summary>
        /// Gets view history for a specific document in the current session.
        /// Queries by (SessionId, DocumentTitle) composite key.
        /// Returns most recent views first.
        /// </summary>
        public static List<ViewHistoryEntry> GetViewHistoryForDocument(string sessionId, string documentTitle, int limit = 1000)
        {
            var entries = new List<ViewHistoryEntry>();

            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = @"
                    SELECT SessionId, DocumentTitle, DocumentPath, ViewId, ViewTitle, ViewType, Timestamp
                    FROM ViewHistory
                    WHERE SessionId = @SessionId
                      AND DocumentTitle = @DocumentTitle
                    ORDER BY Timestamp DESC
                    LIMIT @Limit
                ";

                command.Parameters.AddWithValue("@SessionId", sessionId);
                command.Parameters.AddWithValue("@DocumentTitle", documentTitle);
                command.Parameters.AddWithValue("@Limit", limit);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        entries.Add(new ViewHistoryEntry
                        {
                            SessionId = reader.GetString(0),
                            DocumentTitle = reader.GetString(1),
                            DocumentPath = reader.GetString(2),
                            ViewId = reader.GetInt64(3),
                            ViewTitle = reader.GetString(4),
                            ViewType = reader.GetString(5),
                            Timestamp = DateTimeOffset.FromUnixTimeMilliseconds(reader.GetInt64(6)).LocalDateTime
                        });
                    }
                }
            }

            return entries;
        }

        /// <summary>
        /// Gets view history for all documents in the current session.
        /// Returns most recent views first.
        /// </summary>
        public static List<ViewHistoryEntry> GetViewHistoryForSession(string sessionId, int limit = 1000)
        {
            var entries = new List<ViewHistoryEntry>();

            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = @"
                    SELECT SessionId, DocumentTitle, DocumentPath, ViewId, ViewTitle, ViewType, Timestamp
                    FROM ViewHistory
                    WHERE SessionId = @SessionId
                    ORDER BY Timestamp DESC
                    LIMIT @Limit
                ";

                command.Parameters.AddWithValue("@SessionId", sessionId);
                command.Parameters.AddWithValue("@Limit", limit);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        entries.Add(new ViewHistoryEntry
                        {
                            SessionId = reader.GetString(0),
                            DocumentTitle = reader.GetString(1),
                            DocumentPath = reader.GetString(2),
                            ViewId = reader.GetInt64(3),
                            ViewTitle = reader.GetString(4),
                            ViewType = reader.GetString(5),
                            Timestamp = DateTimeOffset.FromUnixTimeMilliseconds(reader.GetInt64(6)).LocalDateTime
                        });
                    }
                }
            }

            return entries;
        }

        /// <summary>
        /// Gets view history from all previous sessions for a specific document.
        /// Used by OpenLastViews to restore views from previous sessions.
        /// </summary>
        public static List<ViewHistoryEntry> GetPreviousSessionViews(string documentTitle, string excludeSessionId, int limit = 1000)
        {
            var entries = new List<ViewHistoryEntry>();

            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = @"
                    SELECT SessionId, DocumentTitle, DocumentPath, ViewId, ViewTitle, ViewType, Timestamp
                    FROM ViewHistory
                    WHERE DocumentTitle = @DocumentTitle
                      AND SessionId != @ExcludeSessionId
                    ORDER BY Timestamp DESC
                    LIMIT @Limit
                ";

                command.Parameters.AddWithValue("@DocumentTitle", documentTitle);
                command.Parameters.AddWithValue("@ExcludeSessionId", excludeSessionId);
                command.Parameters.AddWithValue("@Limit", limit);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        entries.Add(new ViewHistoryEntry
                        {
                            SessionId = reader.GetString(0),
                            DocumentTitle = reader.GetString(1),
                            DocumentPath = reader.GetString(2),
                            ViewId = reader.GetInt64(3),
                            ViewTitle = reader.GetString(4),
                            ViewType = reader.GetString(5),
                            Timestamp = DateTimeOffset.FromUnixTimeMilliseconds(reader.GetInt64(6)).LocalDateTime
                        });
                    }
                }
            }

            return entries;
        }

        /// <summary>
        /// Gets ALL view history entries for a specific document across ALL sessions.
        /// Returns each view activation as a separate entry (views can repeat across sessions).
        /// Used by OpenPreviousViewsIn commands to show complete history.
        /// </summary>
        public static List<ViewHistoryEntry> GetAllViewHistoryForDocument(string documentTitle, int limit = 1000)
        {
            var entries = new List<ViewHistoryEntry>();

            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = @"
                    SELECT SessionId, DocumentTitle, DocumentPath, ViewId, ViewTitle, ViewType, Timestamp
                    FROM ViewHistory
                    WHERE DocumentTitle = @DocumentTitle
                    ORDER BY Timestamp DESC
                    LIMIT @Limit
                ";

                command.Parameters.AddWithValue("@DocumentTitle", documentTitle);
                command.Parameters.AddWithValue("@Limit", limit);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        entries.Add(new ViewHistoryEntry
                        {
                            SessionId = reader.GetString(0),
                            DocumentTitle = reader.GetString(1),
                            DocumentPath = reader.GetString(2),
                            ViewId = reader.GetInt64(3),
                            ViewTitle = reader.GetString(4),
                            ViewType = reader.GetString(5),
                            Timestamp = DateTimeOffset.FromUnixTimeMilliseconds(reader.GetInt64(6)).LocalDateTime
                        });
                    }
                }
            }

            return entries;
        }

        /// <summary>
        /// Gets ALL view history entries across ALL sessions and ALL documents.
        /// Returns each view activation as a separate entry (views can repeat across sessions).
        /// Used by OpenPreviousViewsInSession to show complete cross-document history.
        /// </summary>
        public static List<ViewHistoryEntry> GetAllViewHistoryForAllDocuments(int limit = 1000)
        {
            var entries = new List<ViewHistoryEntry>();

            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = @"
                    SELECT SessionId, DocumentTitle, DocumentPath, ViewId, ViewTitle, ViewType, Timestamp
                    FROM ViewHistory
                    ORDER BY Timestamp DESC
                    LIMIT @Limit
                ";

                command.Parameters.AddWithValue("@Limit", limit);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        entries.Add(new ViewHistoryEntry
                        {
                            SessionId = reader.GetString(0),
                            DocumentTitle = reader.GetString(1),
                            DocumentPath = reader.GetString(2),
                            ViewId = reader.GetInt64(3),
                            ViewTitle = reader.GetString(4),
                            ViewType = reader.GetString(5),
                            Timestamp = DateTimeOffset.FromUnixTimeMilliseconds(reader.GetInt64(6)).LocalDateTime
                        });
                    }
                }
            }

            return entries;
        }

        /// <summary>
        /// Removes a view from the current document session history.
        /// Used when closing views.
        /// </summary>
        public static void RemoveViewFromHistory(string sessionId, string documentTitle, string viewTitle)
        {
            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = @"
                    DELETE FROM ViewHistory
                    WHERE SessionId = @SessionId
                      AND DocumentTitle = @DocumentTitle
                      AND ViewTitle = @ViewTitle
                ";

                command.Parameters.AddWithValue("@SessionId", sessionId);
                command.Parameters.AddWithValue("@DocumentTitle", documentTitle);
                command.Parameters.AddWithValue("@ViewTitle", viewTitle);

                command.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Cleans up old history entries (optional - for maintenance).
        /// Removes entries older than the specified number of days.
        /// </summary>
        public static void CleanupOldEntries(int daysToKeep = 30)
        {
            using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = @"
                    DELETE FROM ViewHistory
                    WHERE Timestamp < @CutoffDate
                ";

                var cutoffDate = new DateTimeOffset(DateTime.Now.AddDays(-daysToKeep)).ToUnixTimeMilliseconds();
                command.Parameters.AddWithValue("@CutoffDate", cutoffDate);

                command.ExecuteNonQuery();
            }
        }
    }

    /// <summary>
    /// Represents a single view history entry.
    /// </summary>
    public class ViewHistoryEntry
    {
        public string SessionId { get; set; }
        public string DocumentTitle { get; set; }
        public string DocumentPath { get; set; }
        public long ViewId { get; set; }
        public string ViewTitle { get; set; }
        public string ViewType { get; set; }
        public DateTime Timestamp { get; set; }
    }
}
