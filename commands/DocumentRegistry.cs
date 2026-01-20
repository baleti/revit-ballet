using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

#if REVIT2025 || REVIT2026
using Microsoft.Data.Sqlite;
#else
using System.Data.SQLite;
#endif

namespace RevitBallet.Commands
{
    /// <summary>
    /// Represents a document registered in the Revit Ballet network.
    /// </summary>
    public class DocumentEntry
    {
        public string DocumentTitle { get; set; }
        public string DocumentPath { get; set; }
        public string SessionId { get; set; }
        public int Port { get; set; }
        public string Hostname { get; set; }
        public int ProcessId { get; set; }
        public DateTime RegisteredAt { get; set; }
        public DateTime LastHeartbeat { get; set; }
        public DateTime? LastSync { get; set; }
        public DateTime? LastTransaction { get; set; }
    }

    /// <summary>
    /// SQLite-based document registry for peer-to-peer network coordination.
    /// Tracks all open documents across all Revit sessions.
    /// High-performance replacement for CSV-based registry to handle frequent transaction updates.
    /// </summary>
    public static class DocumentRegistry
    {
        private static readonly string DatabasePath = Path.Combine(
            PathHelper.RuntimeDirectory,
            "documents.sqlite"
        );

        private static readonly object lockObject = new object();

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

                // Enable WAL mode for better concurrency
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "PRAGMA journal_mode=WAL;";
                    command.ExecuteNonQuery();
                }

                string createTableSql = @"
                    CREATE TABLE IF NOT EXISTS documents (
                        session_id TEXT NOT NULL,
                        document_path TEXT NOT NULL,
                        document_title TEXT NOT NULL,
                        port INTEGER NOT NULL,
                        hostname TEXT NOT NULL,
                        process_id INTEGER NOT NULL,
                        registered_at TEXT NOT NULL,
                        last_heartbeat TEXT NOT NULL,
                        last_sync TEXT,
                        last_transaction TEXT,
                        PRIMARY KEY (session_id, document_path)
                    );

                    CREATE INDEX IF NOT EXISTS idx_session_id ON documents(session_id);
                    CREATE INDEX IF NOT EXISTS idx_last_heartbeat ON documents(last_heartbeat);
                    CREATE INDEX IF NOT EXISTS idx_hostname ON documents(hostname);
                ";

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = createTableSql;
                    command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Register or update a document entry in the registry.
        /// Uses UPSERT (INSERT OR REPLACE) for atomic updates.
        /// </summary>
        public static void UpsertDocument(DocumentEntry entry)
        {
            lock (lockObject)
            {
                Directory.CreateDirectory(PathHelper.RuntimeDirectory);
                InitializeDatabase();

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
                            INSERT OR REPLACE INTO documents
                            (session_id, document_path, document_title, port, hostname, process_id,
                             registered_at, last_heartbeat, last_sync, last_transaction)
                            VALUES
                            (@sessionId, @documentPath, @documentTitle, @port, @hostname, @processId,
                             @registeredAt, @lastHeartbeat, @lastSync, @lastTransaction)
                        ";

                        AddParameter(command, "@sessionId", entry.SessionId);
                        AddParameter(command, "@documentPath", entry.DocumentPath ?? "");
                        AddParameter(command, "@documentTitle", entry.DocumentTitle ?? "");
                        AddParameter(command, "@port", entry.Port);
                        AddParameter(command, "@hostname", entry.Hostname ?? "");
                        AddParameter(command, "@processId", entry.ProcessId);
                        AddParameter(command, "@registeredAt", entry.RegisteredAt.ToString("o"));
                        AddParameter(command, "@lastHeartbeat", entry.LastHeartbeat.ToString("o"));
                        AddParameter(command, "@lastSync", entry.LastSync?.ToString("o"));
                        AddParameter(command, "@lastTransaction", entry.LastTransaction?.ToString("o"));

                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        /// <summary>
        /// Bulk update multiple documents (used during heartbeat).
        /// More efficient than individual UpsertDocument calls.
        /// </summary>
        public static void UpsertDocuments(List<DocumentEntry> entries)
        {
            if (entries == null || entries.Count == 0)
                return;

            lock (lockObject)
            {
                Directory.CreateDirectory(PathHelper.RuntimeDirectory);
                InitializeDatabase();

#if REVIT2025 || REVIT2026
                using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
#else
                using (var connection = new SQLiteConnection($"Data Source={DatabasePath};Version=3;"))
#endif
                {
                    connection.Open();

                    using (var transaction = connection.BeginTransaction())
                    {
                        foreach (var entry in entries)
                        {
                            using (var command = connection.CreateCommand())
                            {
                                command.CommandText = @"
                                    INSERT OR REPLACE INTO documents
                                    (session_id, document_path, document_title, port, hostname, process_id,
                                     registered_at, last_heartbeat, last_sync, last_transaction)
                                    VALUES
                                    (@sessionId, @documentPath, @documentTitle, @port, @hostname, @processId,
                                     @registeredAt, @lastHeartbeat, @lastSync, @lastTransaction)
                                ";

                                AddParameter(command, "@sessionId", entry.SessionId);
                                AddParameter(command, "@documentPath", entry.DocumentPath ?? "");
                                AddParameter(command, "@documentTitle", entry.DocumentTitle ?? "");
                                AddParameter(command, "@port", entry.Port);
                                AddParameter(command, "@hostname", entry.Hostname ?? "");
                                AddParameter(command, "@processId", entry.ProcessId);
                                AddParameter(command, "@registeredAt", entry.RegisteredAt.ToString("o"));
                                AddParameter(command, "@lastHeartbeat", entry.LastHeartbeat.ToString("o"));
                                AddParameter(command, "@lastSync", entry.LastSync?.ToString("o"));
                                AddParameter(command, "@lastTransaction", entry.LastTransaction?.ToString("o"));

                                command.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                }
            }
        }

        /// <summary>
        /// Update only the LastSync timestamp for a specific document.
        /// Optimized for frequent updates from DocumentSynchronized event.
        /// </summary>
        public static void UpdateLastSync(string sessionId, string documentPath, DateTime timestamp)
        {
            lock (lockObject)
            {
                Directory.CreateDirectory(PathHelper.RuntimeDirectory);
                InitializeDatabase();

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
                            UPDATE documents
                            SET last_sync = @timestamp
                            WHERE session_id = @sessionId
                              AND (document_path = @documentPath OR (document_path = '' AND document_title = @documentPath))
                        ";

                        AddParameter(command, "@timestamp", timestamp.ToString("o"));
                        AddParameter(command, "@sessionId", sessionId);
                        AddParameter(command, "@documentPath", documentPath ?? "");

                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        /// <summary>
        /// Update only the LastTransaction timestamp for a specific document.
        /// Optimized for very frequent updates from DocumentChanged event.
        /// </summary>
        public static void UpdateLastTransaction(string sessionId, string documentPath, DateTime timestamp)
        {
            lock (lockObject)
            {
                Directory.CreateDirectory(PathHelper.RuntimeDirectory);
                InitializeDatabase();

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
                            UPDATE documents
                            SET last_transaction = @timestamp
                            WHERE session_id = @sessionId
                              AND (document_path = @documentPath OR (document_path = '' AND document_title = @documentPath))
                        ";

                        AddParameter(command, "@timestamp", timestamp.ToString("o"));
                        AddParameter(command, "@sessionId", sessionId);
                        AddParameter(command, "@documentPath", documentPath ?? "");

                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        /// <summary>
        /// Get all documents from the registry.
        /// </summary>
        public static List<DocumentEntry> GetAllDocuments()
        {
            lock (lockObject)
            {
                var entries = new List<DocumentEntry>();

                if (!File.Exists(DatabasePath))
                    return entries;

                InitializeDatabase();

#if REVIT2025 || REVIT2026
                using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
#else
                using (var connection = new SQLiteConnection($"Data Source={DatabasePath};Version=3;"))
#endif
                {
                    connection.Open();

                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = "SELECT * FROM documents ORDER BY document_title";

                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                entries.Add(new DocumentEntry
                                {
                                    SessionId = reader["session_id"].ToString(),
                                    DocumentPath = reader["document_path"].ToString(),
                                    DocumentTitle = reader["document_title"].ToString(),
                                    Port = Convert.ToInt32(reader["port"]),
                                    Hostname = reader["hostname"].ToString(),
                                    ProcessId = Convert.ToInt32(reader["process_id"]),
                                    RegisteredAt = DateTime.Parse(reader["registered_at"].ToString()),
                                    LastHeartbeat = DateTime.Parse(reader["last_heartbeat"].ToString()),
                                    LastSync = reader["last_sync"] != DBNull.Value ? DateTime.Parse(reader["last_sync"].ToString()) : (DateTime?)null,
                                    LastTransaction = reader["last_transaction"] != DBNull.Value ? DateTime.Parse(reader["last_transaction"].ToString()) : (DateTime?)null
                                });
                            }
                        }
                    }
                }

                return entries;
            }
        }

        /// <summary>
        /// Get active documents (heartbeat within last 2 minutes).
        /// </summary>
        public static List<DocumentEntry> GetActiveDocuments()
        {
            lock (lockObject)
            {
                var entries = new List<DocumentEntry>();

                if (!File.Exists(DatabasePath))
                    return entries;

                InitializeDatabase();

                var cutoffTime = DateTime.Now.AddSeconds(-120);

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
                            SELECT * FROM documents
                            WHERE last_heartbeat >= @cutoffTime
                            ORDER BY document_title
                        ";

                        AddParameter(command, "@cutoffTime", cutoffTime.ToString("o"));

                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                entries.Add(new DocumentEntry
                                {
                                    SessionId = reader["session_id"].ToString(),
                                    DocumentPath = reader["document_path"].ToString(),
                                    DocumentTitle = reader["document_title"].ToString(),
                                    Port = Convert.ToInt32(reader["port"]),
                                    Hostname = reader["hostname"].ToString(),
                                    ProcessId = Convert.ToInt32(reader["process_id"]),
                                    RegisteredAt = DateTime.Parse(reader["registered_at"].ToString()),
                                    LastHeartbeat = DateTime.Parse(reader["last_heartbeat"].ToString()),
                                    LastSync = reader["last_sync"] != DBNull.Value ? DateTime.Parse(reader["last_sync"].ToString()) : (DateTime?)null,
                                    LastTransaction = reader["last_transaction"] != DBNull.Value ? DateTime.Parse(reader["last_transaction"].ToString()) : (DateTime?)null
                                });
                            }
                        }
                    }
                }

                return entries;
            }
        }

        /// <summary>
        /// Remove all documents for a specific session.
        /// </summary>
        public static void RemoveSession(string sessionId)
        {
            lock (lockObject)
            {
                if (!File.Exists(DatabasePath))
                    return;

                InitializeDatabase();

#if REVIT2025 || REVIT2026
                using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
#else
                using (var connection = new SQLiteConnection($"Data Source={DatabasePath};Version=3;"))
#endif
                {
                    connection.Open();

                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = "DELETE FROM documents WHERE session_id = @sessionId";
                        AddParameter(command, "@sessionId", sessionId);
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        /// <summary>
        /// Remove stale documents (no heartbeat for more than 2 minutes).
        /// </summary>
        public static void RemoveStaleDocuments()
        {
            lock (lockObject)
            {
                if (!File.Exists(DatabasePath))
                    return;

                InitializeDatabase();

                var cutoffTime = DateTime.Now.AddSeconds(-120);

#if REVIT2025 || REVIT2026
                using (var connection = new SqliteConnection($"Data Source={DatabasePath}"))
#else
                using (var connection = new SQLiteConnection($"Data Source={DatabasePath};Version=3;"))
#endif
                {
                    connection.Open();

                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = "DELETE FROM documents WHERE last_heartbeat < @cutoffTime";
                        AddParameter(command, "@cutoffTime", cutoffTime.ToString("o"));
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        /// <summary>
        /// Helper to add parameters to SQLite command (handles version differences).
        /// </summary>
        private static void AddParameter(
#if REVIT2025 || REVIT2026
            SqliteCommand command,
#else
            SQLiteCommand command,
#endif
            string name, object value)
        {
            var parameter = command.CreateParameter();
            parameter.ParameterName = name;
            parameter.Value = value ?? DBNull.Value;
            command.Parameters.Add(parameter);
        }
    }
}
