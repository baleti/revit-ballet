using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;
using Autodesk.Revit.UI;

#if REVIT2025 || REVIT2026
using Microsoft.Data.Sqlite;
#else
using System.Data.SQLite;
#endif

namespace RevitBallet.Commands
{
    public static class CommandUsageTracker
    {
        private struct UsageRecord
        {
            public string CommandName;
            public string ExecutedAt;
            public string Result;
            public long DurationMs;
        }

        private static readonly ConcurrentQueue<UsageRecord> _queue = new ConcurrentQueue<UsageRecord>();

        private static readonly string _dbPath = Path.Combine(
            PathHelper.RuntimeDirectory, "command-usage.sqlite");

        private static volatile bool _dbReady = false;

        static CommandUsageTracker()
        {
            var thread = new Thread(BackgroundWriter)
            {
                IsBackground = true,
                Name = "CommandUsageWriter",
                Priority = ThreadPriority.BelowNormal
            };
            thread.Start();
        }

        // Hot path: enqueues a struct — no I/O, no allocations beyond the queue node.
        public static void Record(string commandName, Result result, long durationMs)
        {
            _queue.Enqueue(new UsageRecord
            {
                CommandName = commandName,
                ExecutedAt = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss"),
                Result = result.ToString(),
                DurationMs = durationMs
            });
        }

        private static void BackgroundWriter()
        {
            InitializeDatabase();
            while (true)
            {
                Thread.Sleep(2000);
                if (!_queue.IsEmpty)
                    FlushQueue();
            }
        }

        private static void InitializeDatabase()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_dbPath));
#if REVIT2025 || REVIT2026
                using (var conn = new SqliteConnection($"Data Source={_dbPath}"))
#else
                using (var conn = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
#endif
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "PRAGMA journal_mode=WAL;";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"
                            CREATE TABLE IF NOT EXISTS CommandUsage (
                                Id        INTEGER PRIMARY KEY AUTOINCREMENT,
                                CommandName TEXT NOT NULL,
                                ExecutedAt  TEXT NOT NULL,
                                Result      TEXT NOT NULL,
                                DurationMs  INTEGER NOT NULL
                            );
                            CREATE INDEX IF NOT EXISTS idx_command ON CommandUsage(CommandName);";
                        cmd.ExecuteNonQuery();
                    }
                }
                _dbReady = true;
            }
            catch { }
        }

        private static void FlushQueue()
        {
            if (!_dbReady) return;
            try
            {
#if REVIT2025 || REVIT2026
                using (var conn = new SqliteConnection($"Data Source={_dbPath}"))
#else
                using (var conn = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
#endif
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText =
                                "INSERT INTO CommandUsage (CommandName, ExecutedAt, Result, DurationMs) " +
                                "VALUES (@n, @t, @r, @d)";
                            AddParameter(cmd, "@n", "");
                            AddParameter(cmd, "@t", "");
                            AddParameter(cmd, "@r", "");
                            AddParameter(cmd, "@d", 0L);

                            while (_queue.TryDequeue(out UsageRecord record))
                            {
                                cmd.Parameters["@n"].Value = record.CommandName;
                                cmd.Parameters["@t"].Value = record.ExecutedAt;
                                cmd.Parameters["@r"].Value = record.Result;
                                cmd.Parameters["@d"].Value = record.DurationMs;
                                cmd.ExecuteNonQuery();
                            }
                        }
                        tx.Commit();
                    }
                }
            }
            catch { }
        }

#if REVIT2025 || REVIT2026
        private static void AddParameter(SqliteCommand cmd, string name, object value)
#else
        private static void AddParameter(SQLiteCommand cmd, string name, object value)
#endif
        {
            var p = cmd.CreateParameter();
            p.ParameterName = name;
            p.Value = value;
            cmd.Parameters.Add(p);
        }
    }
}
