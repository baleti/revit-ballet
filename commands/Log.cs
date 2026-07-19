using System;
using System.IO;

namespace RevitBallet.Commands
{
    /// <summary>
    /// Minimal diagnostic logger for failures that must not interrupt the user.
    /// Appends to %APPDATA%\revit-ballet\runtime\addin.log. Never throws.
    /// </summary>
    public static class Log
    {
        private static readonly object _lock = new object();

        public static string LogFilePath => PathHelper.GetRuntimeFilePath("addin.log");

        public static void Warn(string context, Exception ex) => Write("WARN", context, ex?.ToString());

        public static void Warn(string context, string message) => Write("WARN", context, message);

        public static void Info(string context, string message) => Write("INFO", context, message);

        private static void Write(string level, string context, string detail)
        {
            try
            {
                lock (_lock)
                {
                    File.AppendAllText(LogFilePath,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " [" + level + "] [" + context + "] " + detail + Environment.NewLine);
                }
            }
            catch
            {
                // Logging must never take the addin down.
            }
        }
    }
}
