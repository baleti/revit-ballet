using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using OpenMcdf;

#if REVIT2025 || REVIT2026
using Microsoft.Data.Sqlite;
#else
using System.Data.SQLite;
#endif

namespace RevitBallet.Commands
{
    /// <summary>
    /// Helper class for Revit file operations - does NOT require RevitAPIUI
    /// Can be used from both Revit context and standalone network launcher
    /// </summary>
    public static class RevitFileHelper
    {
        /// <summary>
        /// Scans Documents folder for .rvt files and extracts metadata
        /// </summary>
        public static List<Dictionary<string, object>> GetRevitFilesFromDocuments()
        {
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var rvtFiles = new List<string>();

            // Log to a known location
            string logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "revit-ballet", "runtime", "RevitFileHelper_Debug.log"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(logPath));
            File.WriteAllText(logPath, $"=== RevitFileHelper Debug Log Started at {DateTime.Now} ===\n");

            // Recursively search for .rvt files, handling access denied errors
            GetRvtFilesRecursive(documentsPath, rvtFiles);

            var result = new List<Dictionary<string, object>>();

            foreach (var filePath in rvtFiles)
            {
                try
                {
                    var fileInfo = new FileInfo(filePath);
                    var basicInfo = ExtractBasicFileInfo(filePath, logPath);

                    var entry = new Dictionary<string, object>
                    {
                        ["File Name"] = Path.GetFileNameWithoutExtension(filePath),
                        ["Revit Version"] = basicInfo.RevitVersion ?? "Unknown",
                        ["Central Model"] = string.IsNullOrWhiteSpace(basicInfo.CentralModelPath)
                            ? "(Standalone file)"
                            : basicInfo.CentralModelPath,
                        ["Last Modified"] = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm"),
                        ["Last Opened"] = GetLastOpenedTime(filePath),
                        ["Path"] = filePath
                    };

                    result.Add(entry);
                }
                catch (Exception ex)
                {
                    // Skip files that can't be read
                    System.Diagnostics.Debug.WriteLine($"Failed to read {filePath}: {ex.Message}");
                }
            }

            // Sort by File Name ascending
            File.AppendAllText(logPath, $"\n=== Completed processing {result.Count} files ===\n");
            File.AppendAllText(logPath, $"Log file location: {logPath}\n");

            return result.OrderBy(e => e["File Name"]).ToList();
        }

        /// <summary>
        /// Recursively gets .rvt files, gracefully handling access denied errors
        /// </summary>
        private static void GetRvtFilesRecursive(string directory, List<string> rvtFiles)
        {
            try
            {
                // Get .rvt files in current directory (exclude backup files)
                var filesInDir = Directory.GetFiles(directory, "*.rvt")
                    .Where(f => !Path.GetFileName(f).EndsWith(".0001.rvt"))
                    .ToArray();

                rvtFiles.AddRange(filesInDir);

                // Recursively search subdirectories
                foreach (var subDir in Directory.GetDirectories(directory))
                {
                    GetRvtFilesRecursive(subDir, rvtFiles);
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Skip directories we don't have access to (e.g., My Music, system folders)
            }
            catch (Exception ex)
            {
                // Log other errors but continue
                System.Diagnostics.Debug.WriteLine($"Error accessing {directory}: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts BasicFileInfo from .rvt file using OLE Structured Storage
        /// </summary>
        private static BasicFileInfo ExtractBasicFileInfo(string filePath, string diagLog = null)
        {
            var info = new BasicFileInfo();

            try
            {
                if (diagLog != null)
                {
                    LogToDiag(diagLog, $"\n=== Processing file: {Path.GetFileName(filePath)} ===");
                }

                using (var cf = new CompoundFile(filePath))
                {
                    LogToDiag(diagLog, "Opened CompoundFile successfully");

                    // List all streams for debugging
                    try
                    {
                        cf.RootStorage.VisitEntries(item =>
                        {
                            LogToDiag(diagLog, $"Stream found: {item.Name} (IsStream: {item.IsStream}, IsStorage: {item.IsStorage})");
                        }, false);
                    }
                    catch (Exception ex)
                    {
                        LogToDiag(diagLog, $"Error listing streams: {ex.Message}");
                    }

                    // Read BasicFileInfo stream
                    if (cf.RootStorage.TryGetStream("BasicFileInfo", out CFStream stream))
                    {
                        LogToDiag(diagLog, "Found BasicFileInfo stream");

                        byte[] data = stream.GetData();
                        LogToDiag(diagLog, $"Read {data.Length} bytes from BasicFileInfo stream");

                        // Parse UTF-16LE encoded data
                        string content = Encoding.Unicode.GetString(data);
                        LogToDiag(diagLog, $"Decoded as UTF-16LE, total length: {content.Length} characters");

                        // Log first 1000 chars to see the structure
                        string debugContent = content.Length > 1000 ? content.Substring(0, 1000) : content;
                        LogToDiag(diagLog, $"First 1000 chars:\n{debugContent}");

                        // The BasicFileInfo is UTF-16LE which means null bytes appear between ASCII characters
                        // Strip ALL null characters first
                        string cleanedContent = content.Replace("\0", "");
                        LogToDiag(diagLog, $"After removing all nulls: {cleanedContent.Length} characters");

                        // The format is space-separated at the beginning:
                        // <special_chars> <username> <central_path> <version> <build> <local_path> ...

                        // Find first printable ASCII space or letter
                        int readableStart = 0;
                        for (int i = 0; i < Math.Min(20, cleanedContent.Length); i++)
                        {
                            if (cleanedContent[i] == ' ' || (cleanedContent[i] >= 'a' && cleanedContent[i] <= 'z'))
                            {
                                readableStart = i;
                                break;
                            }
                        }

                        // Take first 500 characters from readableStart for parsing
                        string readableContent = readableStart < cleanedContent.Length
                            ? cleanedContent.Substring(readableStart, Math.Min(500, cleanedContent.Length - readableStart))
                            : cleanedContent;

                        LogToDiag(diagLog, $"Readable content (starting at {readableStart}): '{readableContent.Substring(0, Math.Min(200, readableContent.Length))}'");

                        // Parse the space-separated format
                        // Pattern: <username> [optional chars] <central_unc_path> <year> <build_version> <local_path>
                        // Note: There may be a '[' or other character between username and UNC path
                        var spaceSepMatch = System.Text.RegularExpressions.Regex.Match(
                            readableContent,
                            @"(\\\\[\w\.\-]+\\[^\s]+\.rvt)\s+(\d{4})"
                        );

                        if (spaceSepMatch.Success)
                        {
                            // Extract central model path
                            string centralPath = spaceSepMatch.Groups[1].Value;
                            info.CentralModelPath = centralPath;
                            LogToDiag(diagLog, $"✓ Extracted Central Path from space-separated format: '{centralPath}'");

                            // Extract version
                            info.RevitVersion = spaceSepMatch.Groups[2].Value;
                            LogToDiag(diagLog, $"✓ Extracted Version from space-separated format: '{info.RevitVersion}'");
                        }
                        else
                        {
                            LogToDiag(diagLog, "✗ Failed to parse space-separated format, trying field-based format");

                            // Fall back to the old regex patterns for files that use different format
                            // cleanedContent already has nulls removed from line 160

                            // Extract Format (Revit version)
                            var formatMatch = System.Text.RegularExpressions.Regex.Match(cleanedContent, @"Format:\s*(\d{4})");
                            if (formatMatch.Success)
                            {
                                info.RevitVersion = formatMatch.Groups[1].Value;
                                LogToDiag(diagLog, $"✓ Found Format field: {info.RevitVersion}");
                            }
                            else
                            {
                                // Try finding just a 4-digit year in readable content
                                var yearMatch = System.Text.RegularExpressions.Regex.Match(readableContent, @"\b(20\d{2})\b");
                                if (yearMatch.Success)
                                {
                                    info.RevitVersion = yearMatch.Groups[1].Value;
                                    LogToDiag(diagLog, $"✓ Found year in readable content: {info.RevitVersion}");
                                }
                                else
                                {
                                    LogToDiag(diagLog, "✗ Could not find version");
                                }
                            }

                            // Extract Central Model Path from field format
                            var centralMatch = System.Text.RegularExpressions.Regex.Match(
                                cleanedContent,
                                @"Central Model Path:\s*(.+?)(?=\s*[A-Z][a-z]+\s*:|$)",
                                System.Text.RegularExpressions.RegexOptions.Singleline
                            );

                            if (centralMatch.Success)
                            {
                                string centralPath = centralMatch.Groups[1].Value;
                                centralPath = System.Text.RegularExpressions.Regex.Replace(centralPath, @"[\r\n\t]+", " ").Trim();

                                if (!string.IsNullOrWhiteSpace(centralPath))
                                {
                                    info.CentralModelPath = centralPath;
                                    LogToDiag(diagLog, $"✓ Found Central Model Path field: '{centralPath}'");
                                }
                            }
                        }
                    }
                    else
                    {
                        LogToDiag(diagLog, "✗ BasicFileInfo stream NOT found in file");
                    }
                }

                LogToDiag(diagLog, $"Final result - Version: {info.RevitVersion ?? "NULL"}, Central: {info.CentralModelPath ?? "NULL"}");
            }
            catch (Exception ex)
            {
                string errorMsg = $"ERROR reading BasicFileInfo from {filePath}: {ex.Message}\nStack: {ex.StackTrace}";
                LogToDiag(diagLog, errorMsg);
                System.Diagnostics.Debug.WriteLine(errorMsg);
            }

            return info;
        }

        /// <summary>
        /// Helper to log diagnostic information to a file
        /// </summary>
        private static void LogToDiag(string logPath, string message)
        {
            try
            {
                if (!string.IsNullOrEmpty(logPath))
                {
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] {message}\n");
                }
            }
            catch
            {
                // Ignore logging errors
            }
        }

        /// <summary>
        /// Gets last opened time from LogViewChanges database
        /// </summary>
        private static string GetLastOpenedTime(string filePath)
        {
            try
            {
                string dbPath = PathHelper.GetRuntimeFilePath("LogViewChanges.sqlite");
                if (!File.Exists(dbPath))
                    return "Never";

#if REVIT2025 || REVIT2026
                using (var conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={dbPath}"))
#else
                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath}"))
#endif
                {
                    conn.Open();

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
                            SELECT MAX(Timestamp)
                            FROM ViewHistory
                            WHERE DocumentPath = @path";

                        var param = cmd.CreateParameter();
                        param.ParameterName = "@path";
                        param.Value = filePath;
                        cmd.Parameters.Add(param);

                        var result = cmd.ExecuteScalar();

                        if (result != null && result != DBNull.Value)
                        {
                            // Timestamp is stored as Unix milliseconds (INTEGER)
                            if (long.TryParse(result.ToString(), out long unixMs))
                            {
                                DateTime lastOpened = DateTimeOffset.FromUnixTimeMilliseconds(unixMs).LocalDateTime;
                                return lastOpened.ToString("yyyy-MM-dd HH:mm");
                            }
                        }
                    }
                }
            }
            catch
            {
                // Database might not exist or be accessible
            }

            return "Never";
        }

        /// <summary>
        /// Opens a Revit file in the appropriate Revit version
        /// </summary>
        public static void OpenFileInRevit(string filePath, string revitVersion)
        {
            // Map version to Revit executable
            string revitExe = GetRevitExecutablePath(revitVersion);

            if (revitExe == null || !File.Exists(revitExe))
            {
                string expectedPath = revitExe ?? $@"C:\Program Files\Autodesk\Revit {revitVersion}\Revit.exe";
                throw new FileNotFoundException(
                    $"Revit {revitVersion} is not installed or could not be found.\n\n" +
                    $"Expected location: {expectedPath}\n\n" +
                    $"Please install Revit {revitVersion} or select files from a different version.");
            }

            // Launch Revit with the file
            var startInfo = new ProcessStartInfo
            {
                FileName = revitExe,
                Arguments = $"\"{filePath}\"",
                UseShellExecute = true
            };

            Process.Start(startInfo);
        }

        /// <summary>
        /// Gets the path to Revit.exe for a specific version
        /// </summary>
        private static string GetRevitExecutablePath(string version)
        {
            // Try to parse version (handle "2024", "Unknown", etc.)
            if (!int.TryParse(version, out int versionYear))
            {
                // Default to 2024 if version is unknown
                versionYear = 2024;
            }

            // Try registry first
            string registryPath = GetRevitPathFromRegistry(versionYear.ToString());
            if (registryPath != null)
            {
                string exePath = Path.Combine(registryPath, "Revit.exe");
                if (File.Exists(exePath))
                    return exePath;
            }

            // Fall back to default installation path
            string defaultPath = $@"C:\Program Files\Autodesk\Revit {versionYear}\Revit.exe";
            if (File.Exists(defaultPath))
                return defaultPath;

            return null;
        }

        /// <summary>
        /// Gets Revit installation path from registry
        /// </summary>
        private static string GetRevitPathFromRegistry(string version)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Autodesk\Revit\{version}"))
                {
                    if (key != null)
                    {
                        string installLocation = key.GetValue("InstallLocation") as string;
                        if (!string.IsNullOrEmpty(installLocation) && Directory.Exists(installLocation))
                        {
                            return installLocation;
                        }
                    }
                }
            }
            catch
            {
                // Registry access might fail
            }

            return null;
        }

        private class BasicFileInfo
        {
            public string RevitVersion { get; set; }
            public string CentralModelPath { get; set; }
        }
    }
}
