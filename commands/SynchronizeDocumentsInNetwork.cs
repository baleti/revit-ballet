using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using RevitBallet.Commands;
#if NET8_0_OR_GREATER
using System.Net.Http;
#endif

[Transaction(TransactionMode.Manual)]
public class SynchronizeDocumentsInNetwork : IExternalCommand
{
    /// <summary>
    /// Marks this command as usable outside Revit context via network.
    /// </summary>
    public static bool IsNetworkCommand => true;

    // Script to sync active document using PostCommand (works from ExternalEvent context)
    private const string SynchronizationScript = @"
// Post SynchronizeNow command for the active workshared document
var activeDoc = UIDoc?.Document;
if (activeDoc == null)
{
    Console.WriteLine(""No active document"");
    return;
}

if (!activeDoc.IsWorkshared)
{
    Console.WriteLine(""Active document is not workshared: "" + activeDoc.Title);
    return;
}

var syncCommandId = RevitCommandId.LookupPostableCommandId(PostableCommand.SynchronizeNow);
if (UIApp.CanPostCommand(syncCommandId))
{
    UIApp.PostCommand(syncCommandId);
    Console.WriteLine(""Sync command posted for: "" + activeDoc.Title);
}
else
{
    Console.WriteLine(""Cannot post sync command for: "" + activeDoc.Title);
}
";

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;

        using (var executionLog = CommandExecutionLogger.Start("SynchronizeDocumentsInNetwork", commandData))
        using (var diagnostics = CommandDiagnostics.StartCommand("SynchronizeDocumentsInNetwork", uiApp))
        {
            try
            {
                // Read documents file from runtime directory
                string documentsFilePath = Path.Combine(PathHelper.RuntimeDirectory, "documents");

                if (!File.Exists(documentsFilePath))
                {
                    TaskDialog.Show("Error", "No active documents found. Document registry file does not exist.");
                    diagnostics.LogError("Documents file not found");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }

                // Parse documents file
                var documents = ParseDocumentsFile(documentsFilePath);

                if (documents.Count == 0)
                {
                    TaskDialog.Show("Error", "No active documents found in registry.");
                    diagnostics.LogError("No documents in registry");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }

                // Get current session ID
                string currentSessionId = RevitBallet.RevitBallet.SessionId;

                // Read shared auth token
                string authToken = RevitBalletServer.GetSharedAuthToken();
                if (string.IsNullOrEmpty(authToken))
                {
                    TaskDialog.Show("Error", "Could not read authentication token from network registry.");
                    diagnostics.LogError("Auth token not found");
                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }

                // Prepare data for DataGrid - one row per document
                var gridData = new List<Dictionary<string, object>>();
                var columns = new List<string> { "Document", "Last Sync", "Session ID", "Hostname", "Port", "Last Heartbeat" };

                foreach (var docInfo in documents)
                {
                    bool isCurrent = docInfo.SessionId == currentSessionId;

                    var row = new Dictionary<string, object>
                    {
                        ["Document"] = string.IsNullOrWhiteSpace(docInfo.DocumentTitle) ? "(Home Page)" : docInfo.DocumentTitle,
                        ["Last Sync"] = FormatLastSync(docInfo.LastSync),
                        ["Session ID"] = docInfo.SessionId,
                        ["Port"] = docInfo.Port,
                        ["Hostname"] = docInfo.Hostname,
                        ["Last Heartbeat"] = FormatHeartbeat(docInfo.LastHeartbeat),
                        ["_Port"] = docInfo.Port,
                        ["_Hostname"] = docInfo.Hostname,
                        ["_SessionId"] = docInfo.SessionId,
                        ["_DocumentTitle"] = docInfo.DocumentTitle ?? "",
                        ["_IsCurrent"] = isCurrent
                    };
                    gridData.Add(row);
                }

                // Sort by Document column
                gridData = gridData.OrderBy(row => row["Document"].ToString()).ToList();

                // Find index of current session for initial selection
                int currentSessionIndex = gridData.FindIndex(row => Convert.ToBoolean(row["_IsCurrent"]));
                List<int> initialSelection = currentSessionIndex >= 0 ? new List<int> { currentSessionIndex } : null;

                // Show selection dialog with current session pre-selected
                var selectedRows = CustomGUIs.DataGrid(gridData, columns, false, initialSelection);

                if (selectedRows == null || selectedRows.Count == 0)
                {
                    diagnostics.Log("User cancelled selection");
                    executionLog.SetResult(Result.Cancelled);
                    return Result.Cancelled;
                }

                diagnostics.Log($"User selected {selectedRows.Count} document(s) for synchronization");

                // Group selected documents by session
                var documentsBySession = selectedRows
                    .Where(r => !string.IsNullOrEmpty(r["_DocumentTitle"]?.ToString())) // Skip Home Page entries
                    .GroupBy(r => r["_SessionId"].ToString())
                    .ToDictionary(
                        g => g.Key,
                        g => new SessionSyncInfo
                        {
                            SessionId = g.Key,
                            Port = g.First()["_Port"].ToString(),
                            Hostname = g.First()["_Hostname"].ToString(),
                            IsCurrent = Convert.ToBoolean(g.First()["_IsCurrent"]),
                            DocumentTitles = g.Select(r => r["_DocumentTitle"].ToString()).ToList()
                        });

                if (documentsBySession.Count == 0)
                {
                    TaskDialog.Show("Info", "No documents selected for synchronization (only Home Page entries selected).");
                    executionLog.SetResult(Result.Cancelled);
                    return Result.Cancelled;
                }

                // Synchronize documents grouped by session
                var localResults = new List<SyncResult>();
                int remoteRequestsSent = 0;

                foreach (var kvp in documentsBySession)
                {
                    var sessionInfo = kvp.Value;
                    diagnostics.Log($"Synchronizing {sessionInfo.DocumentTitles.Count} document(s) in session {sessionInfo.SessionId}");

                    if (sessionInfo.IsCurrent)
                    {
                        // For current session, execute synchronization directly for each document
                        foreach (var docTitle in sessionInfo.DocumentTitles)
                        {
                            var result = SynchronizeLocalDocument(uiApp, docTitle, diagnostics);
                            localResults.Add(new SyncResult
                            {
                                SessionId = sessionInfo.SessionId,
                                Document = docTitle,
                                Success = result.Success,
                                Message = result.Message
                            });
                        }
                    }
                    else
                    {
                        // For remote sessions, send code via Roslyn endpoint - fire and forget
                        // Note: This syncs the ACTIVE document in the remote session using PostCommand
                        SendSynchronizationRequestAsync(
                            sessionInfo.Hostname,
                            sessionInfo.Port,
                            authToken,
                            diagnostics);
                        remoteRequestsSent++;
                    }
                }

                // Report local results only
                int localSuccessCount = localResults.Count(r => r.Success);
                int localFailCount = localResults.Count(r => !r.Success);

                diagnostics.Log($"Local sync complete: {localSuccessCount} succeeded, {localFailCount} failed. Remote requests sent: {remoteRequestsSent}");

                // Only show dialog if there were local errors
                if (localFailCount > 0)
                {
                    var errorMessages = localResults.Where(r => !r.Success)
                        .Select(r => $"{r.Document}: {r.Message}")
                        .Distinct()
                        .ToList();

                    TaskDialog.Show("Synchronization Errors",
                        $"Local synchronization completed with {localFailCount} error(s):\n\n" +
                        string.Join("\n", errorMessages) +
                        (remoteRequestsSent > 0 ? $"\n\n({remoteRequestsSent} remote sync request(s) sent in background)" : ""));

                    executionLog.SetResult(Result.Failed);
                    return Result.Failed;
                }

                executionLog.SetResult(Result.Succeeded);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to synchronize network documents: {ex.Message}");
                diagnostics.LogError($"Exception: {ex}");
                executionLog.SetResult(Result.Failed);
                return Result.Failed;
            }
        }
    }

    private SyncResult SynchronizeLocalDocument(UIApplication uiApp, string documentTitle, CommandDiagnostics.DiagnosticSession diagnostics)
    {
        // Find the document by title
        Document targetDoc = null;
        foreach (Document doc in uiApp.Application.Documents)
        {
            if (!doc.IsLinked && doc.Title.Equals(documentTitle, StringComparison.OrdinalIgnoreCase))
            {
                targetDoc = doc;
                break;
            }
        }

        if (targetDoc == null)
        {
            return new SyncResult
            {
                Success = false,
                Message = "Document not found"
            };
        }

        if (targetDoc.IsFamilyDocument)
        {
            return new SyncResult
            {
                Success = false,
                Message = "Family documents cannot be synchronized"
            };
        }

        if (!targetDoc.IsWorkshared)
        {
            return new SyncResult
            {
                Success = false,
                Message = "Document is not workshared"
            };
        }

        try
        {
            var transactOptions = new TransactWithCentralOptions();
            var syncOptions = new SynchronizeWithCentralOptions();
            syncOptions.Comment = "Synchronized via Revit Ballet Network";
            syncOptions.Compact = false;
            syncOptions.SaveLocalBefore = false;
            syncOptions.SaveLocalAfter = true;

            targetDoc.SynchronizeWithCentral(transactOptions, syncOptions);
            diagnostics.Log($"Synchronized local document: {documentTitle}");

            return new SyncResult
            {
                Success = true,
                Message = "Synchronized successfully"
            };
        }
        catch (Exception ex)
        {
            diagnostics.LogError($"Failed to sync local document {documentTitle}: {ex.Message}");
            return new SyncResult
            {
                Success = false,
                Message = ex.Message
            };
        }
    }

    private void SendSynchronizationRequestAsync(string hostname, string port, string authToken, CommandDiagnostics.DiagnosticSession diagnostics)
    {
        // Fire and forget - send request without waiting for response
        // Uses PostCommand(SynchronizeNow) to sync the active document in the remote session
        string url = $"https://127.0.0.1:{port}/roslyn";

        diagnostics.Log($"Sending async sync request to 127.0.0.1:{port} ({hostname})");

#if NET8_0_OR_GREATER
        // Use Task.Run with HttpClient for .NET 8+
        System.Threading.Tasks.Task.Run(async () =>
        {
            try
            {
                using (var handler = new HttpClientHandler())
                {
                    handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;

                    using (var client = new HttpClient(handler))
                    {
                        client.Timeout = TimeSpan.FromMinutes(2);
                        client.DefaultRequestHeaders.Add("X-Auth-Token", authToken);

                        var content = new StringContent(SynchronizationScript, Encoding.UTF8, "text/plain");
                        await client.PostAsync(url, content);
                    }
                }
            }
            catch
            {
                // Fire and forget - ignore errors
            }
        });
#else
        // Use ThreadPool for .NET Framework
        System.Threading.ThreadPool.QueueUserWorkItem(_ =>
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "text/plain";
                request.Headers.Add("X-Auth-Token", authToken);
                request.Timeout = 120000;
                request.ServerCertificateValidationCallback = (sender, cert, chain, errors) => true;

                byte[] bodyBytes = Encoding.UTF8.GetBytes(SynchronizationScript);
                request.ContentLength = bodyBytes.Length;

                using (var requestStream = request.GetRequestStream())
                {
                    requestStream.Write(bodyBytes, 0, bodyBytes.Length);
                }

                // Get response but don't process it
                using (var response = request.GetResponse())
                {
                    // Just read to complete the request
                }
            }
            catch
            {
                // Fire and forget - ignore errors
            }
        });
#endif
    }


    private List<DocumentInfo> ParseDocumentsFile(string filePath)
    {
        var documents = new List<DocumentInfo>();

        try
        {
            var lines = File.ReadAllLines(filePath);

            foreach (var line in lines)
            {
                // Skip comments and empty lines
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                    continue;

                var parts = line.Split(',');
                // Format: DocumentTitle,DocumentPath,SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,LastSync
                if (parts.Length < 8)
                    continue;

                var doc = new DocumentInfo
                {
                    DocumentTitle = parts[0].Trim(),
                    DocumentPath = parts[1].Trim(),
                    SessionId = parts[2].Trim(),
                    Port = parts[3].Trim(),
                    Hostname = parts[4].Trim(),
                    ProcessId = int.Parse(parts[5].Trim()),
                    RegisteredAt = DateTime.Parse(parts[6].Trim()),
                    LastHeartbeat = DateTime.Parse(parts[7].Trim())
                };

                // LastSync is optional (index 8)
                if (parts.Length > 8 && !string.IsNullOrWhiteSpace(parts[8]))
                {
                    DateTime lastSyncParsed;
                    if (DateTime.TryParse(parts[8].Trim(), out lastSyncParsed))
                    {
                        doc.LastSync = lastSyncParsed;
                    }
                }

                // Filter out stale documents (no heartbeat for > 2 minutes)
                if ((DateTime.Now - doc.LastHeartbeat).TotalSeconds <= 120)
                {
                    documents.Add(doc);
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to parse documents file: {ex.Message}", ex);
        }

        return documents;
    }

    private string FormatHeartbeat(DateTime heartbeat)
    {
        var timeAgo = DateTime.Now - heartbeat;

        if (timeAgo.TotalSeconds < 60)
            return "Just now";
        else if (timeAgo.TotalMinutes < 60)
            return $"{(int)timeAgo.TotalMinutes}m ago";
        else if (timeAgo.TotalHours < 24)
            return $"{(int)timeAgo.TotalHours}h ago";
        else
            return $"{(int)timeAgo.TotalDays}d ago";
    }

    private string FormatLastSync(DateTime? lastSync)
    {
        if (!lastSync.HasValue)
            return "-";

        var timeAgo = DateTime.Now - lastSync.Value;

        if (timeAgo.TotalSeconds < 60)
            return "Just now";
        else if (timeAgo.TotalMinutes < 60)
            return $"{(int)timeAgo.TotalMinutes}m ago";
        else if (timeAgo.TotalHours < 24)
            return $"{(int)timeAgo.TotalHours}h ago";
        else
            return $"{(int)timeAgo.TotalDays}d ago";
    }

    private class DocumentInfo
    {
        public string DocumentTitle { get; set; }
        public string DocumentPath { get; set; }
        public string SessionId { get; set; }
        public string Port { get; set; }
        public string Hostname { get; set; }
        public int ProcessId { get; set; }
        public DateTime RegisteredAt { get; set; }
        public DateTime LastHeartbeat { get; set; }
        public DateTime? LastSync { get; set; }
    }

    private class SessionSyncInfo
    {
        public string SessionId { get; set; }
        public string Port { get; set; }
        public string Hostname { get; set; }
        public bool IsCurrent { get; set; }
        public List<string> DocumentTitles { get; set; }
    }

    private class SyncResult
    {
        public string SessionId { get; set; }
        public string Document { get; set; }
        public bool Success { get; set; }
        public string Message { get; set; }
    }

}
