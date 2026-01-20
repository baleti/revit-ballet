#if REVIT2011 || REVIT2012 || REVIT2013 || REVIT2014 || REVIT2015 || REVIT2016 || REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Reflection;
using System.Security.Authentication;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using RevitBallet.Commands;

namespace RevitBallet.Commands
{
    public class RevitBalletServer
    {
        private static RevitBalletServer serverInstance = null;
        private static readonly object lockObject = new object();

        /// <summary>
        /// Initialize and start the Revit Ballet server.
        /// Called by Startup.RunStartupTasks()
        /// </summary>
        public static void InitializeServer()
        {
            // Auto-start server on Revit startup
            lock (lockObject)
            {
                if (serverInstance == null)
                {
                    try
                    {
                        serverInstance = new RevitBalletServer();
                        serverInstance.Start();
                    }
                    catch
                    {
                        // Silently fail - server may already be running from previous load
                    }
                }
            }
        }

        /// <summary>
        /// Terminate the Revit Ballet server.
        /// </summary>
        public static void TerminateServer()
        {
            lock (lockObject)
            {
                if (serverInstance != null)
                {
                    try
                    {
                        serverInstance.Stop();
                    }
                    catch { }
                    serverInstance = null;
                }
            }
        }

        // Instance fields and methods below
        private const int PORT_START = 23717;
        private const int PORT_END = 23817; // Allow 100 concurrent Revit sessions
        private const string NETWORK_FOLDER = "network";
        private const string DOCUMENTS_FILE = "documents";
        private const string TOKEN_FILE = "token";

        private TcpListener listener;
        private CancellationTokenSource cancellationTokenSource;
        private bool isRunning = false;
        private int serverPort;
        private X509Certificate2 serverCertificate;
        private string sessionId;
        private static UIApplication uiApp;
        private ExternalEvent scriptExecutionEvent;
        private ExternalEvent screenshotEvent;
        private ExternalEvent queryEvent;
        private ScriptExecutionHandler scriptHandler;
        private ScreenshotHandler screenshotHandler;
        private QueryHandler queryHandler;
        private int activeConnections = 0;
        private int totalConnectionsAccepted = 0;
        private int totalConnectionsSuccessful = 0;
        private int totalConnectionsFailed = 0;

        public bool IsRunning => isRunning;
        public int Port => serverPort;
        public string SessionId => sessionId;

        public static void SetUIApplication(UIApplication application)
        {
            uiApp = application;
        }

        internal static void LogToRevit(string message)
        {
            try
            {
                // For Revit, we'll write to a log file since we can't easily write to the command line
                // This is different from AutoCAD which has an Editor.WriteMessage
                string logPath = PathHelper.GetRuntimeFilePath("server.log");
                File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n");
            }
            catch
            {
                // Ignore logging errors completely
            }
        }

        public void Start()
        {
            if (isRunning)
                return;

            sessionId = System.Diagnostics.Process.GetCurrentProcess().Id.ToString();

            // Load or generate self-signed certificate for HTTPS
            try
            {
                serverCertificate = SslCertificateHelper.GetOrCreateCertificate();

                // Log certificate diagnostics
                LogToRevit($"[CERT] Thumbprint: {serverCertificate.Thumbprint}");
                LogToRevit($"[CERT] Subject: {serverCertificate.Subject}");
                LogToRevit($"[CERT] Valid: {serverCertificate.NotBefore:yyyy-MM-dd} to {serverCertificate.NotAfter:yyyy-MM-dd}");
                LogToRevit($"[CERT] Has Private Key: {serverCertificate.HasPrivateKey}");
                LogToRevit($"[CERT] Key Algorithm: {serverCertificate.GetKeyAlgorithm()}");
            }
            catch (NotSupportedException ex)
            {
                throw new InvalidOperationException("SSL/TLS certificate generation failed.\n" + ex.Message);
            }

            // Read or generate shared auth token (stored in network/token file)
            GenerateOrReadSharedToken();

            // Find available port
            serverPort = FindAvailablePort(PORT_START, PORT_END);
            if (serverPort == -1)
            {
                throw new InvalidOperationException($"No available ports found between {PORT_START} and {PORT_END}");
            }

            cancellationTokenSource = new CancellationTokenSource();
            listener = new TcpListener(IPAddress.Loopback, serverPort);
            listener.Start();

            // Create ExternalEvent handlers (must be done during API execution)
            scriptHandler = new ScriptExecutionHandler();
            scriptExecutionEvent = ExternalEvent.Create(scriptHandler);

            screenshotHandler = new ScreenshotHandler(sessionId);
            screenshotEvent = ExternalEvent.Create(screenshotHandler);

            queryHandler = new QueryHandler();
            queryEvent = ExternalEvent.Create(queryHandler);

            isRunning = true;

            // Register in network registry
            RegisterInNetwork();

            // Start background listener thread
            Task.Run(() => AcceptRequestsLoop(cancellationTokenSource.Token));

            // Start heartbeat to keep network registry alive
            Task.Run(() => UpdateNetworkHeartbeat(cancellationTokenSource.Token));

#if NET8_0_OR_GREATER
            var tlsVersions = "TLS 1.2, 1.3";
#else
            var tlsVersions = "TLS 1.2";
#endif
            LogToRevit($"[SERVER] Started on port {serverPort} (Session: {sessionId.Substring(0, 8)})");
            LogToRevit($"[SERVER] TLS Protocols: {tlsVersions}");
            LogToRevit($"[SERVER] .NET Version: {Environment.Version}");
        }

        public void Stop()
        {
            if (!isRunning)
                return;

            isRunning = false;

            // Unregister from network
            UnregisterFromNetwork();

            try
            {
                cancellationTokenSource?.Cancel();
            }
            catch { }

            try
            {
                listener?.Stop();
            }
            catch { }

            try
            {
                serverCertificate?.Dispose();
                serverCertificate = null;
            }
            catch { }

            try
            {
                scriptExecutionEvent?.Dispose();
                scriptExecutionEvent = null;
            }
            catch { }

            try
            {
                screenshotEvent?.Dispose();
                screenshotEvent = null;
            }
            catch { }

            try
            {
                queryEvent?.Dispose();
                queryEvent = null;
            }
            catch { }

            try
            {
                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
            }
            catch { }
        }

        private int FindAvailablePort(int startPort, int endPort)
        {
            for (int port = startPort; port <= endPort; port++)
            {
                try
                {
                    var testListener = new TcpListener(IPAddress.Loopback, port);
                    testListener.Start();
                    testListener.Stop();
                    return port;
                }
                catch (SocketException)
                {
                    continue;
                }
            }
            return -1;
        }

        private void RegisterInNetwork()
        {
            var newEntries = CreateDocumentEntries();

            lock (typeof(RevitBalletServer))
            {
                var entries = ReadDocumentRegistry();
                // Remove any existing entries for this session before adding new ones
                entries.RemoveAll(e => e.SessionId == sessionId);
                entries.AddRange(newEntries);
                WriteDocumentRegistry(entries);
            }
        }

        private void UnregisterFromNetwork()
        {
            lock (typeof(RevitBalletServer))
            {
                var entries = ReadDocumentRegistry();
                entries.RemoveAll(e => e.SessionId == sessionId);
                WriteDocumentRegistry(entries);
            }
        }

        /// <summary>
        /// Updates the LastSync timestamp for a specific document in the registry.
        /// Called when a document is synchronized with central.
        /// </summary>
        /// <param name="documentPath">Path of the document that was synchronized</param>
        public static void UpdateLastSyncTime(string documentPath)
        {
            if (serverInstance == null) return;

            try
            {
                lock (typeof(RevitBalletServer))
                {
                    var entries = serverInstance.ReadDocumentRegistry();
                    // Find the document by path (or title if path is empty) in this session
                    var docEntry = entries.FirstOrDefault(e =>
                        e.SessionId == serverInstance.sessionId &&
                        (e.DocumentPath == documentPath || (string.IsNullOrEmpty(e.DocumentPath) && e.DocumentTitle == documentPath)));
                    if (docEntry != null)
                    {
                        docEntry.LastSync = DateTime.Now;
                        serverInstance.WriteDocumentRegistry(entries);
                    }
                }
            }
            catch
            {
                // Silently fail - don't interrupt sync operation
            }
        }

        private async Task UpdateNetworkHeartbeat(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested && isRunning)
            {
                try
                {
                    await Task.Delay(30000, cancellationToken); // Every 30 seconds

                    lock (typeof(RevitBalletServer))
                    {
                        var entries = ReadDocumentRegistry();
                        var now = DateTime.Now;

                        // Get current open documents
                        var openDocs = GetOpenDocuments();

                        // Get existing entries for this session (to preserve LastSync)
                        var myEntries = entries.Where(e => e.SessionId == sessionId).ToList();

                        // Remove all entries for this session
                        entries.RemoveAll(e => e.SessionId == sessionId);

                        // Add entries for currently open documents, preserving LastSync where applicable
                        foreach (var doc in openDocs)
                        {
                            // Find existing entry to preserve LastSync
                            var existingEntry = myEntries.FirstOrDefault(e =>
                                e.DocumentPath == doc.Path || e.DocumentTitle == doc.Title);

                            entries.Add(new DocumentEntry
                            {
                                DocumentTitle = doc.Title,
                                DocumentPath = doc.Path,
                                SessionId = sessionId,
                                Port = serverPort,
                                Hostname = Environment.MachineName,
                                ProcessId = System.Diagnostics.Process.GetCurrentProcess().Id,
                                RegisteredAt = existingEntry?.RegisteredAt ?? now,
                                LastHeartbeat = now,
                                LastSync = existingEntry?.LastSync
                            });
                        }

                        // Clean up dead documents (from sessions with no heartbeat for 2 minutes)
                        entries.RemoveAll(e => (now - e.LastHeartbeat).TotalSeconds > 120);

                        WriteDocumentRegistry(entries);
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch { }
            }
        }

        private List<DocumentEntry> CreateDocumentEntries()
        {
            var entries = new List<DocumentEntry>();
            var openDocs = GetOpenDocuments();
            var now = DateTime.Now;

            foreach (var doc in openDocs)
            {
                entries.Add(new DocumentEntry
                {
                    DocumentTitle = doc.Title,
                    DocumentPath = doc.Path,
                    SessionId = sessionId,
                    Port = serverPort,
                    Hostname = Environment.MachineName,
                    ProcessId = System.Diagnostics.Process.GetCurrentProcess().Id,
                    RegisteredAt = now,
                    LastHeartbeat = now
                });
            }

            return entries;
        }

        private class OpenDocumentInfo
        {
            public string Title { get; set; }
            public string Path { get; set; }
        }

        private List<OpenDocumentInfo> GetOpenDocuments()
        {
            var docs = new List<OpenDocumentInfo>();
            try
            {
                if (uiApp != null)
                {
                    foreach (Document doc in uiApp.Application.Documents)
                    {
                        // Skip linked documents - only include actively opened documents
                        if (!doc.IsLinked)
                        {
                            docs.Add(new OpenDocumentInfo
                            {
                                Title = doc.Title,
                                Path = doc.PathName ?? ""
                            });
                        }
                    }
                }
            }
            catch { }
            return docs;
        }

        private string GetNetworkFolderPath()
        {
            var networkDir = PathHelper.EnsureRuntimeSubdirectoryExists(NETWORK_FOLDER);
            return networkDir;
        }

        private string GetDocumentsFilePath()
        {
            // Documents file is in runtime root, not in network subfolder
            return Path.Combine(PathHelper.RuntimeDirectory, DOCUMENTS_FILE);
        }

        private string GetTokenFilePath()
        {
            return Path.Combine(GetNetworkFolderPath(), TOKEN_FILE);
        }

        private List<DocumentEntry> ReadDocumentRegistry()
        {
            var path = GetDocumentsFilePath();
            var entries = new List<DocumentEntry>();

            if (!File.Exists(path))
                return entries;

            try
            {
                var lines = File.ReadAllLines(path);
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                        continue;

                    var parts = line.Split(',');
                    // Format: DocumentTitle,DocumentPath,SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,LastSync
                    if (parts.Length >= 8)
                    {
                        var entry = new DocumentEntry
                        {
                            DocumentTitle = parts[0],
                            DocumentPath = parts[1],
                            SessionId = parts[2],
                            Port = int.Parse(parts[3]),
                            Hostname = parts[4],
                            ProcessId = int.Parse(parts[5]),
                            RegisteredAt = DateTime.Parse(parts[6]),
                            LastHeartbeat = DateTime.Parse(parts[7])
                        };

                        // LastSync is optional (index 8)
                        if (parts.Length > 8 && !string.IsNullOrWhiteSpace(parts[8]))
                        {
                            DateTime lastSyncParsed;
                            if (DateTime.TryParse(parts[8], out lastSyncParsed))
                            {
                                entry.LastSync = lastSyncParsed;
                            }
                        }

                        entries.Add(entry);
                    }
                }
            }
            catch { }

            return entries;
        }

        private void WriteDocumentRegistry(List<DocumentEntry> entries)
        {
            var path = GetDocumentsFilePath();
            var lines = new List<string>();

            lines.Add("# Revit Ballet Document Registry");
            lines.Add("# DocumentTitle,DocumentPath,SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,LastSync");

            foreach (var entry in entries)
            {
                var lastSyncStr = entry.LastSync.HasValue ? entry.LastSync.Value.ToString("O") : "";
                var line = $"{entry.DocumentTitle},{entry.DocumentPath},{entry.SessionId},{entry.Port},{entry.Hostname},{entry.ProcessId},{entry.RegisteredAt:O},{entry.LastHeartbeat:O},{lastSyncStr}";
                lines.Add(line);
            }

            File.WriteAllLines(path, lines);
        }

        public static List<DocumentEntry> GetActiveDocuments()
        {
            var path = Path.Combine(PathHelper.RuntimeDirectory, DOCUMENTS_FILE);

            var entries = new List<DocumentEntry>();

            if (!File.Exists(path))
                return entries;

            lock (typeof(RevitBalletServer))
            {
                try
                {
                    var lines = File.ReadAllLines(path);
                    foreach (var line in lines)
                    {
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                            continue;

                        var parts = line.Split(',');
                        // Format: DocumentTitle,DocumentPath,SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,LastSync
                        if (parts.Length >= 8)
                        {
                            var entry = new DocumentEntry
                            {
                                DocumentTitle = parts[0],
                                DocumentPath = parts[1],
                                SessionId = parts[2],
                                Port = int.Parse(parts[3]),
                                Hostname = parts[4],
                                ProcessId = int.Parse(parts[5]),
                                RegisteredAt = DateTime.Parse(parts[6]),
                                LastHeartbeat = DateTime.Parse(parts[7])
                            };

                            // LastSync is optional (index 8)
                            if (parts.Length > 8 && !string.IsNullOrWhiteSpace(parts[8]))
                            {
                                DateTime lastSyncParsed;
                                if (DateTime.TryParse(parts[8], out lastSyncParsed))
                                {
                                    entry.LastSync = lastSyncParsed;
                                }
                            }

                            // Only return entries with recent heartbeat
                            if ((DateTime.Now - entry.LastHeartbeat).TotalSeconds < 120)
                            {
                                entries.Add(entry);
                            }
                        }
                    }
                }
                catch { }
            }

            return entries;
        }

        public static string GetSharedAuthToken()
        {
            var networkDir = PathHelper.EnsureRuntimeSubdirectoryExists(NETWORK_FOLDER);
            var tokenPath = Path.Combine(networkDir, TOKEN_FILE);

            if (!File.Exists(tokenPath))
                return null;

            try
            {
                return File.ReadAllText(tokenPath).Trim();
            }
            catch
            {
                return null;
            }
        }

        private string GenerateOrReadSharedToken()
        {
            var tokenPath = GetTokenFilePath();

            // Try to read existing token
            if (File.Exists(tokenPath))
            {
                try
                {
                    var existingToken = File.ReadAllText(tokenPath).Trim();
                    if (!string.IsNullOrWhiteSpace(existingToken))
                    {
                        return existingToken;
                    }
                }
                catch { }
            }

            // Generate new token if none exists or reading failed
            var tokenBytes = new byte[32];
#if NET8_0_OR_GREATER
            RandomNumberGenerator.Fill(tokenBytes);
#else
            using (var rng = System.Security.Cryptography.RandomNumberGenerator.Create())
            {
                rng.GetBytes(tokenBytes);
            }
#endif
            var newToken = Convert.ToBase64String(tokenBytes);

            // Store the token for other sessions to use
            try
            {
                File.WriteAllText(tokenPath, newToken);
            }
            catch { }

            return newToken;
        }

        private async Task AcceptRequestsLoop(CancellationToken cancellationToken)
        {
            LogToRevit($"[SERVER] Accept loop started, listening on port {serverPort}");

            while (!cancellationToken.IsCancellationRequested && isRunning)
            {
                try
                {
                    // CRITICAL FIX: Don't use Pending() - it can miss connections
                    // Use AcceptTcpClientAsync() directly
#if NETFRAMEWORK
                    // .NET Framework doesn't have AcceptTcpClientAsync(CancellationToken) overload
                    // Use Task.WhenAny pattern for cancellation
                    var acceptTask = listener.AcceptTcpClientAsync();
                    var completedTask = await Task.WhenAny(acceptTask, Task.Delay(-1, cancellationToken));

                    if (completedTask != acceptTask)
                    {
                        // Cancellation was requested
                        break;
                    }

                    var client = await acceptTask;
#else
                    // .NET 8.0+ has the CancellationToken overload
                    var client = await listener.AcceptTcpClientAsync(cancellationToken);
#endif
                    var clientEndpoint = client.Client.RemoteEndPoint.ToString();

                    LogToRevit($"[SERVER] Accepted TCP connection from {clientEndpoint}");

                    // Handle request in background thread
                    _ = Task.Run(async () => await HandleHttpsRequest(client, clientEndpoint, cancellationToken));
                }
                catch (OperationCanceledException)
                {
                    LogToRevit($"[SERVER] Accept loop cancelled");
                    break;
                }
                catch (System.Exception ex)
                {
                    if (!cancellationToken.IsCancellationRequested && isRunning)
                    {
                        LogToRevit($"[SERVER] Error in accept loop: {ex.Message}");
                        LogToRevit($"[SERVER] Exception type: {ex.GetType().FullName}");

                        // Don't log stack trace for common errors
                        if (!(ex is ObjectDisposedException))
                        {
                            LogToRevit($"[SERVER] Stack trace: {ex.StackTrace}");
                        }

                        // Small delay before retrying to prevent tight error loops
                        await Task.Delay(1000, cancellationToken);
                    }
                }
            }

            LogToRevit($"[SERVER] Accept loop exited");
        }

        private async Task HandleHttpsRequest(TcpClient client, string clientEndpoint, CancellationToken cancellationToken)
        {
            var connId = Guid.NewGuid().ToString("N").Substring(0, 8);
            Interlocked.Increment(ref activeConnections);
            Interlocked.Increment(ref totalConnectionsAccepted);

            LogToRevit($"[CONN-{connId}] TCP connection from {clientEndpoint} (Active: {activeConnections}, Total Accepted: {totalConnectionsAccepted})");

            var startTime = DateTime.Now;
            bool handshakeSuccess = false;

            try
            {
                using (client)
                {
                    var networkStream = client.GetStream();
                    LogToRevit($"[CONN-{connId}] NetworkStream acquired (CanRead: {networkStream.CanRead}, CanWrite: {networkStream.CanWrite})");

                    // Check if network stream is still connected
                    if (!client.Connected)
                    {
                        LogToRevit($"[CONN-{connId}] ERROR: Client disconnected before SSL handshake");
                        Interlocked.Increment(ref totalConnectionsFailed);
                        return;
                    }

                    var sslStream = new SslStream(networkStream, false);

                    try
                    {
                        if (!serverCertificate.HasPrivateKey)
                        {
                            LogToRevit($"[CONN-{connId}] ERROR: Certificate has no private key!");
                            Interlocked.Increment(ref totalConnectionsFailed);
                            return;
                        }

                        LogToRevit($"[CONN-{connId}] Starting SSL handshake...");
                        var handshakeStart = DateTime.Now;

#if NET8_0_OR_GREATER
                        var protocols = SslProtocols.Tls12 | SslProtocols.Tls13;
#else
                        var protocols = SslProtocols.Tls12;
#endif

                        // Add timeout to handshake to detect hung connections
                        var handshakeTask = sslStream.AuthenticateAsServerAsync(
                            serverCertificate,
                            clientCertificateRequired: false,
                            enabledSslProtocols: protocols,
                            checkCertificateRevocation: false
                        );

                        var timeoutTask = Task.Delay(10000); // 10 second timeout
                        var completedTask = await Task.WhenAny(handshakeTask, timeoutTask);

                        if (completedTask == timeoutTask)
                        {
                            LogToRevit($"[CONN-{connId}] ERROR: SSL handshake timeout (>10s)");
                            LogToRevit($"[CONN-{connId}] Client may have closed connection or sent invalid data");
                            Interlocked.Increment(ref totalConnectionsFailed);
                            return;
                        }

                        await handshakeTask; // Re-await to get any exceptions

                        var handshakeDuration = (DateTime.Now - handshakeStart).TotalMilliseconds;
                        handshakeSuccess = true;
                        Interlocked.Increment(ref totalConnectionsSuccessful);

                        LogToRevit($"[CONN-{connId}] SSL handshake successful ({handshakeDuration:F0}ms)");
                        LogToRevit($"[CONN-{connId}] Protocol: {sslStream.SslProtocol}, Cipher: {sslStream.CipherAlgorithm}, Hash: {sslStream.HashAlgorithm}");

                        await HandleHttpRequest(sslStream, connId, cancellationToken);
                    }
                    catch (AuthenticationException ex)
                    {
                        var handshakeDuration = (DateTime.Now - startTime).TotalMilliseconds;
                        LogToRevit($"[CONN-{connId}] SSL handshake failed after {handshakeDuration:F0}ms: {ex.Message}");
                        LogToRevit($"[CONN-{connId}] Exception type: {ex.GetType().FullName}");
                        LogToRevit($"[CONN-{connId}] Stack trace:\n{ex.StackTrace}");
                        LogToRevit($"[CONN-{connId}] Hint: Client may be using HTTP instead of HTTPS, or TLS version mismatch");
                        Interlocked.Increment(ref totalConnectionsFailed);
                    }
                    catch (System.IO.IOException ex)
                    {
                        var handshakeDuration = (DateTime.Now - startTime).TotalMilliseconds;
                        LogToRevit($"[CONN-{connId}] I/O error after {handshakeDuration:F0}ms: {ex.Message}");
                        LogToRevit($"[CONN-{connId}] Exception type: {ex.GetType().FullName}");

                        if (ex.InnerException != null)
                        {
                            LogToRevit($"[CONN-{connId}] Inner exception: {ex.InnerException.GetType().FullName} - {ex.InnerException.Message}");

                            // Common Windows socket errors
                            var socketEx = ex.InnerException as SocketException;
                            if (socketEx != null)
                            {
                                LogToRevit($"[CONN-{connId}] Socket error code: {socketEx.ErrorCode} (0x{socketEx.ErrorCode:X})");
                                LogToRevit($"[CONN-{connId}] Socket error: {socketEx.SocketErrorCode}");
                            }
                        }

                        LogToRevit($"[CONN-{connId}] Hint: Connection may have been reset by client or intermediary (named pipe relay?)");
                        Interlocked.Increment(ref totalConnectionsFailed);
                    }
                    catch (System.Exception ex)
                    {
                        var handshakeDuration = (DateTime.Now - startTime).TotalMilliseconds;
                        LogToRevit($"[CONN-{connId}] Error in SSL stream after {handshakeDuration:F0}ms: {ex.Message}");
                        LogToRevit($"[CONN-{connId}] Exception type: {ex.GetType().FullName}");
                        LogToRevit($"[CONN-{connId}] Stack trace:\n{ex.StackTrace}");
                        if (ex.InnerException != null)
                        {
                            LogToRevit($"[CONN-{connId}] Inner exception: {ex.InnerException.Message}");
                            LogToRevit($"[CONN-{connId}] Inner stack trace:\n{ex.InnerException.StackTrace}");
                        }
                        Interlocked.Increment(ref totalConnectionsFailed);
                    }
                    finally
                    {
                        try
                        {
                            sslStream?.Dispose();
                        }
                        catch { }

                        LogToRevit($"[CONN-{connId}] SSL stream disposed");
                    }
                }

                var totalDuration = (DateTime.Now - startTime).TotalMilliseconds;
                LogToRevit($"[CONN-{connId}] Connection closed (Duration: {totalDuration:F0}ms, Handshake: {(handshakeSuccess ? "SUCCESS" : "FAILED")})");
            }
            catch (System.Exception ex)
            {
                var totalDuration = (DateTime.Now - startTime).TotalMilliseconds;
                LogToRevit($"[CONN-{connId}] Error handling HTTPS request from {clientEndpoint} after {totalDuration:F0}ms: {ex.Message}");
                LogToRevit($"[CONN-{connId}] Exception type: {ex.GetType().FullName}");
                LogToRevit($"[CONN-{connId}] Stack trace:\n{ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    LogToRevit($"[CONN-{connId}] Inner exception: {ex.InnerException.Message}");
                    LogToRevit($"[CONN-{connId}] Inner stack trace:\n{ex.InnerException.StackTrace}");
                }
                Interlocked.Increment(ref totalConnectionsFailed);
            }
            finally
            {
                Interlocked.Decrement(ref activeConnections);
                LogToRevit($"[CONN-{connId}] Connection cleanup complete (Active: {activeConnections}, Success: {totalConnectionsSuccessful}, Failed: {totalConnectionsFailed})");
            }
        }

        private async Task HandleHttpRequest(Stream stream, string connId, CancellationToken cancellationToken)
        {
            try
            {
                LogToRevit($"[CONN-{connId}] Reading HTTP request...");
                var requestStart = DateTime.Now;

                using (var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
                {
                    var requestLine = await reader.ReadLineAsync();
                    if (string.IsNullOrEmpty(requestLine))
                    {
                        LogToRevit($"[CONN-{connId}] Empty request line (client may have closed connection)");
                        return;
                    }

                    var parts = requestLine.Split(' ');
                    if (parts.Length < 2)
                    {
                        LogToRevit($"[CONN-{connId}] Invalid request line: {requestLine}");
                        return;
                    }

                    var method = parts[0];
                    var path = parts[1];
                    var httpVersion = parts.Length >= 3 ? parts[2] : "HTTP/1.1";
                    LogToRevit($"[CONN-{connId}] Request: {method} {path} {httpVersion}");

                    var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    string line;
                    int headerCount = 0;
                    while ((line = await reader.ReadLineAsync()) != null && line.Length > 0)
                    {
                        var colonIndex = line.IndexOf(':');
                        if (colonIndex > 0)
                        {
                            var key = line.Substring(0, colonIndex).Trim();
                            var value = line.Substring(colonIndex + 1).Trim();
                            headers[key] = value;
                            headerCount++;
                        }
                    }

                    LogToRevit($"[CONN-{connId}] Received {headerCount} headers");

                    // Read body
                    string code = "";
                    if (headers.ContainsKey("Content-Length"))
                    {
                        var contentLength = int.Parse(headers["Content-Length"]);
                        LogToRevit($"[CONN-{connId}] Content-Length: {contentLength}");

                        var buffer = new char[contentLength];
                        var totalRead = 0;
                        var bodyReadStart = DateTime.Now;

                        while (totalRead < contentLength)
                        {
                            var bytesRead = await reader.ReadAsync(buffer, totalRead, contentLength - totalRead);
                            if (bytesRead == 0)
                            {
                                LogToRevit($"[CONN-{connId}] WARNING: Stream ended before reading full body (expected {contentLength}, got {totalRead})");
                                break;
                            }
                            totalRead += bytesRead;
                        }

                        code = new string(buffer, 0, totalRead);
                        var bodyReadDuration = (DateTime.Now - bodyReadStart).TotalMilliseconds;
                        LogToRevit($"[CONN-{connId}] Received {totalRead}/{contentLength} bytes ({bodyReadDuration:F0}ms)");
                    }
                    else
                    {
                        LogToRevit($"[CONN-{connId}] No Content-Length header (body empty or chunked transfer)");
                    }

                    // Check authentication using shared token
                    string providedToken = null;
                    if (headers.ContainsKey("X-Auth-Token"))
                    {
                        providedToken = headers["X-Auth-Token"];
                        LogToRevit($"[CONN-{connId}] Auth method: X-Auth-Token header");
                    }
                    else if (headers.ContainsKey("Authorization"))
                    {
                        var authHeader = headers["Authorization"];
                        if (authHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
                        {
                            providedToken = authHeader.Substring(7).Trim();
                            LogToRevit($"[CONN-{connId}] Auth method: Authorization Bearer header");
                        }
                        else
                        {
                            LogToRevit($"[CONN-{connId}] WARNING: Unrecognized Authorization header format: {authHeader.Substring(0, Math.Min(20, authHeader.Length))}...");
                        }
                    }
                    else
                    {
                        LogToRevit($"[CONN-{connId}] WARNING: No authentication headers found");
                    }

                    var sharedToken = GetSharedAuthToken();
                    if (string.IsNullOrEmpty(providedToken))
                    {
                        LogToRevit($"[CONN-{connId}] Authentication failed: No token provided");
                        var errorResponse = new ScriptResponse
                        {
                            Success = false,
                            Error = "Authentication required. Provide X-Auth-Token header or Authorization: Bearer <token> header."
                        };
                        await SendHttpResponse(stream, errorResponse, 401, "Unauthorized");
                        return;
                    }
                    else if (string.IsNullOrEmpty(sharedToken))
                    {
                        LogToRevit($"[CONN-{connId}] Authentication failed: No shared token configured (server misconfiguration)");
                        var errorResponse = new ScriptResponse
                        {
                            Success = false,
                            Error = "Server authentication not configured."
                        };
                        await SendHttpResponse(stream, errorResponse, 500, "Internal Server Error");
                        return;
                    }
                    else if (providedToken != sharedToken)
                    {
                        LogToRevit($"[CONN-{connId}] Authentication failed: Token mismatch");
                        var errorResponse = new ScriptResponse
                        {
                            Success = false,
                            Error = "Invalid authentication token."
                        };
                        await SendHttpResponse(stream, errorResponse, 401, "Unauthorized");
                        return;
                    }

                    LogToRevit($"[CONN-{connId}] Authentication successful");

                    // Route request based on path
                    if (path.StartsWith("/roslyn", StringComparison.OrdinalIgnoreCase))
                    {
                        LogToRevit($"[CONN-{connId}] Routing to /roslyn endpoint");
                        await HandleRoslynRequest(stream, code, connId);
                    }
                    else if (path.StartsWith("/screenshot", StringComparison.OrdinalIgnoreCase))
                    {
                        LogToRevit($"[CONN-{connId}] Routing to /screenshot endpoint");
                        await HandleScreenshotRequest(stream, connId);
                    }
                    else if (path.StartsWith("/query/", StringComparison.OrdinalIgnoreCase))
                    {
                        LogToRevit($"[CONN-{connId}] Routing to /query endpoint: {path}");
                        await HandleQueryRequest(stream, path, code, connId);
                    }
                    else
                    {
                        LogToRevit($"[CONN-{connId}] Unknown endpoint: {path}");
                        await SendHttpResponse(stream, new ScriptResponse { Success = false, Error = "Unknown endpoint" }, 404, "Not Found");
                    }

                    var requestDuration = (DateTime.Now - requestStart).TotalMilliseconds;
                    LogToRevit($"[CONN-{connId}] Request processing complete ({requestDuration:F0}ms)");
                }
            }
            catch (System.Exception ex)
            {
                LogToRevit($"[CONN-{connId}] Error handling request: {ex.Message}");
                LogToRevit($"[CONN-{connId}] Exception type: {ex.GetType().FullName}");
                LogToRevit($"[CONN-{connId}] Stack trace: {ex.StackTrace}");
            }
        }

        private async Task HandleRoslynRequest(Stream stream, string code, string connId)
        {
            var response = new ScriptResponse();
            response.LogProcessing("Request received by server");

            if (string.IsNullOrWhiteSpace(code))
            {
                LogToRevit($"[CONN-{connId}] Empty request body");
                response.Success = false;
                response.Error = "Empty request body";
                response.LogProcessing("ERROR: Empty request body");
                await SendHttpResponse(stream, response);
                return;
            }

            LogToRevit($"[CONN-{connId}] Compiling script ({code.Length} chars)...");
            response.LogProcessing($"Starting compilation ({code.Length} chars)...");

            var compileStart = DateTime.Now;
            var compilationResult = await CompileScript(code);
            var compileDuration = (DateTime.Now - compileStart).TotalMilliseconds;

            if (!compilationResult.Success)
            {
                LogToRevit($"[CONN-{connId}] Compilation failed ({compileDuration:F0}ms)");
                compilationResult.ErrorResponse.LogProcessing($"Compilation FAILED ({compileDuration:F0}ms)");
                await SendHttpResponse(stream, compilationResult.ErrorResponse);
                return;
            }

            response.LogProcessing($"Compilation successful ({compileDuration:F0}ms)");
            LogToRevit($"[CONN-{connId}] Compilation successful ({compileDuration:F0}ms), executing via External Event...");
            response.LogProcessing("Queueing script for execution on Revit UI thread...");

            var executeStart = DateTime.Now;
            var scriptResponse = await ExecuteScriptViaExternalEvent(compilationResult.CompiledScript);
            var executeDuration = (DateTime.Now - executeStart).TotalMilliseconds;

            scriptResponse.LogProcessing($"Execution completed ({executeDuration:F0}ms) - Success: {scriptResponse.Success}");

            // Merge processing logs from compilation stage
            scriptResponse.ProcessingLog.InsertRange(0, response.ProcessingLog);

            LogToRevit($"[CONN-{connId}] Execution completed ({executeDuration:F0}ms) - Success: {scriptResponse.Success}");
            await SendHttpResponse(stream, scriptResponse);
        }

        private class CompilationResult
        {
            public bool Success { get; set; }
            public Script<object> CompiledScript { get; set; }
            public ScriptResponse ErrorResponse { get; set; }
        }

        private Task<CompilationResult> CompileScript(string code)
        {
            var result = new CompilationResult();

            try
            {
                // Use MetadataReference.CreateFromFile to explicitly load from our bin directory
                // This bypasses AppDomain assembly loading and avoids conflicts with pyRevit
                string binDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                var scriptOptions = ScriptOptions.Default
                    .AddReferences(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(typeof(ScriptGlobals).Assembly.Location))
                    .AddReferences(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(typeof(Document).Assembly.Location))
                    .AddReferences(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(typeof(UIDocument).Assembly.Location))
                    .AddReferences(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(typeof(System.Linq.Enumerable).Assembly.Location));

                // Explicitly add Roslyn dependency DLLs from bin directory (all Revit versions)
                // This avoids conflicts with incompatible versions loaded by other addins (e.g., pyRevit)
                string immutablePath = Path.Combine(binDirectory, "System.Collections.Immutable.dll");
                string metadataPath = Path.Combine(binDirectory, "System.Reflection.Metadata.dll");

                if (File.Exists(immutablePath))
                    scriptOptions = scriptOptions.AddReferences(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(immutablePath));
                if (File.Exists(metadataPath))
                    scriptOptions = scriptOptions.AddReferences(Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(metadataPath));

                scriptOptions = scriptOptions
                    .AddImports("System")
                    .AddImports("System.Linq")
                    .AddImports("System.Collections.Generic")
                    .AddImports("Autodesk.Revit.DB")
                    .AddImports("Autodesk.Revit.UI");

                var script = CSharpScript.Create(code, scriptOptions, typeof(ScriptGlobals));
                var compilation = script.GetCompilation();
                var diagnostics = compilation.GetDiagnostics();

                var errors = diagnostics.Where(d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Error).ToList();
                if (errors.Any())
                {
                    result.Success = false;
                    result.ErrorResponse = new ScriptResponse
                    {
                        Success = false,
                        Error = "Compilation failed",
                        Diagnostics = errors.Select(e => e.ToString()).ToList()
                    };
                    return Task.FromResult(result);
                }

                result.Success = true;
                result.CompiledScript = script;

                var warnings = diagnostics.Where(d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Warning).ToList();
                if (warnings.Any())
                {
                    result.ErrorResponse = new ScriptResponse
                    {
                        Success = true,
                        Diagnostics = warnings.Select(w => w.ToString()).ToList()
                    };
                }
            }
            catch (CompilationErrorException ex)
            {
                result.Success = false;
                result.ErrorResponse = new ScriptResponse
                {
                    Success = false,
                    Error = "Compilation error",
                    Diagnostics = ex.Diagnostics.Select(d => d.ToString()).ToList()
                };
            }
            catch (System.Exception ex)
            {
                LogToRevit($"[{DateTime.Now:HH:mm:ss}] Script compilation exception: {ex.Message}");
                LogToRevit($"[{DateTime.Now:HH:mm:ss}] Exception type: {ex.GetType().FullName}");
                LogToRevit($"[{DateTime.Now:HH:mm:ss}] Stack trace:\n{ex.StackTrace}");

                result.Success = false;
                result.ErrorResponse = new ScriptResponse
                {
                    Success = false,
                    Error = $"{ex.Message}\n\nException Type: {ex.GetType().Name}\nStack Trace: {ex.StackTrace}"
                };
            }

            return Task.FromResult(result);
        }

        private async Task<ScriptResponse> ExecuteScriptViaExternalEvent(Script<object> script)
        {
            var tcs = new TaskCompletionSource<ScriptResponse>();

            // Queue the script execution request in the handler
            scriptHandler.QueueExecution(script, tcs);

            // Raise the external event to execute on the UI thread
            scriptExecutionEvent.Raise();

            // Add timeout to prevent infinite waiting (30 seconds default)
            var timeoutTask = Task.Delay(30000);
            var completedTask = await Task.WhenAny(tcs.Task, timeoutTask);

            if (completedTask == timeoutTask)
            {
                var response = new ScriptResponse
                {
                    Success = false,
                    Error = "Script execution timeout (30 seconds). The script may contain an infinite loop or is waiting for user input."
                };
                response.LogProcessing("ERROR: Execution timeout after 30 seconds");
                return response;
            }

            return await tcs.Task;
        }

        private async Task HandleScreenshotRequest(Stream stream, string connId)
        {
            try
            {
                LogToRevit($"[CONN-{connId}] Capturing screenshot...");
                var screenshotStart = DateTime.Now;
                var screenshotPath = await CaptureScreenshot();
                var screenshotDuration = (DateTime.Now - screenshotStart).TotalMilliseconds;

                LogToRevit($"[CONN-{connId}] Screenshot captured ({screenshotDuration:F0}ms): {screenshotPath}");

                var response = new ScriptResponse
                {
                    Success = true,
                    Output = screenshotPath
                };
                await SendHttpResponse(stream, response);
            }
            catch (System.Exception ex)
            {
                LogToRevit($"[CONN-{connId}] Screenshot request exception: {ex.Message}");
                LogToRevit($"[CONN-{connId}] Exception type: {ex.GetType().FullName}");
                LogToRevit($"[CONN-{connId}] Stack trace:\n{ex.StackTrace}");

                await SendHttpResponse(stream, new ScriptResponse
                {
                    Success = false,
                    Error = $"{ex.Message}\n\nException Type: {ex.GetType().Name}"
                });
            }
        }

        private async Task HandleQueryRequest(Stream stream, string path, string body, string connId)
        {
            var response = new ScriptResponse();
            response.LogProcessing("Query request received by server");

            try
            {
                LogToRevit($"[CONN-{connId}] Processing query request: {path}");
                response.LogProcessing($"Request path: {path}");

                var executeStart = DateTime.Now;
                var queryResponse = await ExecuteQueryViaExternalEvent(path, body);
                var executeDuration = (DateTime.Now - executeStart).TotalMilliseconds;

                queryResponse.LogProcessing($"Execution completed ({executeDuration:F0}ms) - Success: {queryResponse.Success}");
                queryResponse.ProcessingLog.InsertRange(0, response.ProcessingLog);

                LogToRevit($"[CONN-{connId}] Query completed ({executeDuration:F0}ms) - Success: {queryResponse.Success}");
                await SendHttpResponse(stream, queryResponse);
            }
            catch (System.Exception ex)
            {
                LogToRevit($"[CONN-{connId}] Query request exception: {ex.Message}");
                LogToRevit($"[CONN-{connId}] Exception type: {ex.GetType().FullName}");
                LogToRevit($"[CONN-{connId}] Stack trace:\n{ex.StackTrace}");

                response.Success = false;
                response.Error = $"{ex.Message}\n\nException Type: {ex.GetType().Name}";
                await SendHttpResponse(stream, response);
            }
        }

        private async Task<ScriptResponse> ExecuteQueryViaExternalEvent(string path, string body)
        {
            var tcs = new TaskCompletionSource<ScriptResponse>();

            // Queue the query execution request in the handler
            queryHandler.QueueExecution(path, body, tcs);

            // Raise the external event to execute on the UI thread
            queryEvent.Raise();

            // Add timeout to prevent infinite waiting (30 seconds default)
            var timeoutTask = Task.Delay(30000);
            var completedTask = await Task.WhenAny(tcs.Task, timeoutTask);

            if (completedTask == timeoutTask)
            {
                var response = new ScriptResponse
                {
                    Success = false,
                    Error = "Query execution timeout (30 seconds)."
                };
                response.LogProcessing("ERROR: Execution timeout after 30 seconds");
                return response;
            }

            return await tcs.Task;
        }

        private Task<string> CaptureScreenshot()
        {
            var tcs = new TaskCompletionSource<string>();

            // Queue the screenshot request in the handler
            screenshotHandler.QueueCapture(tcs);

            // Raise the external event to execute on the UI thread
            screenshotEvent.Raise();

            return tcs.Task;
        }

        private async Task SendHttpResponse(Stream stream, ScriptResponse scriptResponse, int statusCode = 0, string statusText = null)
        {
            var jsonResponse = scriptResponse.ToJson();
            var responseBytes = Encoding.UTF8.GetBytes(jsonResponse);

            if (statusCode == 0)
            {
                statusCode = scriptResponse.Success ? 200 : 400;
            }

            if (statusText == null)
            {
                statusText = scriptResponse.Success ? "OK" : "Bad Request";
            }

            var httpResponse = $"HTTP/1.1 {statusCode} {statusText}\r\n" +
                             $"Content-Type: application/json\r\n" +
                             $"Content-Length: {responseBytes.Length}\r\n" +
                             $"\r\n";

            var headerBytes = Encoding.UTF8.GetBytes(httpResponse);

            await stream.WriteAsync(headerBytes, 0, headerBytes.Length);
            await stream.WriteAsync(responseBytes, 0, responseBytes.Length);
            await stream.FlushAsync();
        }
    }

    /// <summary>
    /// External Event Handler for executing Roslyn scripts on the Revit UI thread
    /// </summary>
    internal class ScriptExecutionHandler : IExternalEventHandler
    {
        private Script<object> pendingScript;
        private TaskCompletionSource<ScriptResponse> pendingTcs;
        private readonly object lockObject = new object();

        public ScriptExecutionHandler()
        {
        }

        public void QueueExecution(Script<object> script, TaskCompletionSource<ScriptResponse> tcs)
        {
            lock (lockObject)
            {
                pendingScript = script;
                pendingTcs = tcs;
            }
        }

        public void Execute(UIApplication app)
        {
            Script<object> script;
            TaskCompletionSource<ScriptResponse> tcs;

            // Get the pending request
            lock (lockObject)
            {
                script = pendingScript;
                tcs = pendingTcs;
                pendingScript = null;
                pendingTcs = null;
            }

            if (script == null || tcs == null)
            {
                return; // No pending execution
            }

            var response = new ScriptResponse();
            var outputCapture = new StringWriter();
            var originalOut = Console.Out;

            try
            {
                response.LogProcessing("Script execution started on Revit UI thread");
                Console.SetOut(outputCapture);

                var uidoc = app.ActiveUIDocument;
                var doc = uidoc?.Document;

                // Allow scripts to run even without an active document
                // UIDoc and Doc will be null, but UIApp is always available
                var globals = new ScriptGlobals
                {
                    UIApp = app,
                    UIDoc = uidoc,
                    Doc = doc
                };

                response.LogProcessing($"Script context: UIApp={app != null}, UIDoc={uidoc != null}, Doc={doc != null}");
                response.LogProcessing("Running script...");

                var executeStart = DateTime.Now;
                // Execute the script (no Transaction - user must create one in their script if needed)
                var result = script.RunAsync(globals).Result;
                var executeMs = (DateTime.Now - executeStart).TotalMilliseconds;

                response.LogProcessing($"Script completed successfully ({executeMs:F0}ms)");
                response.Success = true;
                response.Output = outputCapture.ToString();

                if (result.ReturnValue != null)
                {
                    if (!string.IsNullOrEmpty(response.Output))
                        response.Output += "\n";
                    response.Output += $"Return value: {result.ReturnValue}";
                }

                tcs.SetResult(response);
            }
            catch (CompilationErrorException ex)
            {
                response.LogProcessing("ERROR: Compilation error during execution");
                response.Success = false;
                response.Error = "Compilation error";
                response.Diagnostics = ex.Diagnostics.Select(d => d.ToString()).ToList();
                tcs.SetResult(response);
            }
            catch (System.Exception ex)
            {
                RevitBalletServer.LogToRevit($"[{DateTime.Now:HH:mm:ss}] Script execution exception: {ex.Message}");
                RevitBalletServer.LogToRevit($"[{DateTime.Now:HH:mm:ss}] Exception type: {ex.GetType().FullName}");
                RevitBalletServer.LogToRevit($"[{DateTime.Now:HH:mm:ss}] Stack trace:\n{ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    RevitBalletServer.LogToRevit($"[{DateTime.Now:HH:mm:ss}] Inner exception: {ex.InnerException.Message}");
                    RevitBalletServer.LogToRevit($"[{DateTime.Now:HH:mm:ss}] Inner stack trace:\n{ex.InnerException.StackTrace}");
                }

                response.LogProcessing($"ERROR: {ex.GetType().Name} - {ex.Message}");
                response.Success = false;
                response.Error = $"{ex.Message}\n\nException Type: {ex.GetType().Name}\n\nStack Trace:\n{ex.StackTrace}";
                response.Output = outputCapture.ToString();
                tcs.SetResult(response);
            }
            finally
            {
                Console.SetOut(originalOut);
            }
        }

        public string GetName()
        {
            return "RevitBalletScriptExecutionHandler";
        }
    }

    /// <summary>
    /// External Event Handler for capturing screenshots on the Revit UI thread
    /// </summary>
    internal class ScreenshotHandler : IExternalEventHandler
    {
        private TaskCompletionSource<string> pendingTcs;
        private readonly string sessionId;
        private readonly object lockObject = new object();

        public ScreenshotHandler(string sessionId)
        {
            this.sessionId = sessionId;
        }

        public void QueueCapture(TaskCompletionSource<string> tcs)
        {
            lock (lockObject)
            {
                pendingTcs = tcs;
            }
        }

        public void Execute(UIApplication app)
        {
            TaskCompletionSource<string> tcs;

            // Get the pending request
            lock (lockObject)
            {
                tcs = pendingTcs;
                pendingTcs = null;
            }

            if (tcs == null)
            {
                return; // No pending capture
            }

            try
            {
                // Get Revit main application window handle
                IntPtr windowHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

                // Create screenshots directory
                string screenshotDir = PathHelper.EnsureRuntimeSubdirectoryExists("screenshots");

                // Clean up old screenshots (keep last 20)
                CleanupOldScreenshots(screenshotDir, 20);

                // Generate filename: datetime-sessionid.png
                string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss-fff");
                string filename = $"{timestamp}-{sessionId}.png";
                string filepath = Path.Combine(screenshotDir, filename);

                // Capture screenshot using Windows API
                CaptureWindowToFile(windowHandle, filepath);

                // Convert absolute path to relative path from revit-ballet root
                string relativePath = ConvertToRelativePath(filepath);
                tcs.SetResult(relativePath);
            }
            catch (System.Exception ex)
            {
                tcs.SetException(ex);
            }
        }

        public string GetName()
        {
            return "RevitBalletScreenshotHandler";
        }

        private string ConvertToRelativePath(string absolutePath)
        {
            // Convert absolute path to relative path from revit-ballet root
            // e.g., C:\Users\...\AppData\Roaming\revit-ballet\runtime\screenshots\file.png
            // becomes revit-ballet\runtime\screenshots\file.png
            string revitBalletBase = PathHelper.RevitBalletDirectory;

            if (absolutePath.StartsWith(revitBalletBase, StringComparison.OrdinalIgnoreCase))
            {
                // Remove the base path and leading separator
                string relativePart = absolutePath.Substring(revitBalletBase.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                // Prepend "revit-ballet\"
                return Path.Combine("revit-ballet", relativePart);
            }

            // If path doesn't start with revit-ballet base, return as-is
            return absolutePath;
        }

        private void CaptureWindowToFile(IntPtr windowHandle, string filepath)
        {
            // Get window client rectangle (excludes borders)
            RECT rect;
            GetClientRect(windowHandle, out rect);
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            // Create bitmap
            using (var bitmap = new System.Drawing.Bitmap(width, height))
            {
                using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                {
                    // Get device context from the window
                    IntPtr windowDC = GetDC(windowHandle);
                    IntPtr targetDC = graphics.GetHdc();

                    try
                    {
                        // Copy pixels from window DC to bitmap DC
                        BitBlt(targetDC, 0, 0, width, height, windowDC, 0, 0, SRCCOPY);
                    }
                    finally
                    {
                        graphics.ReleaseHdc(targetDC);
                        ReleaseDC(windowHandle, windowDC);
                    }
                }

                // Save to file
                bitmap.Save(filepath, System.Drawing.Imaging.ImageFormat.Png);
            }
        }

        private void CleanupOldScreenshots(string directory, int keepCount)
        {
            try
            {
                var files = new DirectoryInfo(directory)
                    .GetFiles("*.png")
                    .OrderByDescending(f => f.CreationTime)
                    .Skip(keepCount)
                    .ToList();

                foreach (var file in files)
                {
                    try
                    {
                        file.Delete();
                    }
                    catch
                    {
                        // Ignore deletion failures
                    }
                }
            }
            catch
            {
                // Ignore cleanup failures
            }
        }

        // Windows API declarations for screenshot capture
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest, IntPtr hdcSource, int xSrc, int ySrc, int RasterOp);

        private const int SRCCOPY = 0x00CC0020;

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }

    /// <summary>
    /// External Event Handler for executing pre-compiled queries on the Revit UI thread
    /// </summary>
    internal class QueryHandler : IExternalEventHandler
    {
        private string pendingPath;
        private string pendingBody;
        private TaskCompletionSource<ScriptResponse> pendingTcs;
        private readonly object lockObject = new object();

        public void QueueExecution(string path, string body, TaskCompletionSource<ScriptResponse> tcs)
        {
            lock (lockObject)
            {
                pendingPath = path;
                pendingBody = body;
                pendingTcs = tcs;
            }
        }

        public void Execute(UIApplication app)
        {
            string path;
            string body;
            TaskCompletionSource<ScriptResponse> tcs;

            // Get the pending request
            lock (lockObject)
            {
                path = pendingPath;
                body = pendingBody;
                tcs = pendingTcs;
                pendingPath = null;
                pendingBody = null;
                pendingTcs = null;
            }

            if (path == null || tcs == null)
            {
                return; // No pending execution
            }

            var response = new ScriptResponse();
            var outputCapture = new StringWriter();
            var originalOut = Console.Out;

            try
            {
                response.LogProcessing("Query execution started on Revit UI thread");
                Console.SetOut(outputCapture);

                var executeStart = DateTime.Now;

                // Parse request body (JSON with documentTitle)
                var requestData = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(body ?? "{}");
                string documentTitle = requestData.ContainsKey("documentTitle") ? requestData["documentTitle"] : null;

                response.LogProcessing($"Query context: DocumentTitle={documentTitle}");

                // Find target document
                Document targetDoc = null;
                if (!string.IsNullOrEmpty(documentTitle))
                {
                    foreach (Document d in app.Application.Documents)
                    {
                        if (!d.IsLinked && d.Title == documentTitle)
                        {
                            targetDoc = d;
                            break;
                        }
                    }
                }

                if (targetDoc == null && !string.IsNullOrEmpty(documentTitle))
                {
                    response.Success = false;
                    response.Error = $"Document not found: {documentTitle}";
                    response.Output = outputCapture.ToString();
                    tcs.SetResult(response);
                    return;
                }

                // Route to appropriate query handler
                if (path.Equals("/query/familytypes/counts", StringComparison.OrdinalIgnoreCase))
                {
                    QueryFamilyTypeCounts(targetDoc, response);
                }
                else if (path.Equals("/query/familytypes/elements", StringComparison.OrdinalIgnoreCase))
                {
                    QueryFamilyTypeElements(targetDoc, requestData, response);
                }
                else if (path.Equals("/query/categories/counts", StringComparison.OrdinalIgnoreCase))
                {
                    QueryCategoryCounts(targetDoc, response);
                }
                else if (path.Equals("/query/categories/elements", StringComparison.OrdinalIgnoreCase))
                {
                    QueryCategoryElements(targetDoc, requestData, response);
                }
                else if (path.Equals("/query/worksets/counts", StringComparison.OrdinalIgnoreCase))
                {
                    QueryWorksetCounts(targetDoc, response);
                }
                else if (path.Equals("/query/worksets/elements", StringComparison.OrdinalIgnoreCase))
                {
                    QueryWorksetElements(targetDoc, requestData, response);
                }
                else
                {
                    response.Success = false;
                    response.Error = $"Unknown query endpoint: {path}";
                }

                var executeMs = (DateTime.Now - executeStart).TotalMilliseconds;
                response.LogProcessing($"Query completed ({executeMs:F0}ms)");
                response.Output = outputCapture.ToString();
                tcs.SetResult(response);
            }
            catch (System.Exception ex)
            {
                RevitBalletServer.LogToRevit($"[{DateTime.Now:HH:mm:ss}] Query execution exception: {ex.Message}");
                RevitBalletServer.LogToRevit($"[{DateTime.Now:HH:mm:ss}] Exception type: {ex.GetType().FullName}");
                RevitBalletServer.LogToRevit($"[{DateTime.Now:HH:mm:ss}] Stack trace:\n{ex.StackTrace}");

                response.LogProcessing($"ERROR: {ex.GetType().Name} - {ex.Message}");
                response.Success = false;
                response.Error = $"{ex.Message}\n\nException Type: {ex.GetType().Name}\n\nStack Trace:\n{ex.StackTrace}";
                response.Output = outputCapture.ToString();
                tcs.SetResult(response);
            }
            finally
            {
                Console.SetOut(originalOut);
            }
        }

        private void QueryFamilyTypeCounts(Document doc, ScriptResponse response)
        {
            var elementTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(ElementType))
                .Cast<ElementType>()
                .ToList();

            // Count instances by type
            var typeInstanceCounts = new Dictionary<ElementId, int>();
            var allInstances = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Where(x => x.GetTypeId() != null && x.GetTypeId() != ElementId.InvalidElementId)
                .ToList();

            foreach (var instance in allInstances)
            {
                ElementId typeId = instance.GetTypeId();
                if (!typeInstanceCounts.ContainsKey(typeId))
                    typeInstanceCounts[typeId] = 0;
                typeInstanceCounts[typeId]++;
            }

            foreach (var elementType in elementTypes)
            {
                string typeName = elementType.Name;
                string familyName = "";
                string categoryName = "";

                if (elementType is FamilySymbol fs)
                {
                    familyName = fs.Family.Name;
                    categoryName = fs.Category != null ? fs.Category.Name : "N/A";
                    if (categoryName.Contains("Import Symbol") || familyName.Contains("Import Symbol"))
                        continue;
                }
                else
                {
                    Parameter familyParam = elementType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                    familyName = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
                        ? familyParam.AsString()
                        : "System Type";
                    categoryName = elementType.Category != null ? elementType.Category.Name : "N/A";
                    if (categoryName.Contains("Import Symbol") || familyName.Contains("Import Symbol"))
                        continue;
                }

                int count = typeInstanceCounts.ContainsKey(elementType.Id) ? typeInstanceCounts[elementType.Id] : 0;
                Console.WriteLine($"FAMILYTYPE|{categoryName}|{familyName}|{typeName}|{count}");
            }

            response.Success = true;
        }

        private void QueryFamilyTypeElements(Document doc, Dictionary<string, string> requestData, ScriptResponse response)
        {
            // Parse family type keys from request
            string familyTypesJson = requestData.ContainsKey("familyTypes") ? requestData["familyTypes"] : "[]";
            var familyTypes = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(familyTypesJson);

            // Build lookup of matching type IDs
            var matchingTypeIds = new HashSet<ElementId>();
            var typeIdToKey = new Dictionary<ElementId, Tuple<string, string, string>>();

            var elementTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(ElementType))
                .Cast<ElementType>()
                .ToList();

            foreach (var elementType in elementTypes)
            {
                string typeName = elementType.Name;
                string familyName = "";
                string categoryName = "";

                if (elementType is FamilySymbol fs)
                {
                    familyName = fs.Family.Name;
                    categoryName = fs.Category != null ? fs.Category.Name : "N/A";
                }
                else
                {
                    Parameter familyParam = elementType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                    familyName = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
                        ? familyParam.AsString()
                        : "System Type";
                    categoryName = elementType.Category != null ? elementType.Category.Name : "N/A";
                }

                var key = Tuple.Create(categoryName, familyName, typeName);
                if (familyTypes.Any(ft => ft["category"] == key.Item1 && ft["family"] == key.Item2 && ft["typeName"] == key.Item3))
                {
                    matchingTypeIds.Add(elementType.Id);
                    typeIdToKey[elementType.Id] = key;
                }
            }

            // Collect instances
            var instances = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Where(x => x.GetTypeId() != null && matchingTypeIds.Contains(x.GetTypeId()))
                .ToList();

            foreach (var elem in instances)
            {
                var key = typeIdToKey[elem.GetTypeId()];
                Console.WriteLine($"ELEMENT|{key.Item1}|{key.Item2}|{key.Item3}|{elem.UniqueId}|{elem.Id.IntegerValue}");
            }

            response.Success = true;
        }

        private void QueryCategoryCounts(Document doc, ScriptResponse response)
        {
            var collector = new FilteredElementCollector(doc);
            var elements = collector.WhereElementIsNotElementType();
            var categoryGroups = elements.Where(e => e.Category != null).GroupBy(e => e.Category.Name).OrderBy(g => g.Key);

            foreach (var group in categoryGroups)
            {
                Console.WriteLine($"CATEGORY|{group.Key}|{group.Count()}");
            }

            response.Success = true;
        }

        private void QueryCategoryElements(Document doc, Dictionary<string, string> requestData, ScriptResponse response)
        {
            // Parse category names from request
            string categoriesJson = requestData.ContainsKey("categories") ? requestData["categories"] : "[]";
            var categories = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(categoriesJson);
            var categorySet = new HashSet<string>(categories);

            var collector = new FilteredElementCollector(doc);
            var elements = collector.WhereElementIsNotElementType().Where(e => e.Category != null && categorySet.Contains(e.Category.Name));
            var categoryGroups = elements.GroupBy(e => e.Category.Name);

            foreach (var group in categoryGroups)
            {
                Console.WriteLine($"CATEGORY|{group.Key}");
                foreach (var elem in group)
                {
                    Console.WriteLine($"ELEMENT|{elem.UniqueId}|{elem.Id.IntegerValue}");
                }
            }

            response.Success = true;
        }

        private void QueryWorksetCounts(Document doc, ScriptResponse response)
        {
            if (!doc.IsWorkshared)
            {
                response.Success = true;
                response.Output = "INFO|Document is not workshared";
                return;
            }

            var collector = new FilteredElementCollector(doc).WhereElementIsNotElementType();
            var worksetElementCounts = new Dictionary<WorksetId, int>();

            foreach (Element e in collector)
            {
                WorksetId wsId = e.WorksetId;
                if (wsId == WorksetId.InvalidWorksetId) continue;
                if (!worksetElementCounts.ContainsKey(wsId))
                    worksetElementCounts[wsId] = 0;
                worksetElementCounts[wsId]++;
            }

            foreach (var pair in worksetElementCounts)
            {
                Workset ws = doc.GetWorksetTable().GetWorkset(pair.Key);
                if (ws == null) continue;

                string wsType = "";
                switch (ws.Kind)
                {
                    case WorksetKind.UserWorkset: wsType = "User"; break;
                    case WorksetKind.StandardWorkset: wsType = "Standard"; break;
                    case WorksetKind.FamilyWorkset: wsType = "Family"; break;
                    case WorksetKind.ViewWorkset: wsType = "View"; break;
                    default: wsType = ws.Kind.ToString(); break;
                }

                Console.WriteLine($"WORKSET|{ws.Name}|{wsType}|{pair.Value}");
            }

            response.Success = true;
        }

        private void QueryWorksetElements(Document doc, Dictionary<string, string> requestData, ScriptResponse response)
        {
            if (!doc.IsWorkshared)
            {
                response.Success = true;
                response.Output = "INFO|Document is not workshared";
                return;
            }

            // Parse workset names from request
            string worksetsJson = requestData.ContainsKey("worksets") ? requestData["worksets"] : "[]";
            var worksets = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(worksetsJson);
            var worksetNames = new HashSet<string>(worksets);

            var collector = new FilteredElementCollector(doc).WhereElementIsNotElementType();

            foreach (Element e in collector)
            {
                WorksetId wsId = e.WorksetId;
                if (wsId == WorksetId.InvalidWorksetId) continue;

                Workset ws = doc.GetWorksetTable().GetWorkset(wsId);
                if (ws == null || !worksetNames.Contains(ws.Name)) continue;

                Console.WriteLine($"ELEMENT|{ws.Name}|{e.UniqueId}|{e.Id.IntegerValue}");
            }

            response.Success = true;
        }

        public string GetName()
        {
            return "RevitBalletQueryHandler";
        }
    }

    /// <summary>
    /// Represents a single document entry in the document registry.
    /// One entry per open document across all Revit sessions.
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
    }

    /// <summary>
    /// Globals available to Roslyn scripts
    /// </summary>
    public class ScriptGlobals
    {
        public UIApplication UIApp { get; set; }
        public UIDocument UIDoc { get; set; }
        public Document Doc { get; set; }
    }

    /// <summary>
    /// Response object for script execution
    /// </summary>
    public class ScriptResponse
    {
        public bool Success { get; set; }
        public string Output { get; set; }
        public string Error { get; set; }
        public List<string> Diagnostics { get; set; }
        public List<string> ProcessingLog { get; set; }

        public ScriptResponse()
        {
            Diagnostics = new List<string>();
            ProcessingLog = new List<string>();
        }

        public void LogProcessing(string message)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            ProcessingLog.Add($"[{timestamp}] {message}");
        }

        public string ToJson()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    /// <summary>
    /// Helper class for generating and managing self-signed SSL certificates for HTTPS server
    /// </summary>
    internal static class SslCertificateHelper
    {
        private static readonly string CertificateStorePath = PathHelper.GetRuntimeFilePath("ServerCert.pfx");

        private const string CertificatePassword = "revit-ballet-localhost";

        public static X509Certificate2 GetOrCreateCertificate()
        {
            if (File.Exists(CertificateStorePath))
            {
                try
                {
                    var cert = new X509Certificate2(CertificateStorePath, CertificatePassword,
                        X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);

                    if (cert.NotAfter > DateTime.Now.AddDays(30))
                    {
                        return cert;
                    }

                    cert.Dispose();
                }
                catch { }
            }

            var certificate = GenerateSelfSignedCertificate();

            try
            {
                var certBytes = certificate.Export(X509ContentType.Pfx, CertificatePassword);
                File.WriteAllBytes(CertificateStorePath, certBytes);
            }
            catch { }

            return certificate;
        }

        private static X509Certificate2 GenerateSelfSignedCertificate()
        {
#if NET48 || NET8_0_OR_GREATER
            var rsa = RSA.Create(2048);
            try
            {
                var request = new CertificateRequest(
                    "CN=localhost",
                    rsa,
                    HashAlgorithmName.SHA256,
                    RSASignaturePadding.Pkcs1
                );

                var sanBuilder = new SubjectAlternativeNameBuilder();
                sanBuilder.AddDnsName("localhost");
                sanBuilder.AddIpAddress(System.Net.IPAddress.Loopback);
                sanBuilder.AddIpAddress(System.Net.IPAddress.Parse("127.0.0.1"));
                request.CertificateExtensions.Add(sanBuilder.Build());

                request.CertificateExtensions.Add(
                    new X509KeyUsageExtension(
                        X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
                        critical: true
                    )
                );

                request.CertificateExtensions.Add(
                    new X509EnhancedKeyUsageExtension(
                        new OidCollection { new Oid("1.3.6.1.5.5.7.3.1") },
                        critical: true
                    )
                );

                var certificate = request.CreateSelfSigned(
                    DateTimeOffset.Now.AddDays(-1),
                    DateTimeOffset.Now.AddDays(365)
                );

                var certBytes = certificate.Export(X509ContentType.Pfx, CertificatePassword);
                certificate.Dispose();

                var finalCertificate = new X509Certificate2(certBytes, CertificatePassword,
                    X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.MachineKeySet);

                return finalCertificate;
            }
            finally
            {
                rsa?.Dispose();
            }
#else
            throw new NotSupportedException(
                "Self-signed certificate generation is not supported on .NET Framework 4.6/4.7.\n" +
                "For Revit 2017-2018, please generate a certificate manually using:\n" +
                "  PowerShell: New-SelfSignedCertificate -DnsName 'localhost' -CertStoreLocation 'Cert:\\CurrentUser\\My'\n" +
                "Then export it to: " + CertificateStorePath
            );
#endif
        }
    }
}

#endif
