#if REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
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
        private const string SESSIONS_FILE = "sessions";
        private const string TOKEN_FILE = "token";

        private TcpListener listener;
        private CancellationTokenSource cancellationTokenSource;
        private bool isRunning = false;
        private int serverPort;
        private X509Certificate2 serverCertificate;
        private string sessionId;
        private static UIApplication uiApp;

        public bool IsRunning => isRunning;
        public int Port => serverPort;
        public string SessionId => sessionId;

        public static void SetUIApplication(UIApplication application)
        {
            uiApp = application;
        }

        private void LogToRevit(string message)
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

            sessionId = Guid.NewGuid().ToString();

            // Load or generate self-signed certificate for HTTPS
            try
            {
                serverCertificate = SslCertificateHelper.GetOrCreateCertificate();
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
            isRunning = true;

            // Register in network registry
            RegisterInNetwork();

            // Start background listener thread
            Task.Run(() => AcceptRequestsLoop(cancellationTokenSource.Token));

            // Start heartbeat to keep network registry alive
            Task.Run(() => UpdateNetworkHeartbeat(cancellationTokenSource.Token));

            LogToRevit($"[Revit Ballet] Server started on port {serverPort} (Session: {sessionId.Substring(0, 8)})");
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
            var entry = CreateNetworkEntry();

            lock (typeof(RevitBalletServer))
            {
                var entries = ReadNetworkRegistry();
                entries.Add(entry);
                WriteNetworkRegistry(entries);
            }
        }

        private void UnregisterFromNetwork()
        {
            lock (typeof(RevitBalletServer))
            {
                var entries = ReadNetworkRegistry();
                entries.RemoveAll(e => e.SessionId == sessionId);
                WriteNetworkRegistry(entries);
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
                        var entries = ReadNetworkRegistry();
                        var myEntry = entries.FirstOrDefault(e => e.SessionId == sessionId);
                        if (myEntry != null)
                        {
                            myEntry.LastHeartbeat = DateTime.Now;
                            myEntry.Documents = GetOpenDocuments();
                            WriteNetworkRegistry(entries);
                        }
                        else
                        {
                            // Re-register if not found
                            entries.Add(CreateNetworkEntry());
                            WriteNetworkRegistry(entries);
                        }

                        // Clean up dead sessions (no heartbeat for 2 minutes)
                        var deadEntries = entries.Where(e =>
                            (DateTime.Now - e.LastHeartbeat).TotalSeconds > 120).ToList();
                        if (deadEntries.Count > 0)
                        {
                            foreach (var dead in deadEntries)
                            {
                                entries.Remove(dead);
                            }
                            WriteNetworkRegistry(entries);
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch { }
            }
        }

        private NetworkEntry CreateNetworkEntry()
        {
            return new NetworkEntry
            {
                SessionId = sessionId,
                Port = serverPort,
                Hostname = Environment.MachineName,
                ProcessId = System.Diagnostics.Process.GetCurrentProcess().Id,
                Documents = GetOpenDocuments(),
                RegisteredAt = DateTime.Now,
                LastHeartbeat = DateTime.Now
            };
        }

        private List<string> GetOpenDocuments()
        {
            var docs = new List<string>();
            try
            {
                if (uiApp != null)
                {
                    foreach (Document doc in uiApp.Application.Documents)
                    {
                        docs.Add(doc.Title);
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

        private string GetSessionsFilePath()
        {
            return Path.Combine(GetNetworkFolderPath(), SESSIONS_FILE);
        }

        private string GetTokenFilePath()
        {
            return Path.Combine(GetNetworkFolderPath(), TOKEN_FILE);
        }

        private List<NetworkEntry> ReadNetworkRegistry()
        {
            var path = GetSessionsFilePath();
            var entries = new List<NetworkEntry>();

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
                    if (parts.Length >= 6)
                    {
                        var entry = new NetworkEntry
                        {
                            SessionId = parts[0],
                            Port = int.Parse(parts[1]),
                            Hostname = parts[2],
                            ProcessId = int.Parse(parts[3]),
                            RegisteredAt = DateTime.Parse(parts[4]),
                            LastHeartbeat = DateTime.Parse(parts[5]),
                            Documents = new List<string>()
                        };

                        // Documents are all remaining parts (comma-separated)
                        for (int i = 6; i < parts.Length; i++)
                        {
                            if (!string.IsNullOrWhiteSpace(parts[i]))
                            {
                                entry.Documents.Add(parts[i]);
                            }
                        }

                        entries.Add(entry);
                    }
                }
            }
            catch { }

            return entries;
        }

        private void WriteNetworkRegistry(List<NetworkEntry> entries)
        {
            var path = GetSessionsFilePath();
            var lines = new List<string>();

            lines.Add("# Revit Ballet Network Registry - Sessions");
            lines.Add("# SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,Documents...");

            foreach (var entry in entries)
            {
                var docsString = string.Join(",", entry.Documents);
                var line = $"{entry.SessionId},{entry.Port},{entry.Hostname},{entry.ProcessId},{entry.RegisteredAt:O},{entry.LastHeartbeat:O},{docsString}";
                lines.Add(line);
            }

            File.WriteAllLines(path, lines);
        }

        public static List<NetworkEntry> GetActiveNetworkSessions()
        {
            var networkDir = PathHelper.EnsureRuntimeSubdirectoryExists(NETWORK_FOLDER);
            var path = Path.Combine(networkDir, SESSIONS_FILE);

            var entries = new List<NetworkEntry>();

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
                        if (parts.Length >= 6)
                        {
                            var entry = new NetworkEntry
                            {
                                SessionId = parts[0],
                                Port = int.Parse(parts[1]),
                                Hostname = parts[2],
                                ProcessId = int.Parse(parts[3]),
                                RegisteredAt = DateTime.Parse(parts[4]),
                                LastHeartbeat = DateTime.Parse(parts[5]),
                                Documents = new List<string>()
                            };

                            for (int i = 6; i < parts.Length; i++)
                            {
                                if (!string.IsNullOrWhiteSpace(parts[i]))
                                {
                                    entry.Documents.Add(parts[i]);
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
            while (!cancellationToken.IsCancellationRequested && isRunning)
            {
                try
                {
                    if (listener.Pending())
                    {
                        var client = await listener.AcceptTcpClientAsync();
                        var clientEndpoint = client.Client.RemoteEndPoint.ToString();

                        // Handle request in background thread
                        _ = Task.Run(async () => await HandleHttpsRequest(client, clientEndpoint, cancellationToken));
                    }
                    else
                    {
                        await Task.Delay(100, cancellationToken);
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch { }
            }
        }

        private async Task HandleHttpsRequest(TcpClient client, string clientEndpoint, CancellationToken cancellationToken)
        {
            LogToRevit($"[{DateTime.Now:HH:mm:ss}] TCP connection from {clientEndpoint}");

            try
            {
                using (client)
                {
                    var networkStream = client.GetStream();
                    var sslStream = new SslStream(networkStream, false);

                    try
                    {
                        if (!serverCertificate.HasPrivateKey)
                        {
                            LogToRevit($"[{DateTime.Now:HH:mm:ss}] ERROR: Certificate has no private key!");
                            return;
                        }

                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Starting SSL handshake...");

#if NET8_0_OR_GREATER
                        var protocols = SslProtocols.Tls12 | SslProtocols.Tls13;
#else
                        var protocols = SslProtocols.Tls12;
#endif
                        await sslStream.AuthenticateAsServerAsync(
                            serverCertificate,
                            clientCertificateRequired: false,
                            enabledSslProtocols: protocols,
                            checkCertificateRevocation: false
                        );

                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] SSL handshake successful");

                        await HandleHttpRequest(sslStream, clientEndpoint, cancellationToken);
                    }
                    catch (AuthenticationException ex)
                    {
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] SSL handshake failed: {ex.Message}");
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Hint: Client may be using HTTP instead of HTTPS");
                    }
                    catch (System.Exception ex)
                    {
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Error in SSL stream: {ex.Message}");
                    }
                    finally
                    {
                        try
                        {
                            sslStream?.Dispose();
                        }
                        catch { }
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogToRevit($"[{DateTime.Now:HH:mm:ss}] Error handling HTTPS request from {clientEndpoint}: {ex.Message}");
            }
        }

        private async Task HandleHttpRequest(Stream stream, string clientEndpoint, CancellationToken cancellationToken)
        {
            try
            {
                LogToRevit($"[{DateTime.Now:HH:mm:ss}] Connection from {clientEndpoint}");

                using (var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
                {
                    var requestLine = await reader.ReadLineAsync();
                    if (string.IsNullOrEmpty(requestLine))
                    {
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Empty request line");
                        return;
                    }

                    var parts = requestLine.Split(' ');
                    if (parts.Length < 2)
                    {
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Invalid request line: {requestLine}");
                        return;
                    }

                    var path = parts[1];
                    LogToRevit($"[{DateTime.Now:HH:mm:ss}] Request path: {path}");

                    var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    string line;
                    while ((line = await reader.ReadLineAsync()) != null && line.Length > 0)
                    {
                        var colonIndex = line.IndexOf(':');
                        if (colonIndex > 0)
                        {
                            var key = line.Substring(0, colonIndex).Trim();
                            var value = line.Substring(colonIndex + 1).Trim();
                            headers[key] = value;
                        }
                    }

                    // Read body
                    string code = "";
                    if (headers.ContainsKey("Content-Length"))
                    {
                        var contentLength = int.Parse(headers["Content-Length"]);
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Content-Length: {contentLength}");

                        var buffer = new char[contentLength];
                        var totalRead = 0;

                        while (totalRead < contentLength)
                        {
                            var bytesRead = await reader.ReadAsync(buffer, totalRead, contentLength - totalRead);
                            if (bytesRead == 0) break;
                            totalRead += bytesRead;
                        }

                        code = new string(buffer, 0, totalRead);
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Received {totalRead} bytes of code");
                    }

                    // Check authentication using shared token
                    string providedToken = null;
                    if (headers.ContainsKey("X-Auth-Token"))
                    {
                        providedToken = headers["X-Auth-Token"];
                    }
                    else if (headers.ContainsKey("Authorization"))
                    {
                        var authHeader = headers["Authorization"];
                        if (authHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
                        {
                            providedToken = authHeader.Substring(7).Trim();
                        }
                    }

                    var sharedToken = GetSharedAuthToken();
                    if (string.IsNullOrEmpty(providedToken) || string.IsNullOrEmpty(sharedToken) || providedToken != sharedToken)
                    {
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Authentication failed");
                        var errorResponse = new ScriptResponse
                        {
                            Success = false,
                            Error = "Authentication required. Provide X-Auth-Token header or Authorization: Bearer <token> header."
                        };
                        await SendHttpResponse(stream, errorResponse, 401, "Unauthorized");
                        return;
                    }

                    LogToRevit($"[{DateTime.Now:HH:mm:ss}] Authenticated request received");

                    // Route request based on path
                    if (path.StartsWith("/roslyn", StringComparison.OrdinalIgnoreCase))
                    {
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Routing to /roslyn endpoint");
                        await HandleRoslynRequest(stream, code);
                    }
                    else if (path.StartsWith("/screenshot", StringComparison.OrdinalIgnoreCase))
                    {
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Routing to /screenshot endpoint");
                        await HandleScreenshotRequest(stream);
                    }
                    else
                    {
                        LogToRevit($"[{DateTime.Now:HH:mm:ss}] Unknown endpoint: {path}");
                        await SendHttpResponse(stream, new ScriptResponse { Success = false, Error = "Unknown endpoint" }, 404, "Not Found");
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogToRevit($"[{DateTime.Now:HH:mm:ss}] Error handling request: {ex.Message}");
                LogToRevit($"[{DateTime.Now:HH:mm:ss}] Stack trace: {ex.StackTrace}");
            }
        }

        private async Task HandleRoslynRequest(Stream stream, string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                LogToRevit($"[{DateTime.Now:HH:mm:ss}] Empty request body");
                await SendHttpResponse(stream, new ScriptResponse { Success = false, Error = "Empty request body" });
                return;
            }

            LogToRevit($"[{DateTime.Now:HH:mm:ss}] Compiling script ({code.Length} chars)...");
            var compilationResult = await CompileScript(code);

            if (!compilationResult.Success)
            {
                LogToRevit($"[{DateTime.Now:HH:mm:ss}] Compilation failed");
                await SendHttpResponse(stream, compilationResult.ErrorResponse);
                return;
            }

            LogToRevit($"[{DateTime.Now:HH:mm:ss}] Script compiled successfully, executing via External Event...");
            var scriptResponse = await ExecuteScriptViaExternalEvent(compilationResult.CompiledScript);

            LogToRevit($"[{DateTime.Now:HH:mm:ss}] Execution completed - Success: {scriptResponse.Success}");
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
                var scriptOptions = ScriptOptions.Default
                    .AddReferences(typeof(ScriptGlobals).Assembly)
                    .AddReferences(typeof(Document).Assembly)
                    .AddReferences(typeof(UIDocument).Assembly)
                    .AddReferences(typeof(System.Linq.Enumerable).Assembly)
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
                result.Success = false;
                result.ErrorResponse = new ScriptResponse
                {
                    Success = false,
                    Error = ex.Message
                };
            }

            return Task.FromResult(result);
        }

        private Task<ScriptResponse> ExecuteScriptViaExternalEvent(Script<object> script)
        {
            var tcs = new TaskCompletionSource<ScriptResponse>();

            // Create an external event handler to execute the script on the Revit UI thread
            var handler = new ScriptExecutionHandler(script, tcs);
            var externalEvent = ExternalEvent.Create(handler);

            // Raise the external event to execute on the UI thread
            externalEvent.Raise();

            return tcs.Task;
        }

        private async Task HandleScreenshotRequest(Stream stream)
        {
            try
            {
                var screenshotPath = await CaptureScreenshot();
                var response = new ScriptResponse
                {
                    Success = true,
                    Output = screenshotPath
                };
                await SendHttpResponse(stream, response);
            }
            catch (System.Exception ex)
            {
                await SendHttpResponse(stream, new ScriptResponse { Success = false, Error = ex.Message });
            }
        }

        private Task<string> CaptureScreenshot()
        {
            var tcs = new TaskCompletionSource<string>();

            // Create an external event handler to capture screenshot on the Revit UI thread
            var handler = new ScreenshotHandler(tcs, sessionId);
            var externalEvent = ExternalEvent.Create(handler);

            // Raise the external event to execute on the UI thread
            externalEvent.Raise();

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
        private Script<object> script;
        private TaskCompletionSource<ScriptResponse> tcs;

        public ScriptExecutionHandler(Script<object> script, TaskCompletionSource<ScriptResponse> tcs)
        {
            this.script = script;
            this.tcs = tcs;
        }

        public void Execute(UIApplication app)
        {
            var response = new ScriptResponse();
            var outputCapture = new StringWriter();
            var originalOut = Console.Out;

            try
            {
                Console.SetOut(outputCapture);

                var uidoc = app.ActiveUIDocument;
                var doc = uidoc?.Document;

                if (doc == null)
                {
                    response.Success = false;
                    response.Error = "No active Revit document";
                    tcs.SetResult(response);
                    return;
                }

                var globals = new ScriptGlobals
                {
                    UIApp = app,
                    UIDoc = uidoc,
                    Doc = doc
                };

                // Execute the script (no Transaction - user must create one in their script if needed)
                var result = script.RunAsync(globals).Result;

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
                response.Success = false;
                response.Error = "Compilation error";
                response.Diagnostics = ex.Diagnostics.Select(d => d.ToString()).ToList();
                tcs.SetResult(response);
            }
            catch (System.Exception ex)
            {
                response.Success = false;
                response.Error = ex.Message;
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
        private TaskCompletionSource<string> tcs;
        private string sessionId;

        public ScreenshotHandler(TaskCompletionSource<string> tcs, string sessionId)
        {
            this.tcs = tcs;
            this.sessionId = sessionId;
        }

        public void Execute(UIApplication app)
        {
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

                tcs.SetResult(filepath);
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

        private void CaptureWindowToFile(IntPtr windowHandle, string filepath)
        {
            // Get window rectangle
            RECT rect;
            GetWindowRect(windowHandle, out rect);
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            // Create bitmap
            using (var bitmap = new System.Drawing.Bitmap(width, height))
            {
                using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                {
                    // Use CopyFromScreen to capture everything visible on screen
                    // This captures the actual rendered pixels including ribbon, panels, command line
                    graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new System.Drawing.Size(width, height));
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
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }

    public class NetworkEntry
    {
        public string SessionId { get; set; }
        public int Port { get; set; }
        public string Hostname { get; set; }
        public int ProcessId { get; set; }
        public List<string> Documents { get; set; }
        public DateTime RegisteredAt { get; set; }
        public DateTime LastHeartbeat { get; set; }
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

        public ScriptResponse()
        {
            Diagnostics = new List<string>();
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
        private static readonly string CertificateStorePath = PathHelper.GetRuntimeFilePath("server-cert.pfx");

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
