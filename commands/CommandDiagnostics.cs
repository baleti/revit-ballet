using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace RevitBallet.Commands
{
    /// <summary>
    /// Centralized diagnostics for tracking command execution and detecting issues.
    /// </summary>
    public static class CommandDiagnostics
    {
        private static Stack<string> callStack = new Stack<string>();

        /// <summary>
        /// Call this at the START of a command to track execution and check for issues.
        /// Returns a diagnostic session that should be disposed when command completes.
        /// </summary>
        public static DiagnosticSession StartCommand(string commandName, UIApplication uiApp)
        {
            return new DiagnosticSession(commandName, uiApp);
        }

        /// <summary>
        /// Diagnostic session that tracks command execution from start to finish.
        /// Use with 'using' statement to ensure proper cleanup.
        /// </summary>
        public class DiagnosticSession : IDisposable
        {
            private readonly string commandName;
            private readonly UIApplication uiApp;
            private readonly Stopwatch stopwatch;
            private readonly List<string> diagnosticLines;
            private readonly DateTime startTime;

            public DiagnosticSession(string commandName, UIApplication uiApp)
            {
                this.commandName = commandName;
                this.uiApp = uiApp;
                this.stopwatch = Stopwatch.StartNew();
                this.startTime = DateTime.Now;
                this.diagnosticLines = new List<string>();

                // Track call stack
                callStack.Push(commandName);

                // Start diagnostic log
                diagnosticLines.Add($"=== COMMAND START: {commandName} at {startTime:yyyy-MM-dd HH:mm:ss.fff} ===");
                diagnosticLines.Add($"Call Stack Depth: {callStack.Count}");
                if (callStack.Count > 1)
                {
                    diagnosticLines.Add($"Called From: {string.Join(" → ", callStack.ToArray())}");
                }
                diagnosticLines.Add("");

                // Check for open transactions
                var transactionIssues = TransactionMonitor.CheckForOpenTransactions(uiApp);
                if (transactionIssues.Count > 0)
                {
                    diagnosticLines.Add("⚠ WARNING: Open transactions detected at START:");
                    foreach (var issue in transactionIssues)
                    {
                        diagnosticLines.Add($"  • {issue}");
                    }
                    diagnosticLines.Add("");
                }
                else
                {
                    diagnosticLines.Add("✓ No open transactions at START");
                    diagnosticLines.Add("");
                }

                // Log active document
                var activeDoc = uiApp.ActiveUIDocument?.Document;
                if (activeDoc != null)
                {
                    diagnosticLines.Add($"Active Document: {activeDoc.Title}");
                    diagnosticLines.Add($"Active View: {uiApp.ActiveUIDocument.ActiveView?.Name ?? "None"}");
                    diagnosticLines.Add("");
                }
            }

            /// <summary>
            /// Add a log entry to the diagnostic session.
            /// </summary>
            public void Log(string message)
            {
                diagnosticLines.Add($"[{stopwatch.ElapsedMilliseconds}ms] {message}");
            }

            /// <summary>
            /// Log an error to the diagnostic session.
            /// </summary>
            public void LogError(string error, Exception ex = null)
            {
                diagnosticLines.Add($"[{stopwatch.ElapsedMilliseconds}ms] ❌ ERROR: {error}");
                if (ex != null)
                {
                    diagnosticLines.Add($"  Exception: {ex.GetType().Name}");
                    diagnosticLines.Add($"  Message: {ex.Message}");
                    diagnosticLines.Add($"  StackTrace: {ex.StackTrace}");
                }
            }

            /// <summary>
            /// Called when command completes (via Dispose).
            /// </summary>
            public void Dispose()
            {
                stopwatch.Stop();

                // Pop from call stack
                if (callStack.Count > 0 && callStack.Peek() == commandName)
                {
                    callStack.Pop();
                }

                // Check for open transactions at END
                diagnosticLines.Add("");
                diagnosticLines.Add($"=== COMMAND END: {commandName} ===");
                diagnosticLines.Add($"Duration: {stopwatch.ElapsedMilliseconds}ms");
                diagnosticLines.Add("");

                var transactionIssues = TransactionMonitor.CheckForOpenTransactions(uiApp);
                if (transactionIssues.Count > 0)
                {
                    diagnosticLines.Add("⚠⚠⚠ CRITICAL: Open transactions detected at END:");
                    foreach (var issue in transactionIssues)
                    {
                        diagnosticLines.Add($"  • {issue}");
                    }
                    diagnosticLines.Add("");
                    diagnosticLines.Add("THIS WILL CAUSE REVIT TO ROLLBACK CHANGES!");
                }
                else
                {
                    diagnosticLines.Add("✓ No open transactions at END (clean exit)");
                }

                // Write diagnostic file
                WriteDiagnostic();
            }

            private void WriteDiagnostic()
            {
                // Diagnostic file writing disabled
                // try
                // {
                //     string diagnosticPath = System.IO.Path.Combine(
                //         PathHelper.RuntimeDirectory,
                //         "diagnostics",
                //         $"{commandName}-{startTime:yyyyMMdd-HHmmss-fff}.txt");
                //
                //     System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
                //     System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);
                // }
                // catch
                // {
                //     // Silently fail - don't interrupt command execution
                // }
            }
        }
    }
}
