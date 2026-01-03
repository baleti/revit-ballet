using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

namespace RevitBallet.Commands
{
    /// <summary>
    /// Utility to detect and diagnose dangling transactions.
    /// </summary>
    public static class TransactionMonitor
    {
        /// <summary>
        /// Checks if any document has an open transaction.
        /// Returns list of documents with open transactions and diagnostic info.
        /// </summary>
        public static List<string> CheckForOpenTransactions(UIApplication uiApp)
        {
            var issues = new List<string>();

            foreach (Document doc in uiApp.Application.Documents)
            {
                if (doc.IsLinked) continue;

                try
                {
                    // Try to start a transaction - this will fail if one is already open
                    using (var testTrans = new Transaction(doc, "TransactionMonitor_Test"))
                    {
                        TransactionStatus status = testTrans.Start();
                        testTrans.RollBack();

                        // No issue - transaction started successfully
                    }
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                {
                    issues.Add($"Document '{doc.Title}' has an open transaction or sub-transaction: {ex.Message}");
                }
                catch (Exception ex)
                {
                    issues.Add($"Document '{doc.Title}' unexpected error: {ex.Message}");
                }
            }

            return issues;
        }

        /// <summary>
        /// Checks if a specific document has an open transaction.
        /// </summary>
        public static bool HasOpenTransaction(Document doc)
        {
            if (doc == null || doc.IsLinked) return false;

            try
            {
                using (var testTrans = new Transaction(doc, "TransactionMonitor_Test"))
                {
                    testTrans.Start();
                    testTrans.RollBack();
                    return false; // No open transaction
                }
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                return true; // Transaction is open
            }
            catch
            {
                return false; // Other error, assume no transaction
            }
        }

        /// <summary>
        /// Writes transaction status to diagnostic file.
        /// Call this before critical operations like document switching.
        /// </summary>
        public static void WriteDiagnostic(UIApplication uiApp, string operationName)
        {
            var diagnosticLines = new List<string>();
            diagnosticLines.Add($"=== Transaction Monitor: {operationName} at {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ===");
            diagnosticLines.Add("");

            var issues = CheckForOpenTransactions(uiApp);

            if (issues.Count == 0)
            {
                diagnosticLines.Add("✓ No open transactions detected");
            }
            else
            {
                diagnosticLines.Add($"⚠ WARNING: {issues.Count} document(s) with open transactions:");
                diagnosticLines.Add("");
                foreach (var issue in issues)
                {
                    diagnosticLines.Add($"  • {issue}");
                }
            }

            try
            {
                string diagnosticPath = System.IO.Path.Combine(
                    PathHelper.RuntimeDirectory,
                    "diagnostics",
                    $"TransactionMonitor-{DateTime.Now:yyyyMMdd-HHmmss-fff}.txt");

                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(diagnosticPath));
                System.IO.File.WriteAllLines(diagnosticPath, diagnosticLines);
            }
            catch
            {
                // Silently fail - don't interrupt operations
            }
        }
    }
}
