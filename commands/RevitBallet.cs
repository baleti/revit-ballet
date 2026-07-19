using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.UI;
using RevitBallet.Commands;

namespace RevitBallet
{
    /// <summary>
    /// Main application entry point for revit-ballet.
    /// Handles startup and shutdown tasks for the Revit add-in.
    /// </summary>
    public class RevitBallet : IExternalApplication
    {
        /// <summary>
        /// Session ID for this Revit process instance (ProcessId as string).
        /// Shared across all documents opened in this Revit session.
        /// </summary>
        public static string SessionId => System.Diagnostics.Process.GetCurrentProcess().Id.ToString();

        public Result OnStartup(UIControlledApplication application)
        {

            // Run all startup tasks (directory initialization and update migration)
            Startup.RunStartupTasks(application);

            // Initialize SQLite database for view history
            try
            {
                LogViewChangesDatabase.InitializeDatabase();
            }
            catch (Exception ex)
            {
                Log.Warn("Startup.ViewHistoryDatabase", ex);
            }

            // Initialize view logging
            LogViewChanges.Initialize(application);

            // Initialize the server
            try
            {
                RevitBalletServer.InitializeServer();
            }
            catch (Exception ex)
            {
                Log.Warn("Startup.Server", ex);
            }

            // Initialize DataGrid column handler registry for automatic editing
            try
            {
                CustomGUIs.ColumnHandlerRegistry.RegisterStandardHandlers();
            }
            catch (Exception ex)
            {
                Log.Warn("Startup.ColumnHandlerRegistry", ex);
            }

            // Subscribe to sync event to track last synchronization time
            try
            {
                application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronized;
            }
            catch (Exception ex)
            {
                Log.Warn("Startup.SubscribeDocumentSynchronized", ex);
            }

            // Subscribe to DocumentChanged event to track last transaction time
            try
            {
                application.ControlledApplication.DocumentChanged += OnDocumentChanged;
            }
            catch (Exception ex)
            {
                Log.Warn("Startup.SubscribeDocumentChanged", ex);
            }

            return Result.Succeeded;
        }

        private static void OnDocumentSynchronized(object sender, Autodesk.Revit.DB.Events.DocumentSynchronizedWithCentralEventArgs e)
        {
            try
            {
                // Pass document path (or title if path is empty) to identify which document was synced
                var doc = e.Document;
                var docIdentifier = !string.IsNullOrEmpty(doc.PathName) ? doc.PathName : doc.Title;
                RevitBalletServer.UpdateLastSyncTime(docIdentifier);
            }
            catch (Exception ex)
            {
                Log.Warn("OnDocumentSynchronized", ex);
            }
        }

        private static void OnDocumentChanged(object sender, Autodesk.Revit.DB.Events.DocumentChangedEventArgs e)
        {
            try
            {
                // Only track committed transactions (not rollbacks, undo, or redo)
                if (e.Operation == Autodesk.Revit.DB.Events.UndoOperation.TransactionCommitted)
                {
                    var doc = e.GetDocument();
                    var docIdentifier = !string.IsNullOrEmpty(doc.PathName) ? doc.PathName : doc.Title;
                    RevitBalletServer.UpdateLastTransactionTime(docIdentifier);
                }
            }
            catch (Exception ex)
            {
                Log.Warn("OnDocumentChanged", ex);
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Cleanup view logging
            LogViewChanges.Cleanup(application);

            // Unsubscribe from sync event
            try
            {
                application.ControlledApplication.DocumentSynchronizedWithCentral -= OnDocumentSynchronized;
            }
            catch (Exception ex)
            {
                Log.Warn("Shutdown.UnsubscribeDocumentSynchronized", ex);
            }

            // Unsubscribe from DocumentChanged event
            try
            {
                application.ControlledApplication.DocumentChanged -= OnDocumentChanged;
            }
            catch (Exception ex)
            {
                Log.Warn("Shutdown.UnsubscribeDocumentChanged", ex);
            }

            // Terminate the server
            try
            {
                RevitBalletServer.TerminateServer();
            }
            catch (Exception ex)
            {
                Log.Warn("Shutdown.Server", ex);
            }

            return Result.Succeeded;
        }
    }
}

