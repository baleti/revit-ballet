#if REVIT2011 || REVIT2012 || REVIT2013 || REVIT2014 || REVIT2015 || REVIT2016 || REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using RevitBallet.Commands;

namespace RevitBallet
{
    /// <summary>
    /// Logs view changes and maintains view history using SQLite database.
    /// </summary>
    public static class LogViewChanges
    {
        private static bool serverInitialized = false;

        /// <summary>
        /// When true, view activation events are not logged.
        /// Used to suppress logging of intermediate views during document switches.
        /// </summary>
        private static bool suppressLogging = false;

        /// <summary>
        /// Initializes view change logging by registering event handlers.
        /// </summary>
        public static void Initialize(UIControlledApplication application)
        {
            application.ViewActivated += OnViewActivated;
            application.ControlledApplication.DocumentOpened += OnDocumentOpened;
            application.ControlledApplication.DocumentClosing += OnDocumentClosing;
        }

        /// <summary>
        /// Cleans up view change logging by unregistering event handlers.
        /// </summary>
        public static void Cleanup(UIControlledApplication application)
        {
            application.ViewActivated -= OnViewActivated;
            application.ControlledApplication.DocumentOpened -= OnDocumentOpened;
            application.ControlledApplication.DocumentClosing -= OnDocumentClosing;
        }

        /// <summary>
        /// Gets the SessionId for the current Revit process (ProcessId as string).
        /// Works in both main addin and InvokeAddinCommand scenarios.
        /// </summary>
        public static string GetSessionId()
        {
            return System.Diagnostics.Process.GetCurrentProcess().Id.ToString();
        }

        /// <summary>
        /// Gets the DocumentSessionId for a specific document.
        /// DocumentSessionId is the same as SessionId (ProcessId) - uniqueness is achieved via (SessionId, DocumentTitle).
        /// </summary>
        public static string GetDocumentSessionId(Document doc)
        {
            if (doc == null)
                return null;

            return GetSessionId();
        }

        /// <summary>
        /// Gets or creates a DocumentSessionId for a specific document.
        /// DocumentSessionId is the same as SessionId (ProcessId) - uniqueness is achieved via (SessionId, DocumentTitle).
        /// </summary>
        public static string GetOrCreateDocumentSessionId(Document doc)
        {
            if (doc == null)
                return null;

            return GetSessionId();
        }

        /// <summary>
        /// Temporarily suppresses view activation logging.
        /// Use this before cross-document switches to avoid logging intermediate views.
        /// Must be paired with ResumeLogging().
        /// </summary>
        public static void SuppressLogging()
        {
            suppressLogging = true;
        }

        /// <summary>
        /// Resumes view activation logging after suppression.
        /// Must be called after SuppressLogging() to re-enable logging.
        /// </summary>
        public static void ResumeLogging()
        {
            suppressLogging = false;
        }

        private static void OnViewActivated(object sender, ViewActivatedEventArgs e)
        {
            // Set UIApplication for the server on first view activation
            if (!serverInitialized && sender is UIApplication uiApp)
            {
                RevitBalletServer.SetUIApplication(uiApp);
                serverInitialized = true;
            }

            Document doc = e.Document;

            // Check if the document is a family document
            if (doc.IsFamilyDocument)
            {
                return; // Exit if it's a family document
            }

            // Skip logging if suppressed (during cross-document switches)
            if (suppressLogging)
            {
                return;
            }

            string sessionId = GetSessionId();
            string documentSessionId = sessionId; // DocumentSessionId is same as SessionId (ProcessId)

            // Validate ViewId before logging
            ElementId viewId = e.CurrentActiveView?.Id;
            if (viewId == null || viewId == ElementId.InvalidElementId)
            {
                return; // Don't log invalid views
            }

            // Log view activation to database
            try
            {
                LogViewChangesDatabase.LogViewActivation(
                    sessionId: sessionId,
                    documentSessionId: documentSessionId,
                    documentTitle: doc.Title,
                    documentPath: doc.PathName ?? "",
                    viewId: viewId,
                    viewTitle: e.CurrentActiveView.Title,
                    viewType: e.CurrentActiveView.ViewType.ToString(),
                    timestamp: DateTime.Now
                );
            }
            catch
            {
                // Silently fail - don't interrupt Revit operations
            }
        }

        private static void OnDocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            // No action needed - DocumentSessionId is derived from ProcessId
        }

        private static void OnDocumentClosing(object sender, DocumentClosingEventArgs e)
        {
            // No action needed - no in-memory state to clean up
        }
    }
}

#endif
