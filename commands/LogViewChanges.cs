#if REVIT2011 || REVIT2012 || REVIT2013 || REVIT2014 || REVIT2015 || REVIT2016 || REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI.Events;
using System.IO;
using System;
using System.Linq;
using System.Collections.Generic;
using RevitBallet.Commands;

namespace RevitBallet
{
    /// <summary>
    /// Logs view changes and maintains view history for each project.
    /// </summary>
    public static class LogViewChanges
    {
        private static string logFilePath;
        private static bool serverInitialized = false;

        /// <summary>
        /// Initializes view change logging by registering event handlers.
        /// </summary>
        public static void Initialize(UIControlledApplication application)
        {
            application.ViewActivated += OnViewActivated;
            application.ControlledApplication.DocumentOpened += OnDocumentOpened;
        }

        /// <summary>
        /// Cleans up view change logging by unregistering event handlers.
        /// </summary>
        public static void Cleanup(UIControlledApplication application)
        {
            application.ViewActivated -= OnViewActivated;
            application.ControlledApplication.DocumentOpened -= OnDocumentOpened;
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

            string projectName = doc != null ? doc.Title : "UnknownProject";

            // Get the log file path using PathHelper (ensures directory exists)
            logFilePath = PathHelper.GetLogViewChangesPath(projectName);

            List<string> logEntries = File.Exists(logFilePath) ? File.ReadAllLines(logFilePath).ToList() : new List<string>();

            // Add the new entry
            logEntries.Add($"{e.CurrentActiveView.Id} {e.CurrentActiveView.Title}");

            // Remove duplicates, preserving the order (keeping the first entry from the bottom)
            HashSet<string> seen = new HashSet<string>();
            int insertIndex = logEntries.Count;
            for (int i = logEntries.Count - 1; i >= 0; i--)
            {
                if (seen.Add(logEntries[i]))
                {
                    logEntries[--insertIndex] = logEntries[i];
                }
            }
            logEntries = logEntries.Skip(insertIndex).ToList();

            File.WriteAllLines(logFilePath, logEntries);
        }

        private static void OnDocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            Document doc = e.Document;
            string projectName = doc != null ? doc.Title : "UnknownProject";

            // Check if the document is a family document
            if (doc.IsFamilyDocument)
            {
                return; // Exit if it's a family document
            }

            // Get the log file path using PathHelper (ensures directory exists)
            logFilePath = PathHelper.GetLogViewChangesPath(projectName);

            // Clear the contents of the log file or create it if it doesn't exist
            if (File.Exists(logFilePath))
            {
                // Create a backup copy with a .last suffix
                string backupFilePath = logFilePath + ".last";
                File.Copy(logFilePath, backupFilePath, true);
            }

            // Clear/create the log file (WriteAllText creates the file if it doesn't exist)
            File.WriteAllText(logFilePath, string.Empty);
        }
    }
}

#endif
