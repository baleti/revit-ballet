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

            // Handle reopening: restore from .last backup if it exists
            if (File.Exists(logFilePath))
            {
                // Create a backup copy with a .last suffix
                string backupFilePath = logFilePath + ".last";
                File.Copy(logFilePath, backupFilePath, true);

                // Update element IDs in case they changed (detachment, copying)
                UpdateElementIdsInLog(doc, logFilePath);
            }
            else
            {
                // New document - try to restore from .last backup
                string backupFilePath = logFilePath + ".last";
                if (File.Exists(backupFilePath))
                {
                    File.Copy(backupFilePath, logFilePath, false);
                    UpdateElementIdsInLog(doc, logFilePath);
                }
                else
                {
                    // Create new empty log
                    File.WriteAllText(logFilePath, string.Empty);
                }
            }
        }

        /// <summary>
        /// Updates element IDs in the log file to match current document IDs.
        /// This handles cases where IDs change due to detachment or copying.
        /// </summary>
        private static void UpdateElementIdsInLog(Document doc, string logFilePath)
        {
            if (!File.Exists(logFilePath))
                return;

            var logEntries = File.ReadAllLines(logFilePath).ToList();
            bool modified = false;

            // Build a lookup of all views by title
            var viewsByTitle = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .GroupBy(v => v.Title)
                .ToDictionary(g => g.Key, g => g.First());

            // Update each entry with current element ID
            for (int i = 0; i < logEntries.Count; i++)
            {
                string entry = logEntries[i].Trim();
                if (string.IsNullOrEmpty(entry))
                    continue;

                // Parse: "ElementId Title"
                var parts = entry.Split(new[] { ' ' }, 2);
                if (parts.Length != 2)
                    continue;

                string title = parts[1];

                // Find view by title
                if (viewsByTitle.TryGetValue(title, out View view))
                {
                    // Update with current element ID
                    string newEntry = $"{view.Id} {title}";
                    if (newEntry != entry)
                    {
                        logEntries[i] = newEntry;
                        modified = true;
                    }
                }
                // If view not found by title, keep the old entry (might have been deleted/renamed)
            }

            // Write back if modified
            if (modified)
            {
                File.WriteAllLines(logFilePath, logEntries);
            }
        }
    }
}

#endif
