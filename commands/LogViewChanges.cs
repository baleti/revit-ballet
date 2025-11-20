#if REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI.Events;
using System.IO;
using System;
using System.Linq;
using System.Collections.Generic;
using RevitBallet.Commands;

public class LogViewChanges : IExternalApplication
{
    private string logFilePath;
    private static bool serverInitialized = false;

    public Result OnStartup(UIControlledApplication application)
    {
        // Run all startup tasks (directory initialization and update migration)
        Startup.RunStartupTasks(application);

        application.ViewActivated += OnViewActivated;
        application.ControlledApplication.DocumentOpened += OnDocumentOpened;

        // Initialize the server (it doesn't need UIApplication immediately)
        try
        {
            RevitBalletServer.InitializeServer();
        }
        catch
        {
            // Silently fail - don't interrupt Revit startup
        }

        return Result.Succeeded;
    }

    public Result OnShutdown(UIControlledApplication application)
    {
        application.ViewActivated -= OnViewActivated;
        application.ControlledApplication.DocumentOpened -= OnDocumentOpened;

        // Terminate the server
        try
        {
            RevitBalletServer.TerminateServer();
        }
        catch
        {
            // Silently fail
        }

        return Result.Succeeded;
    }

    private void OnViewActivated(object sender, ViewActivatedEventArgs e)
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

    private void OnDocumentOpened(object sender, DocumentOpenedEventArgs e)
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

        // Clear the contents of the log file
        if (File.Exists(logFilePath))
        {
            // Create a backup copy with a .last suffix
            string backupFilePath = logFilePath + ".last";
            File.Copy(logFilePath, backupFilePath, true);

            // Clear the contents of the log file
            File.WriteAllText(logFilePath, string.Empty);
        }
    }
}

#endif
