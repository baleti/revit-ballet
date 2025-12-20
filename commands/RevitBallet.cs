#if REVIT2011 || REVIT2012 || REVIT2013 || REVIT2014 || REVIT2015 || REVIT2016 || REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
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
                // Log error but don't interrupt Revit startup
                System.Diagnostics.Debug.WriteLine($"Failed to initialize view history database: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }

            // Initialize view logging
            LogViewChanges.Initialize(application);

            // Initialize the server
            try
            {
                RevitBalletServer.InitializeServer();
            }
            catch
            {
                // Silently fail - don't interrupt Revit startup
            }

            // Initialize DataGrid column handler registry for automatic editing
            try
            {
                CustomGUIs.ColumnHandlerRegistry.RegisterStandardHandlers();
            }
            catch
            {
                // Silently fail - don't interrupt Revit startup
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Cleanup view logging
            LogViewChanges.Cleanup(application);

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
    }
}

#endif
