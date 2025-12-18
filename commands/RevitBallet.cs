#if REVIT2011 || REVIT2012 || REVIT2013 || REVIT2014 || REVIT2015 || REVIT2016 || REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
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
        public Result OnStartup(UIControlledApplication application)
        {
            // Run all startup tasks (directory initialization and update migration)
            Startup.RunStartupTasks(application);

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
