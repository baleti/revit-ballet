using System;
using System.IO;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitBallet.Commands
{
    /// <summary>
    /// Centralized helper for managing application data paths and ensuring required directories exist.
    /// </summary>
    public static class PathHelper
    {
        // Base paths
        private static readonly string AppDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        private static readonly string RevitBalletBase = Path.Combine(AppDataPath, "revit-ballet");
        private static readonly string RuntimeBase = Path.Combine(RevitBalletBase, "runtime");

        /// <summary>
        /// Gets the base revit-ballet directory path in AppData.
        /// </summary>
        public static string RevitBalletDirectory => RevitBalletBase;

        /// <summary>
        /// Gets the runtime directory path in AppData.
        /// </summary>
        public static string RuntimeDirectory => RuntimeBase;

        /// <summary>
        /// Ensures that the base revit-ballet and runtime directories exist.
        /// This should be called during application startup.
        /// </summary>
        public static void EnsureBaseDirectoriesExist()
        {
            EnsureDirectoryExists(RevitBalletBase);
            EnsureDirectoryExists(RuntimeBase);
        }

        /// <summary>
        /// Ensures that a specific subdirectory exists within the runtime directory.
        /// </summary>
        /// <param name="subdirectory">The subdirectory name or path relative to runtime directory.</param>
        /// <returns>The full path to the subdirectory.</returns>
        public static string EnsureRuntimeSubdirectoryExists(string subdirectory)
        {
            string fullPath = Path.Combine(RuntimeBase, subdirectory);
            EnsureDirectoryExists(fullPath);
            return fullPath;
        }

        /// <summary>
        /// Gets the path to a file in the runtime directory, ensuring the directory exists.
        /// </summary>
        /// <param name="fileName">The file name.</param>
        /// <returns>The full path to the file.</returns>
        public static string GetRuntimeFilePath(string fileName)
        {
            EnsureDirectoryExists(RuntimeBase);
            return Path.Combine(RuntimeBase, fileName);
        }

        /// <summary>
        /// Gets the path to a file in a runtime subdirectory, ensuring the directory exists.
        /// </summary>
        /// <param name="subdirectory">The subdirectory name or path relative to runtime directory.</param>
        /// <param name="fileName">The file name.</param>
        /// <returns>The full path to the file.</returns>
        public static string GetRuntimeSubdirectoryFilePath(string subdirectory, string fileName)
        {
            string directoryPath = EnsureRuntimeSubdirectoryExists(subdirectory);
            return Path.Combine(directoryPath, fileName);
        }

        /// <summary>
        /// Ensures a directory exists, creating it if necessary.
        /// </summary>
        /// <param name="path">The directory path.</param>
        public static void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        /// <summary>
        /// Gets the LogViewChanges directory path for a specific project.
        /// </summary>
        /// <param name="projectName">The project name.</param>
        /// <returns>The full path to the project's log directory.</returns>
        public static string GetLogViewChangesPath(string projectName)
        {
            return GetRuntimeSubdirectoryFilePath("LogViewChanges", projectName);
        }
    }
}
