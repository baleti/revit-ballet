using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using RevitBallet.Commands;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitBallet.Commands
{
    /// <summary>
    /// Handles all startup tasks for revit-ballet, including directory initialization and update migration.
    /// </summary>
    public static class Startup
    {
        /// <summary>
        /// Runs all startup tasks required for revit-ballet.
        /// This should be called from IExternalApplication.OnStartup().
        /// </summary>
        /// <param name="application">The UIControlledApplication instance</param>
        public static void RunStartupTasks(UIControlledApplication application)
        {
            // Ensure base directories exist
            PathHelper.EnsureBaseDirectoriesExist();

            // Handle update migration from revit-ballet.update to revit-ballet folder
            HandleUpdateMigration(application);
        }

        /// <summary>
        /// Handles migration of files from update folders to the main revit-ballet folder.
        /// Supports both standard (revit-ballet.update) and timestamped (revit-ballet.update.YYYYMMDD...) folders.
        /// This two-phase process ensures safe updates:
        /// 1. First startup: Running from update folder, copy files to main folder
        /// 2. Second startup: Running from main folder, verify and delete all update folders
        /// </summary>
        /// <param name="application">The UIControlledApplication instance</param>
        private static void HandleUpdateMigration(UIControlledApplication application)
        {
            try
            {
                string revitVersion = application.ControlledApplication.VersionNumber;
                string addinsPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Autodesk",
                    "Revit",
                    "Addins",
                    revitVersion
                );

                string mainFolder = Path.Combine(addinsPath, "revit-ballet");
                string addinPath = Path.Combine(addinsPath, "revit-ballet.addin");

                // Find all update folders (revit-ballet.update and revit-ballet.update.*)
                var updateFolders = Directory.Exists(addinsPath)
                    ? Directory.GetDirectories(addinsPath, "revit-ballet.update*")
                    : new string[0];

                if (updateFolders.Length == 0)
                {
                    return; // No update pending
                }

                // Check if we're currently running from any update folder
                string currentAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                bool runningFromUpdate = currentAssemblyPath.Contains("revit-ballet.update");

                // Find which update folder we're running from (if any)
                // Sort by length descending to match the most specific (longest) folder first,
                // preventing "revit-ballet.update" from matching "revit-ballet.update.20240115..."
                string currentUpdateFolder = runningFromUpdate
                    ? updateFolders
                        .OrderByDescending(f => f.Length)
                        .FirstOrDefault(f => currentAssemblyPath.StartsWith(f + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                    : null;

                if (runningFromUpdate && currentUpdateFolder != null)
                {
                    // Phase 1: We're running from an update folder - perform migration
                    PerformUpdateMigration(mainFolder, currentUpdateFolder, addinPath);

                    // Clean up other update folders we're not running from
                    foreach (var folder in updateFolders.Where(f => !f.Equals(currentUpdateFolder, StringComparison.OrdinalIgnoreCase)))
                    {
                        TryDeleteUpdateFolder(folder);
                    }
                }
                else
                {
                    // Phase 2: We're running from main folder - cleanup all update folders
                    CleanupAllUpdateFolders(mainFolder, updateFolders, addinPath);
                }
            }
            catch
            {
                // Silently fail - don't interrupt Revit startup
            }
        }

        /// <summary>
        /// Attempts to delete an update folder. Fails silently if folder is in use.
        /// </summary>
        private static void TryDeleteUpdateFolder(string folderPath)
        {
            try
            {
                if (Directory.Exists(folderPath))
                {
                    Directory.Delete(folderPath, recursive: true);
                }
            }
            catch
            {
                // Silently fail - folder may be in use by another Revit version
            }
        }

        /// <summary>
        /// Phase 1: Copies files from update folder to main folder and updates .addin file.
        /// Only flips the manifest back to main folder if the key DLL was actually copied —
        /// if the copy fails (DLL locked by another Revit instance), we leave the manifest
        /// pointing to the update folder so the next clean restart retries automatically.
        /// </summary>
        private static void PerformUpdateMigration(string mainFolder, string updateFolder, string addinPath)
        {
            if (!Directory.Exists(mainFolder))
                Directory.CreateDirectory(mainFolder);

            // Copy all files from update folder to main folder.
            // Track whether the key DLL copied successfully so we can decide whether to
            // flip the manifest.  Non-DLL files (native interop etc.) are best-effort.
            bool keyDllCopied = false;
            foreach (string srcFile in Directory.GetFiles(updateFolder))
            {
                string fileName = Path.GetFileName(srcFile);
                string destFile = Path.Combine(mainFolder, fileName);
                try
                {
                    File.Copy(srcFile, destFile, overwrite: true);
                    if (fileName.Equals("revit-ballet.dll", StringComparison.OrdinalIgnoreCase))
                        keyDllCopied = true;
                }
                catch
                {
                    // File is still locked — leave it for the next restart.
                }
            }

            // Only redirect the manifest to the main folder once the key DLL is there.
            // If the copy failed, keep the manifest pointing to the update folder so that
            // the next Revit startup (after all instances are closed) retries Phase 1.
            if (!keyDllCopied) return;

            if (File.Exists(addinPath))
            {
                try
                {
                    var doc = System.Xml.Linq.XDocument.Load(addinPath);
                    foreach (var assemblyElement in doc.Descendants("Assembly"))
                    {
                        string currentValue = assemblyElement.Value;
                        if (currentValue.Contains("revit-ballet.update"))
                        {
                            string updateFolderName = Path.GetFileName(updateFolder);
                            assemblyElement.Value = currentValue.Replace(updateFolderName, "revit-ballet");
                        }
                    }
                    doc.Save(addinPath);
                }
                catch
                {
                    // Silently fail — worst case the manifest still points to the update folder,
                    // which works fine and will be corrected on next startup.
                }
            }
        }

        /// <summary>
        /// Phase 2: Verifies migration was successful and deletes all update folders.
        /// </summary>
        private static void CleanupAllUpdateFolders(string mainFolder, string[] updateFolders, string addinPath)
        {
            // First check if .addin points to main folder (not any update folder)
            bool addinPointsToMain = true;
            if (File.Exists(addinPath))
            {
                try
                {
                    var doc = System.Xml.Linq.XDocument.Load(addinPath);
                    foreach (var assemblyElement in doc.Descendants("Assembly"))
                    {
                        string currentValue = assemblyElement.Value;
                        if (currentValue.Contains("revit-ballet.update"))
                        {
                            addinPointsToMain = false;
                            break;
                        }
                    }
                }
                catch
                {
                    addinPointsToMain = false;
                }
            }

            if (!addinPointsToMain)
            {
                // .addin still points to an update folder - don't delete anything
                return;
            }

            // Delete all update folders
            foreach (var updateFolder in updateFolders)
            {
                bool canDelete = true;

                // Check if files from this update folder exist in main folder
                if (Directory.Exists(mainFolder) && Directory.Exists(updateFolder))
                {
                    foreach (string updateFile in Directory.GetFiles(updateFolder))
                    {
                        string fileName = Path.GetFileName(updateFile);
                        string mainFile = Path.Combine(mainFolder, fileName);

                        if (!File.Exists(mainFile))
                        {
                            canDelete = false;
                            break;
                        }

                        // Compare file sizes as a quick check (skip byte-by-byte for performance)
                        try
                        {
                            var updateInfo = new FileInfo(updateFile);
                            var mainInfo = new FileInfo(mainFile);

                            if (updateInfo.Length != mainInfo.Length)
                            {
                                canDelete = false;
                                break;
                            }
                        }
                        catch
                        {
                            canDelete = false;
                            break;
                        }
                    }
                }

                // Delete update folder if safe to do so
                if (canDelete)
                {
                    TryDeleteUpdateFolder(updateFolder);
                }
            }
        }
    }
}
