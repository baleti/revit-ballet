using Autodesk.Revit.UI;
using System;
using System.IO;
using RevitBallet.Commands;

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
        /// Handles migration of files from the revit-ballet.update folder to the revit-ballet folder.
        /// This two-phase process ensures safe updates:
        /// 1. First startup: Running from update folder, copy files to main folder
        /// 2. Second startup: Running from main folder, verify and delete update folder
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
                string updateFolder = Path.Combine(addinsPath, "revit-ballet.update");
                string addinPath = Path.Combine(addinsPath, "revit-ballet.addin");

                // Check if update folder exists
                if (!Directory.Exists(updateFolder))
                {
                    return; // No update pending
                }

                // Check if we're currently running from the update folder
                string currentAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                bool runningFromUpdate = currentAssemblyPath.Contains("revit-ballet.update");

                if (runningFromUpdate)
                {
                    // Phase 1: We're running from update folder - perform migration
                    PerformUpdateMigration(mainFolder, updateFolder, addinPath);
                }
                else
                {
                    // Phase 2: We're running from main folder - verify and cleanup
                    CleanupUpdateFolder(mainFolder, updateFolder, addinPath);
                }
            }
            catch
            {
                // Silently fail - don't interrupt Revit startup
            }
        }

        /// <summary>
        /// Phase 1: Copies files from update folder to main folder and updates .addin file.
        /// </summary>
        private static void PerformUpdateMigration(string mainFolder, string updateFolder, string addinPath)
        {
            // Delete old files from main folder
            if (Directory.Exists(mainFolder))
            {
                foreach (string file in Directory.GetFiles(mainFolder, "*.dll"))
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                        // Skip if can't delete
                    }
                }
            }
            else
            {
                Directory.CreateDirectory(mainFolder);
            }

            // Copy all files from update folder to main folder
            foreach (string file in Directory.GetFiles(updateFolder))
            {
                string fileName = Path.GetFileName(file);
                string destFile = Path.Combine(mainFolder, fileName);
                try
                {
                    File.Copy(file, destFile, overwrite: true);
                }
                catch
                {
                    // Skip if can't copy
                }
            }

            // Update .addin file to point to main folder
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
                            assemblyElement.Value = currentValue.Replace("revit-ballet.update", "revit-ballet");
                        }
                    }
                    doc.Save(addinPath);
                }
                catch
                {
                    // Skip if can't update addin file
                }
            }
        }

        /// <summary>
        /// Phase 2: Verifies migration was successful and deletes update folder.
        /// </summary>
        private static void CleanupUpdateFolder(string mainFolder, string updateFolder, string addinPath)
        {
            bool canDelete = true;

            // Check if .addin points to main folder
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
                            canDelete = false;
                            break;
                        }
                    }
                }
                catch
                {
                    canDelete = false;
                }
            }

            // Check if files in both folders are the same
            if (canDelete && Directory.Exists(mainFolder))
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

                    try
                    {
                        byte[] updateBytes = File.ReadAllBytes(updateFile);
                        byte[] mainBytes = File.ReadAllBytes(mainFile);

                        if (updateBytes.Length != mainBytes.Length)
                        {
                            canDelete = false;
                            break;
                        }

                        for (int i = 0; i < updateBytes.Length; i++)
                        {
                            if (updateBytes[i] != mainBytes[i])
                            {
                                canDelete = false;
                                break;
                            }
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
                try
                {
                    Directory.Delete(updateFolder, recursive: true);
                }
                catch
                {
                    // Failed to delete - will try again next time
                }
            }
        }
    }
}
