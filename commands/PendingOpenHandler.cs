using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace RevitBallet.Commands
{
    // Reads a pending-open config written by OpenRvtDocumentInNewSession and opens the
    // specified .rvt file on the first Idling event after Revit starts.
    // Config file: %appdata%/revit-ballet/pending-open/{ProcessId}.toml
    internal static class PendingOpenHandler
    {
        private static readonly string PendingOpenDir =
            Path.Combine(PathHelper.RevitBalletDirectory, "pending-open");

        private static bool _handled;

        public static void Register(UIControlledApplication application)
        {
            _handled = false;
            application.Idling += OnIdling;
        }

        private static void OnIdling(object sender, IdlingEventArgs e)
        {
            if (_handled) return;

            var uiApp = sender as UIApplication;
            if (uiApp == null) return;

            string configPath = Path.Combine(PendingOpenDir,
                $"{System.Diagnostics.Process.GetCurrentProcess().Id}.toml");

            if (!File.Exists(configPath)) return;

            // Mark as handled immediately so we never retry even if an exception occurs below.
            _handled = true;
            uiApp.Idling -= OnIdling;

            try
            {
                var config = TomlHelper.Read(configPath);
                try { File.Delete(configPath); } catch { }
                ProcessPendingOpen(uiApp, config);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[PendingOpenHandler] Failed: {ex.Message}");
            }
        }

        private static void ProcessPendingOpen(UIApplication uiApp, Dictionary<string, string> config)
        {
            string filePath    = TomlHelper.GetString(config, "FilePath");
            bool isWorkshared  = TomlHelper.GetBool(config, "IsWorkshared");
            bool isCentral     = TomlHelper.GetBool(config, "IsCentral");
            bool createNewLocal = TomlHelper.GetBool(config, "CreateNewLocal", true);
            string worksetOption = TomlHelper.GetString(config, "WorksetOption", "LastViewed");
            string worksetsStr  = TomlHelper.GetString(config, "WorksetsToOpen", "");

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[PendingOpenHandler] File not found: {filePath}");
                return;
            }

            // Create local copy of central model if requested.
            if (isWorkshared && isCentral && createNewLocal)
            {
                string localPath = GenerateLocalCopyPath(filePath);
                try
                {
                    var centralModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
                    var localModelPath   = ModelPathUtils.ConvertUserVisiblePathToModelPath(localPath);
                    WorksharingUtils.CreateNewLocal(centralModelPath, localModelPath);
                    filePath = localPath;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[PendingOpenHandler] CreateNewLocal failed: {ex.Message}. Opening central directly.");
                    // Fall through and open the central file directly.
                }
            }

            var openOptions = new OpenOptions();

            if (worksetOption == "OpenAll")
            {
                openOptions.SetOpenWorksetsConfiguration(
                    new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets));
            }
            else if (worksetOption == "CloseAll")
            {
                openOptions.SetOpenWorksetsConfiguration(
                    new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));
            }
            else if (worksetOption == "Specify" && !string.IsNullOrEmpty(worksetsStr))
            {
                openOptions.SetOpenWorksetsConfiguration(
                    BuildSpecifyConfig(uiApp, filePath, isWorkshared, worksetsStr));
            }
            // "LastViewed" → don't set workset config; Revit uses the file's saved state.

            var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            uiApp.Application.OpenDocumentFile(modelPath, openOptions);
        }

        // Opens the file briefly (detached, all worksets closed) just to resolve workset
        // names → IDs, then closes it and returns the proper WorksetConfiguration.
        private static WorksetConfiguration BuildSpecifyConfig(
            UIApplication uiApp, string filePath, bool isWorkshared, string worksetsStr)
        {
            var requested = new HashSet<string>(
                worksetsStr.Split('|'), StringComparer.OrdinalIgnoreCase);

            var config = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);

            if (!isWorkshared) return config;

            Document tempDoc = null;
            try
            {
                var tempOptions = new OpenOptions();
                tempOptions.DetachFromCentralOption =
                    DetachFromCentralOption.DetachAndPreserveWorksets;
                tempOptions.SetOpenWorksetsConfiguration(
                    new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));

                var tempPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
                tempDoc = uiApp.Application.OpenDocumentFile(tempPath, tempOptions);

                var toOpen = new FilteredWorksetCollector(tempDoc)
                    .OfKind(WorksetKind.UserWorkset)
                    .ToWorksets()
                    .Where(ws => requested.Contains(ws.Name))
                    .Select(ws => ws.Id)
                    .ToList();

                if (toOpen.Count > 0)
                    config.Open(toOpen);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[PendingOpenHandler] Workset resolution failed: {ex.Message}");
            }
            finally
            {
                try { tempDoc?.Close(false); } catch { }
            }

            return config;
        }

        private static string GenerateLocalCopyPath(string centralPath)
        {
            string dir  = Path.GetDirectoryName(centralPath) ?? "";
            string name = Path.GetFileNameWithoutExtension(centralPath);
            string user = Environment.UserName;

            string candidate = Path.Combine(dir, $"{name}_{user}.rvt");
            if (!File.Exists(candidate)) return candidate;

            string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return Path.Combine(dir, $"{name}_{user}_{ts}.rvt");
        }
    }
}
