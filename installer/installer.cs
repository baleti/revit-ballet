using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;

namespace RevitBalletInstaller
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string exeName = Path.GetFileName(Assembly.GetExecutingAssembly().Location).ToLower();
            bool isUninstaller = exeName.Contains("uninstall") || (args.Length > 0 && args[0] == "/uninstall");

            if (isUninstaller)
            {
                new BalletUninstaller().Run();
            }
            else
            {
                new BalletInstaller().Run();
            }
        }
    }

    public class BalletInstaller
    {
        private const string PRODUCT_NAME = "Revit Ballet";
        private const string UNINSTALL_KEY = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\RevitBallet";

        private readonly List<RevitInstallation> installations = new List<RevitInstallation>();
        private readonly List<InstallationResult> installationResults = new List<InstallationResult>();
        private Dictionary<string, FileMapping> fileMappings = null;

        private enum InstallState
        {
            FreshInstall,
            UpdatedSuccessfully,
            UpdatedNeedsRestart
        }

        private class FileMapping
        {
            public string ResourceName { get; set; }
            public string TargetFileName { get; set; }
            public List<string> Years { get; set; }
        }

        public void Run()
        {
            try
            {
                bool alreadyInstalled = IsInstalled();

                // Load file mapping from embedded resources
                LoadFileMapping();

                DetectRevitInstallations();

                if (installations.Count == 0)
                {
                    MessageBox.Show("No Revit installations found.",
                        "Revit Ballet", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // If already installed, just update - no dialog needed

                string targetDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "revit-ballet");
                Directory.CreateDirectory(targetDir);

                // Extract .addin file and dll to each Revit version's Addins folder
                InstallToRevitVersions(targetDir);

                // Also copy DLLs to runtime location for InvokeAddinCommand
                CopyDllsToRuntimeLocation(targetDir);

                CopyInstallerForUninstall(targetDir);

                // Check for missing KeyboardShortcuts.xml files and show warning
                ShowMissingKeyboardShortcutsWarning();

                // Modify keyboard shortcuts
                var modifiedInstallations2 = new List<RevitInstallation>();
                foreach (var installation in installations)
                {
                    bool modified = ModifyKeyboardShortcuts(installation);
                    if (modified)
                    {
                        modifiedInstallations2.Add(installation);
                    }
                }

                RegisterUninstaller(targetDir);

                // Register addins as trusted to avoid "Load Always/Load Once" prompts
                RegisterTrustedAddins();

                ShowSuccessDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Installation error: {ex.Message}\n\n{ex.StackTrace}",
                    "Revit Ballet Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowSuccessDialog()
        {
            var form = new Form
            {
                Text = "Revit Ballet - Installation Complete",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.White
            };

            int currentY = 30;

            var titleLabel = new Label
            {
                Text = "Revit Ballet",
                Font = new Font("Segoe UI", 24, FontStyle.Bold),
                ForeColor = Color.FromArgb(50, 50, 50),
                Location = new Point(0, currentY),
                Size = new Size(600, 40),
                TextAlign = ContentAlignment.MiddleCenter
            };
            form.Controls.Add(titleLabel);
            currentY += 60;

            // Group results by state
            var freshInstalls = installationResults.Where(r => r.State == InstallState.FreshInstall).ToList();
            var updatedSuccessfully = installationResults.Where(r => r.State == InstallState.UpdatedSuccessfully).ToList();
            var needsRestart = installationResults.Where(r => r.State == InstallState.UpdatedNeedsRestart).ToList();

            string messageText = "";

            if (freshInstalls.Count > 0)
            {
                messageText += "Successfully installed to:\n";
                foreach (var result in freshInstalls)
                {
                    messageText += $"  • Revit {result.Year}\n";
                }
                messageText += "\n";
            }

            if (updatedSuccessfully.Count > 0)
            {
                messageText += "Successfully updated:\n";
                foreach (var result in updatedSuccessfully)
                {
                    messageText += $"  • Revit {result.Year}\n";
                }
                messageText += "\n";
            }

            if (needsRestart.Count > 0)
            {
                messageText += "Updated (restart Revit to use new version):\n";
                foreach (var result in needsRestart)
                {
                    messageText += $"  • Revit {result.Year}\n";
                }
                messageText += "\nNote: These versions had Revit running during installation.\n";
                messageText += "New commands are available immediately via InvokeAddinCommand,\n";
                messageText += "but you need to restart Revit to load the updated addin.\n";
            }

            if (installationResults.Count == 0)
            {
                messageText = "Installation completed, but no Revit versions were updated.\nThis might indicate missing DLL resources.";
            }

            using (var g = form.CreateGraphics())
            {
                var textSize = g.MeasureString(messageText, new Font("Segoe UI", 10), 520);
                var textHeight = (int)Math.Ceiling(textSize.Height);

                var messageLabel = new Label
                {
                    Text = messageText,
                    Font = new Font("Segoe UI", 10),
                    ForeColor = Color.FromArgb(60, 60, 60),
                    Location = new Point(40, currentY),
                    Size = new Size(520, textHeight + 20),
                    TextAlign = ContentAlignment.TopLeft
                };
                form.Controls.Add(messageLabel);

                currentY += textHeight + 40;
            }

            var btnOK = new Button
            {
                Text = "OK",
                Font = new Font("Segoe UI", 10),
                Size = new Size(100, 30),
                Location = new Point(250, currentY),
                FlatStyle = FlatStyle.System,
                DialogResult = DialogResult.OK
            };
            form.Controls.Add(btnOK);
            form.AcceptButton = btnOK;

            currentY += 60;
            form.ClientSize = new Size(600, currentY);

            form.ShowDialog();
        }

        private bool IsInstalled()
        {
            try
            {
                return Registry.CurrentUser.OpenSubKey(UNINSTALL_KEY) != null;
            }
            catch
            {
                return false;
            }
        }

        private void DetectRevitInstallations()
        {
            // Detection strategy:
            // - Check ProgramData to see which Revit versions are installed
            // - But install addins to AppData (current user) to avoid requiring admin privileges
            // - Revit loads addins from both ProgramData and AppData locations

            var detectionPaths = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), // C:\ProgramData
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)        // C:\Users\<user>\AppData\Roaming
            };

            var detectedYears = new HashSet<string>();

            // Scan both locations to detect which Revit versions are installed
            foreach (string basePath in detectionPaths)
            {
                string revitPath = Path.Combine(basePath, "Autodesk", "Revit", "Addins");

                if (!Directory.Exists(revitPath))
                    continue;

                // Look for year folders (2011-2030)
                string[] yearDirs = Directory.GetDirectories(revitPath);

                foreach (string yearDir in yearDirs)
                {
                    string year = Path.GetFileName(yearDir);
                    if (int.TryParse(year, out int yearNum) && yearNum >= 2011 && yearNum <= 2030)
                    {
                        detectedYears.Add(year);
                    }
                }
            }

            // Create installation entries - ALWAYS use AppData for installation (no admin required)
            string userAppData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            foreach (string year in detectedYears)
            {
                string userAddinsPath = Path.Combine(userAppData, "Autodesk", "Revit", "Addins", year);
                string revitDataPath = Path.Combine(userAppData, "Autodesk", "Revit", $"Autodesk Revit {year}");

                var installation = new RevitInstallation
                {
                    Year = year,
                    AddinsPath = userAddinsPath, // Always use user's AppData (no admin needed)
                    RevitDataPath = revitDataPath
                };

                installations.Add(installation);
            }
        }

        private void InstallToRevitVersions(string sourceDir)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var addinResource = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("revit-ballet.addin"));

            if (addinResource == null)
            {
                throw new Exception("Required resource revit-ballet.addin not found in installer.");
            }

            foreach (var installation in installations)
            {
                // Find version-specific DLL resource (e.g., installer._2024.revit-ballet.dll)
                string dllResourcePattern = $"_{installation.Year}.revit-ballet.dll";
                var dllResource = assembly.GetManifestResourceNames()
                    .FirstOrDefault(r => r.Contains(dllResourcePattern));

                if (dllResource == null)
                {
                    // Try without year prefix (fallback)
                    dllResource = assembly.GetManifestResourceNames()
                        .FirstOrDefault(r => r.EndsWith("revit-ballet.dll") && !r.Contains("._20"));
                }

                if (dllResource == null)
                {
                    // Silently skip versions without DLLs
                    continue;
                }

                string mainFolder = Path.Combine(installation.AddinsPath, "revit-ballet");
                string updateFolder = Path.Combine(installation.AddinsPath, "revit-ballet.update");
                string addinPath = Path.Combine(installation.AddinsPath, "revit-ballet.addin");

                Directory.CreateDirectory(mainFolder);

                bool wasInstalled = File.Exists(Path.Combine(mainFolder, "revit-ballet.dll"));
                bool filesAreLocked = false;

                // Try to copy all files to main folder
                try
                {
                    // Extract main DLL
                    string mainDllPath = Path.Combine(mainFolder, "revit-ballet.dll");
                    ExtractResourceSafe(dllResource, mainDllPath);

                    // Extract dependencies
                    ExtractDependenciesSafe(installation.Year, mainFolder);

                    // Copy KeyboardShortcuts.xml template if provided
                    var kbResource = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("KeyboardShortcuts.xml"));
                    if (kbResource != null)
                    {
                        string kbTemplatePath = Path.Combine(mainFolder, "KeyboardShortcuts-template.xml");
                        try
                        {
                            ExtractResourceSafe(kbResource, kbTemplatePath);
                        }
                        catch
                        {
                            // Non-critical
                        }
                    }

                    // Successfully updated main folder - no need to create .update folder
                    // But we still need to create/update .addin file if it's a fresh install
                    if (!File.Exists(addinPath))
                    {
                        string addinContent = GetResourceText(addinResource);
                        var doc = XDocument.Parse(addinContent);

                        foreach (var assemblyElement in doc.Descendants("Assembly"))
                        {
                            string currentValue = assemblyElement.Value;
                            if (!currentValue.Contains("\\") && !currentValue.Contains("/"))
                            {
                                assemblyElement.Value = Path.Combine(mainFolder, currentValue);
                            }
                        }

                        doc.Save(addinPath);
                    }

                    installationResults.Add(new InstallationResult
                    {
                        Year = installation.Year,
                        State = wasInstalled ? InstallState.UpdatedSuccessfully : InstallState.FreshInstall
                    });
                }
                catch (IOException)
                {
                    // Files are locked (Revit is running) - use update folder strategy
                    filesAreLocked = true;
                    Directory.CreateDirectory(updateFolder);

                    // Extract to update folder
                    string updateDllPath = Path.Combine(updateFolder, "revit-ballet.dll");
                    ExtractResource(dllResource, updateDllPath);
                    ExtractDependencies(installation.Year, updateFolder);

                    // Copy KeyboardShortcuts.xml template to update folder
                    var kbResource = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("KeyboardShortcuts.xml"));
                    if (kbResource != null)
                    {
                        string kbTemplatePath = Path.Combine(updateFolder, "KeyboardShortcuts-template.xml");
                        try
                        {
                            ExtractResource(kbResource, kbTemplatePath);
                        }
                        catch
                        {
                            // Non-critical
                        }
                    }

                    // Update .addin file to point to update folder
                    string addinContent = GetResourceText(addinResource);
                    var doc = XDocument.Parse(addinContent);

                    foreach (var assemblyElement in doc.Descendants("Assembly"))
                    {
                        string currentValue = assemblyElement.Value;
                        if (!currentValue.Contains("\\") && !currentValue.Contains("/"))
                        {
                            assemblyElement.Value = Path.Combine(updateFolder, currentValue);
                        }
                    }

                    doc.Save(addinPath);

                    installationResults.Add(new InstallationResult
                    {
                        Year = installation.Year,
                        State = InstallState.UpdatedNeedsRestart
                    });
                }
            }
        }

        private void ExtractResourceSafe(string resourceName, string targetPath)
        {
            // This method throws IOException if file is locked
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    // Try to open target file exclusively to check if it's locked
                    using (FileStream fs = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        stream.CopyTo(fs);
                    }
                }
                else
                {
                    throw new Exception($"Resource {resourceName} not found");
                }
            }
        }

        private void ExtractDependenciesSafe(string year, string targetFolder)
        {
            // This method throws IOException if any file is locked
            var assembly = Assembly.GetExecutingAssembly();

            if (fileMappings != null && fileMappings.Count > 0)
            {
                // Use file mapping to extract deduplicated dependencies
                foreach (var mapping in fileMappings.Values)
                {
                    // Skip if this file doesn't apply to this year
                    if (!mapping.Years.Contains(year))
                        continue;

                    // Skip the main DLL (handled separately)
                    if (mapping.TargetFileName == "revit-ballet.dll")
                        continue;

                    // Find the resource with this name
                    var resourceName = assembly.GetManifestResourceNames()
                        .FirstOrDefault(r => r.EndsWith(mapping.ResourceName));

                    if (resourceName == null)
                        continue;

                    string targetPath = Path.Combine(targetFolder, mapping.TargetFileName);

                    // This will throw if file is locked
                    ExtractResourceSafe(resourceName, targetPath);
                }
            }
            else
            {
                // Fallback to old method if no mapping file
                string yearPattern = $"_{year}.";
                var dependencyResources = assembly.GetManifestResourceNames()
                    .Where(r => r.Contains(yearPattern) && r.EndsWith(".dll") && !r.EndsWith("revit-ballet.dll"))
                    .ToList();

                foreach (var depResource in dependencyResources)
                {
                    int yearDotIndex = depResource.IndexOf(yearPattern);
                    if (yearDotIndex < 0) continue;

                    string fileName = depResource.Substring(yearDotIndex + yearPattern.Length);
                    string targetPath = Path.Combine(targetFolder, fileName);

                    ExtractResourceSafe(depResource, targetPath);
                }
            }
        }

        private void ExtractDependencies(string year, string targetFolder)
        {
            var assembly = Assembly.GetExecutingAssembly();

            if (fileMappings != null && fileMappings.Count > 0)
            {
                // Use file mapping to extract deduplicated dependencies
                foreach (var mapping in fileMappings.Values)
                {
                    // Skip if this file doesn't apply to this year
                    if (!mapping.Years.Contains(year))
                        continue;

                    // Skip the main DLL (handled separately)
                    if (mapping.TargetFileName == "revit-ballet.dll")
                        continue;

                    // Find the resource with this name
                    var resourceName = assembly.GetManifestResourceNames()
                        .FirstOrDefault(r => r.EndsWith(mapping.ResourceName));

                    if (resourceName == null)
                        continue;

                    string targetPath = Path.Combine(targetFolder, mapping.TargetFileName);

                    try
                    {
                        ExtractResource(resourceName, targetPath);
                    }
                    catch
                    {
                        // Non-critical - continue
                    }
                }
            }
            else
            {
                // Fallback to old method if no mapping file
                string yearPattern = $"_{year}.";
                var dependencyResources = assembly.GetManifestResourceNames()
                    .Where(r => r.Contains(yearPattern) && r.EndsWith(".dll") && !r.EndsWith("revit-ballet.dll"))
                    .ToList();

                foreach (var depResource in dependencyResources)
                {
                    int lastDot = depResource.LastIndexOf('.');
                    if (lastDot <= 0) continue;

                    int yearDotIndex = depResource.IndexOf(yearPattern);
                    if (yearDotIndex < 0) continue;

                    string fileName = depResource.Substring(yearDotIndex + yearPattern.Length);
                    string targetPath = Path.Combine(targetFolder, fileName);

                    try
                    {
                        ExtractResource(depResource, targetPath);
                    }
                    catch
                    {
                        // Non-critical - continue
                    }
                }
            }
        }

        private void CopyDllsToRuntimeLocation(string targetDir)
        {
            // InvokeAddinCommand expects DLLs at: %AppData%\revit-ballet\commands\bin\{year}\
            // Copy DLLs from embedded resources to this location for development workflow

            var assembly = Assembly.GetExecutingAssembly();

            foreach (var installation in installations)
            {
                // Create runtime directory for this Revit version
                string runtimeBinDir = Path.Combine(targetDir, "commands", "bin", installation.Year);
                Directory.CreateDirectory(runtimeBinDir);

                // Find version-specific DLL resource
                string dllResourcePattern = $"_{installation.Year}.revit-ballet.dll";
                var dllResource = assembly.GetManifestResourceNames()
                    .FirstOrDefault(r => r.Contains(dllResourcePattern));

                if (dllResource == null)
                {
                    // Try without year prefix (fallback)
                    dllResource = assembly.GetManifestResourceNames()
                        .FirstOrDefault(r => r.EndsWith("revit-ballet.dll") && !r.Contains("._20"));
                }

                if (dllResource == null)
                {
                    continue; // Skip this version
                }

                // Extract main DLL to runtime location
                string runtimeDllPath = Path.Combine(runtimeBinDir, "revit-ballet.dll");
                ExtractResource(dllResource, runtimeDllPath);

                // Extract dependencies to runtime location using file mapping or fallback
                if (fileMappings != null && fileMappings.Count > 0)
                {
                    foreach (var mapping in fileMappings.Values)
                    {
                        // Skip if this file doesn't apply to this year
                        if (!mapping.Years.Contains(installation.Year))
                            continue;

                        // Skip the main DLL (handled above)
                        if (mapping.TargetFileName == "revit-ballet.dll")
                            continue;

                        // Find the resource with this name
                        var resourceName = assembly.GetManifestResourceNames()
                            .FirstOrDefault(r => r.EndsWith(mapping.ResourceName));

                        if (resourceName == null)
                            continue;

                        string targetPath = Path.Combine(runtimeBinDir, mapping.TargetFileName);

                        try
                        {
                            ExtractResource(resourceName, targetPath);
                        }
                        catch
                        {
                            // Non-critical - continue
                        }
                    }
                }
                else
                {
                    // Fallback to old method
                    string yearPattern = $"_{installation.Year}.";
                    var dependencyResources = assembly.GetManifestResourceNames()
                        .Where(r => r.Contains(yearPattern) && r.EndsWith(".dll") && !r.EndsWith("revit-ballet.dll"))
                        .ToList();

                    foreach (var depResource in dependencyResources)
                    {
                        int yearDotIndex = depResource.IndexOf(yearPattern);
                        if (yearDotIndex < 0) continue;

                        string fileName = depResource.Substring(yearDotIndex + yearPattern.Length);
                        string targetPath = Path.Combine(runtimeBinDir, fileName);

                        try
                        {
                            ExtractResource(depResource, targetPath);
                        }
                        catch
                        {
                            // Non-critical - continue
                        }
                    }
                }
            }
        }

        private string GetResourceText(string resourceName)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        private void LoadFileMapping()
        {
            fileMappings = new Dictionary<string, FileMapping>();

            var assembly = Assembly.GetExecutingAssembly();
            var mappingResource = assembly.GetManifestResourceNames()
                .FirstOrDefault(r => r.EndsWith("file-mapping.txt"));

            if (mappingResource == null)
            {
                // No mapping file found - this is fine, might be an old-style installer
                return;
            }

            string mappingContent = GetResourceText(mappingResource);
            if (string.IsNullOrEmpty(mappingContent))
                return;

            // Parse the mapping file
            var lines = mappingContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                // Skip comments and empty lines
                if (line.Trim().StartsWith("#") || string.IsNullOrWhiteSpace(line))
                    continue;

                // Format: ResourceName|TargetFileName|Years (comma-separated)
                var parts = line.Split('|');
                if (parts.Length != 3)
                    continue;

                string resourceName = parts[0].Trim();
                string targetFileName = parts[1].Trim();
                string yearsStr = parts[2].Trim();

                var years = yearsStr.Split(',').Select(y => y.Trim()).ToList();

                // Index by resource name for easy lookup
                fileMappings[resourceName] = new FileMapping
                {
                    ResourceName = resourceName,
                    TargetFileName = targetFileName,
                    Years = years
                };
            }
        }

        private bool ModifyKeyboardShortcuts(RevitInstallation installation)
        {
            try
            {
                string kbPath = Path.Combine(installation.RevitDataPath, "KeyboardShortcuts.xml");

                if (!File.Exists(kbPath))
                {
                    // KeyboardShortcuts.xml doesn't exist - silently skip
                    return false;
                }

                // Load existing KeyboardShortcuts.xml
                var doc = XDocument.Load(kbPath);

                // Get shortcuts to add from template
                var assembly = Assembly.GetExecutingAssembly();
                var kbResource = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("KeyboardShortcuts.xml"));

                if (kbResource == null)
                {
                    return false;
                }

                string templateContent = GetResourceText(kbResource);
                var templateDoc = XDocument.Parse(templateContent);

                // Extract External Tools shortcuts from template
                var externalToolsShortcuts = templateDoc.Descendants("ShortcutItem")
                    .Where(e => e.Attribute("CommandName")?.Value.StartsWith("External Tools:") == true)
                    .ToList();

                if (externalToolsShortcuts.Count == 0)
                {
                    return false;
                }

                // Get root Shortcuts element
                var shortcutsRoot = doc.Element("Shortcuts");
                if (shortcutsRoot == null)
                {
                    return false;
                }

                // Create backup
                string backupPath = kbPath + ".bak";
                if (!File.Exists(backupPath))
                {
                    File.Copy(kbPath, backupPath);
                }

                bool modified = false;

                // Add or update each External Tools shortcut
                foreach (var shortcut in externalToolsShortcuts)
                {
                    string commandId = shortcut.Attribute("CommandId")?.Value?.ToLower();

                    if (string.IsNullOrEmpty(commandId))
                        continue;

                    // Find existing shortcut with matching CommandId
                    var existingShortcut = shortcutsRoot.Descendants("ShortcutItem")
                        .FirstOrDefault(e => e.Attribute("CommandId")?.Value?.ToLower() == commandId);

                    if (existingShortcut != null)
                    {
                        // Update existing shortcut - copy all attributes from template
                        // This handles cases where shortcut exists but Shortcuts attribute is empty
                        bool needsUpdate = false;

                        foreach (var attr in shortcut.Attributes())
                        {
                            var existingValue = existingShortcut.Attribute(attr.Name)?.Value;
                            if (existingValue != attr.Value)
                            {
                                existingShortcut.SetAttributeValue(attr.Name, attr.Value);
                                needsUpdate = true;
                            }
                        }

                        if (needsUpdate)
                        {
                            modified = true;
                        }
                    }
                    else
                    {
                        // Create new element with attributes (not a copy) to avoid XML formatting issues
                        var newShortcut = new XElement("ShortcutItem");
                        foreach (var attr in shortcut.Attributes())
                        {
                            newShortcut.SetAttributeValue(attr.Name, attr.Value);
                        }
                        shortcutsRoot.Add(newShortcut);
                        modified = true;
                    }
                }

                if (modified)
                {
                    doc.Save(kbPath);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to modify keyboard shortcuts for Revit {installation.Year}: {ex.Message}",
                    "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        private void ShowMissingKeyboardShortcutsWarning()
        {
            var missingVersions = new List<string>();

            foreach (var installation in installations)
            {
                string kbPath = Path.Combine(installation.RevitDataPath, "KeyboardShortcuts.xml");
                if (!File.Exists(kbPath))
                {
                    missingVersions.Add(installation.Year);
                }
            }

            if (missingVersions.Count > 0)
            {
                // Sort versions numerically
                missingVersions.Sort((a, b) => int.Parse(a).CompareTo(int.Parse(b)));

                string versionsText = string.Join(", ", missingVersions.Select(y => $"Revit {y}"));

                MessageBox.Show(
                    $"KeyboardShortcuts.xml not found for: {versionsText}\n\n" +
                    "This is a Revit limitation - the file only exists after you modify at least one keyboard shortcut.\n\n" +
                    "To fix this:\n" +
                    "1. Open each Revit version listed above\n" +
                    "2. Go to View > User Interface > Keyboard Shortcuts\n" +
                    "3. Modify any shortcut (for example, add 'ET' to 'Type Properties' command)\n" +
                    "4. Click OK\n" +
                    "5. Re-run this installer\n\n" +
                    "The installer will now skip keyboard shortcut installation for these versions.",
                    "Revit Ballet - Manual Step Required",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void CopyInstallerForUninstall(string targetDir)
        {
            try
            {
                string source = Assembly.GetExecutingAssembly().Location;
                string dest = Path.Combine(targetDir, "uninstaller.exe");
                File.Copy(source, dest, true);
            }
            catch { }
        }

        private void RegisterUninstaller(string targetDir)
        {
            try
            {
                string uninstallerPath = Path.Combine(targetDir, "uninstaller.exe");

                using (var key = Registry.CurrentUser.CreateSubKey(UNINSTALL_KEY))
                {
                    key.SetValue("DisplayName", PRODUCT_NAME);
                    key.SetValue("DisplayVersion", "1.0.0");
                    key.SetValue("Publisher", "Revit Ballet");
                    key.SetValue("InstallLocation", targetDir);
                    key.SetValue("UninstallString", $"\"{uninstallerPath}\"");
                    key.SetValue("QuietUninstallString", $"\"{uninstallerPath}\" /quiet");
                    key.SetValue("NoModify", 1);
                    key.SetValue("NoRepair", 1);
                    key.SetValue("EstimatedSize", 1000);
                    key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
                    key.SetValue("DisplayIcon", uninstallerPath);
                }
            }
            catch { }
        }

        private void RegisterTrustedAddins()
        {
            try
            {
                // Read .addin file to extract all AddInId GUIDs
                var assembly = Assembly.GetExecutingAssembly();
                var addinResource = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("revit-ballet.addin"));

                if (addinResource == null)
                    return;

                string addinContent = GetResourceText(addinResource);
                var doc = XDocument.Parse(addinContent);

                // Extract all AddInId values
                var guids = doc.Descendants("AddInId")
                    .Select(e => e.Value.Trim())
                    .Where(g => !string.IsNullOrEmpty(g))
                    .ToList();

                if (guids.Count == 0)
                    return;

                // Register for each installed Revit version
                foreach (var installation in installations)
                {
                    string regPath = $@"SOFTWARE\Autodesk\Revit\Autodesk Revit {installation.Year}\CodeSigning";

                    try
                    {
                        using (var key = Registry.CurrentUser.CreateSubKey(regPath))
                        {
                            foreach (var guid in guids)
                            {
                                // Set DWORD value 1 to indicate "Always Load"
                                key.SetValue(guid, 1, RegistryValueKind.DWord);
                            }
                        }
                    }
                    catch
                    {
                        // Non-critical - continue with other versions
                    }
                }
            }
            catch
            {
                // Non-critical - installer can continue without this
            }
        }

        private List<RevitInstallation> CheckMissingShortcuts()
        {
            var installationsNeedingShortcuts = new List<RevitInstallation>();

            foreach (var installation in installations)
            {
                string kbPath = Path.Combine(installation.RevitDataPath, "KeyboardShortcuts.xml");

                if (!File.Exists(kbPath))
                {
                    // File doesn't exist - will need to be created by user first
                    continue;
                }

                try
                {
                    var doc = XDocument.Load(kbPath);
                    var shortcutsRoot = doc.Element("Shortcuts");

                    if (shortcutsRoot == null)
                        continue;

                    // Check if any External Tools shortcuts are missing
                    var existingExternalTools = shortcutsRoot.Descendants("ShortcutItem")
                        .Where(e => e.Attribute("CommandName")?.Value.StartsWith("External Tools:") == true)
                        .Select(e => e.Attribute("CommandId")?.Value?.ToLower())
                        .Where(id => !string.IsNullOrEmpty(id))
                        .ToHashSet();

                    // Get template shortcuts
                    var assembly = Assembly.GetExecutingAssembly();
                    var kbResource = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("KeyboardShortcuts.xml"));

                    if (kbResource == null)
                        continue;

                    string templateContent = GetResourceText(kbResource);
                    var templateDoc = XDocument.Parse(templateContent);

                    var templateShortcuts = templateDoc.Descendants("ShortcutItem")
                        .Where(e => e.Attribute("CommandName")?.Value.StartsWith("External Tools:") == true)
                        .Select(e => e.Attribute("CommandId")?.Value?.ToLower())
                        .Where(id => !string.IsNullOrEmpty(id))
                        .ToList();

                    // Check if any template shortcuts are missing
                    bool hasMissing = templateShortcuts.Any(id => !existingExternalTools.Contains(id));

                    if (hasMissing)
                    {
                        installationsNeedingShortcuts.Add(installation);
                    }
                }
                catch
                {
                    // If we can't read the file, skip this installation
                }
            }

            return installationsNeedingShortcuts;
        }

        private bool OfferToAddShortcuts(List<RevitInstallation> installationsNeedingShortcuts)
        {
            // Get list of shortcuts that will be added
            var assembly = Assembly.GetExecutingAssembly();
            var kbResource = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("KeyboardShortcuts.xml"));

            if (kbResource == null)
                return false;

            string templateContent = GetResourceText(kbResource);
            var templateDoc = XDocument.Parse(templateContent);

            var shortcuts = templateDoc.Descendants("ShortcutItem")
                .Where(e => e.Attribute("CommandName")?.Value.StartsWith("External Tools:") == true)
                .Select(e => new
                {
                    Command = e.Attribute("CommandName")?.Value?.Replace("External Tools:", "").Trim(),
                    Shortcut = e.Attribute("Shortcuts")?.Value ?? ""
                })
                .Where(s => !string.IsNullOrEmpty(s.Command))
                .ToList();

            // Create form with DataGridView
            var form = new Form
            {
                Text = "Add Keyboard Shortcuts",
                Width = 700,
                Height = 600,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var label = new Label
            {
                Text = "Revit Ballet keyboard shortcuts are missing for: " +
                       string.Join(", ", installationsNeedingShortcuts.Select(i => $"Revit {i.Year}")) +
                       $"\n\nThe following {shortcuts.Count} shortcuts will be added:",
                Location = new Point(20, 20),
                Size = new Size(640, 60),
                Font = new Font("Segoe UI", 10)
            };
            form.Controls.Add(label);

            var grid = new DataGridView
            {
                Location = new Point(20, 90),
                Size = new Size(640, 400),
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Command",
                HeaderText = "Command",
                FillWeight = 60
            });

            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Shortcut",
                HeaderText = "Keyboard Shortcut",
                FillWeight = 40
            });

            foreach (var shortcut in shortcuts)
            {
                grid.Rows.Add(shortcut.Command, shortcut.Shortcut);
            }

            form.Controls.Add(grid);

            var btnYes = new Button
            {
                Text = "Add Shortcuts",
                Location = new Point(460, 510),
                Size = new Size(120, 30),
                DialogResult = DialogResult.Yes
            };
            form.Controls.Add(btnYes);
            form.AcceptButton = btnYes;

            var btnNo = new Button
            {
                Text = "Cancel",
                Location = new Point(590, 510),
                Size = new Size(70, 30),
                DialogResult = DialogResult.No
            };
            form.Controls.Add(btnNo);
            form.CancelButton = btnNo;

            return form.ShowDialog() == DialogResult.Yes;
        }

        private void ExtractResource(string resourceName, string targetPath)
        {
            try
            {
                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        using (FileStream fs = new FileStream(targetPath, FileMode.Create))
                            stream.CopyTo(fs);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to extract resource {resourceName}: {ex.Message}");
            }
        }

        private class RevitInstallation
        {
            public string Year { get; set; }
            public string AddinsPath { get; set; }
            public string RevitDataPath { get; set; }
        }

        private class InstallationResult
        {
            public string Year { get; set; }
            public InstallState State { get; set; }
        }
    }

    public class BalletUninstaller
    {
        public void Run()
        {
            try
            {
                if (MessageBox.Show("Uninstall Revit Ballet?", "Revit Ballet",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    return;

                RemoveFromRevit();
                RestoreKeyboardShortcuts();

                string targetDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "revit-ballet");
                if (Directory.Exists(targetDir))
                {
                    try
                    {
                        Directory.Delete(targetDir, true);
                    }
                    catch { }
                }

                RemoveFromRegistry();

                MessageBox.Show("Revit Ballet uninstalled successfully.", "Uninstaller",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Uninstallation failed: {ex.Message}", "Uninstaller",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RemoveFromRevit()
        {
            // Remove addins from user's AppData only (installer only writes to AppData)
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string revitPath = Path.Combine(appData, "Autodesk", "Revit", "Addins");

            if (!Directory.Exists(revitPath))
                return;

            // Look for year folders (2011-2030)
            string[] yearDirs = Directory.GetDirectories(revitPath);

            foreach (string yearDir in yearDirs)
            {
                string addinFile = Path.Combine(yearDir, "revit-ballet.addin");
                if (File.Exists(addinFile))
                {
                    try { File.Delete(addinFile); } catch { }
                }

                string balletFolder = Path.Combine(yearDir, "revit-ballet");
                if (Directory.Exists(balletFolder))
                {
                    try { Directory.Delete(balletFolder, true); } catch { }
                }
            }
        }

        private void RestoreKeyboardShortcuts()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string autodeskPath = Path.Combine(appData, "Autodesk", "Revit");

            if (!Directory.Exists(autodeskPath))
                return;

            // Find all KeyboardShortcuts.xml files
            string[] kbFiles = Directory.GetFiles(autodeskPath, "KeyboardShortcuts.xml", SearchOption.AllDirectories);

            foreach (string kbPath in kbFiles)
            {
                try
                {
                    if (!File.Exists(kbPath))
                        continue;

                    var doc = XDocument.Load(kbPath);
                    var shortcutsRoot = doc.Element("Shortcuts");

                    if (shortcutsRoot == null)
                        continue;

                    // Remove all External Tools shortcuts
                    var externalToolsShortcuts = shortcutsRoot.Descendants("ShortcutItem")
                        .Where(e => e.Attribute("CommandName")?.Value.StartsWith("External Tools:") == true)
                        .ToList();

                    bool modified = false;
                    foreach (var shortcut in externalToolsShortcuts)
                    {
                        shortcut.Remove();
                        modified = true;
                    }

                    if (modified)
                    {
                        doc.Save(kbPath);
                    }

                    // Restore backup if it exists
                    string backupPath = kbPath + ".bak";
                    if (File.Exists(backupPath))
                    {
                        File.Delete(backupPath);
                    }
                }
                catch { }
            }
        }

        private void RemoveFromRegistry()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\RevitBallet", false);
            }
            catch { }
        }
    }
}
