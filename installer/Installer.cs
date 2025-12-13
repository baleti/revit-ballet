using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;

namespace RevitBalletInstaller
{
    internal static class Program
    {
        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int dwProcessId);

        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        private const int ATTACH_PARENT_PROCESS = -1;

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string exeName = Path.GetFileName(Assembly.GetExecutingAssembly().Location).ToLower();
            bool isUninstaller = exeName.Contains("uninstall") || (args.Length > 0 && args[0] == "/uninstall");
            bool quietMode = args.Any(arg => arg.Equals("/q", StringComparison.OrdinalIgnoreCase) ||
                                             arg.Equals("/quiet", StringComparison.OrdinalIgnoreCase) ||
                                             arg.Equals("-q", StringComparison.OrdinalIgnoreCase) ||
                                             arg.Equals("--quiet", StringComparison.OrdinalIgnoreCase));

            // In quiet mode, attach to parent console to enable stdout/stderr output
            if (quietMode)
            {
                AttachConsole(ATTACH_PARENT_PROCESS);
            }

            if (isUninstaller)
            {
                new BalletUninstaller().Run(quietMode);
            }
            else
            {
                new BalletInstaller().Run(quietMode);
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

        public void Run(bool quietMode = false)
        {
            try
            {
                bool alreadyInstalled = IsInstalled();

                // Load file mapping from embedded resources
                LoadFileMapping();

                DetectRevitInstallations();

                if (installations.Count == 0)
                {
                    if (!quietMode)
                    {
                        MessageBox.Show("No Revit installations found.",
                            "Revit Ballet", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        Console.WriteLine("No Revit installations found.");
                    }
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

                // Deploy keyboard shortcuts (creates file if missing, or adds shortcuts if exists)
                foreach (var installation in installations)
                {
                    DeployKeyboardShortcuts(installation);
                }

                RegisterUninstaller(targetDir);

                // Register addins as trusted to avoid "Load Always/Load Once" prompts
                RegisterTrustedAddins();

                if (!quietMode)
                {
                    ShowSuccessDialog();
                }
                else
                {
                    OutputSuccessToConsole();
                }
            }
            catch (Exception ex)
            {
                if (!quietMode)
                {
                    MessageBox.Show($"Installation error: {ex.Message}\n\n{ex.StackTrace}",
                        "Revit Ballet Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    // In quiet mode, output error to stderr and exit with error code
                    Console.Error.WriteLine($"Installation error: {ex.Message}");
                    Console.Error.WriteLine(ex.StackTrace);
                    Environment.Exit(1);
                }
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

            Image logo = null;

            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var logoResourceName = assembly.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith(".ico"));

                if (logoResourceName != null)
                {
                    using (var stream = assembly.GetManifestResourceStream(logoResourceName))
                    {
                        if (stream != null)
                        {
                            using (var icon = new Icon(stream))
                            {
                                logo = icon.ToBitmap();
                            }
                        }
                    }
                }
            }
            catch { }

            int currentY = 30;

            if (logo != null)
            {
                var pictureBox = new PictureBox
                {
                    Image = logo,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Location = new Point(175, currentY),
                    Size = new Size(200, 150)
                };
                form.Controls.Add(pictureBox);
                currentY += 160;
            }
            else
            {
                var titleLabel = new Label
                {
                    Text = "Revit Ballet",
                    Font = new Font("Segoe UI", 24, FontStyle.Bold),
                    ForeColor = Color.FromArgb(50, 50, 50),
                    Location = new Point(0, currentY),
                    Size = new Size(550, 40),
                    TextAlign = ContentAlignment.MiddleCenter
                };
                form.Controls.Add(titleLabel);
                currentY += 60;
            }

            // Group results by state and sort by year (numerically)
            var freshInstalls = installationResults.Where(r => r.State == InstallState.FreshInstall)
                .OrderBy(r => int.Parse(r.Year)).ToList();
            var updatedSuccessfully = installationResults.Where(r => r.State == InstallState.UpdatedSuccessfully)
                .OrderBy(r => int.Parse(r.Year)).ToList();
            var needsRestart = installationResults.Where(r => r.State == InstallState.UpdatedNeedsRestart)
                .OrderBy(r => int.Parse(r.Year)).ToList();

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
                messageText += "\nRestart Revit to load updates or use commands immediately via InvokeAddinCommand.\n";
            }

            if (installationResults.Count == 0)
            {
                messageText = "Installation completed, but no Revit versions were updated.\nThis might indicate missing DLL resources.";
            }

            using (var g = form.CreateGraphics())
            {
                var textSize = g.MeasureString(messageText, new Font("Segoe UI", 10), 450);
                var textHeight = (int)Math.Ceiling(textSize.Height);

                var messageLabel = new Label
                {
                    Text = messageText,
                    Font = new Font("Segoe UI", 10),
                    ForeColor = Color.FromArgb(60, 60, 60),
                    Location = new Point(50, currentY),
                    Size = new Size(450, textHeight + 10),
                    TextAlign = ContentAlignment.TopCenter
                };
                form.Controls.Add(messageLabel);

                currentY += textHeight + 30;
            }

            var btnOK = new Button
            {
                Text = "OK",
                Font = new Font("Segoe UI", 10),
                Size = new Size(100, 30),
                Location = new Point(225, currentY),
                FlatStyle = FlatStyle.System,
                DialogResult = DialogResult.OK
            };
            form.Controls.Add(btnOK);
            form.AcceptButton = btnOK;

            currentY += 50;
            form.ClientSize = new Size(550, currentY);

            form.ShowDialog();
        }

        private void OutputSuccessToConsole()
        {
            // Group results by state and sort by year (numerically)
            var freshInstalls = installationResults.Where(r => r.State == InstallState.FreshInstall)
                .OrderBy(r => int.Parse(r.Year)).ToList();
            var updatedSuccessfully = installationResults.Where(r => r.State == InstallState.UpdatedSuccessfully)
                .OrderBy(r => int.Parse(r.Year)).ToList();
            var needsRestart = installationResults.Where(r => r.State == InstallState.UpdatedNeedsRestart)
                .OrderBy(r => int.Parse(r.Year)).ToList();

            if (freshInstalls.Count > 0)
            {
                Console.WriteLine("Successfully installed to:");
                foreach (var result in freshInstalls)
                {
                    Console.WriteLine($"  - Revit {result.Year}");
                }
            }

            if (updatedSuccessfully.Count > 0)
            {
                Console.WriteLine("Successfully updated:");
                foreach (var result in updatedSuccessfully)
                {
                    Console.WriteLine($"  - Revit {result.Year}");
                }
            }

            if (needsRestart.Count > 0)
            {
                Console.WriteLine("Updated (restart Revit to use new version):");
                foreach (var result in needsRestart)
                {
                    Console.WriteLine($"  - Revit {result.Year}");
                }
                Console.WriteLine("Restart Revit to load updates or use commands immediately via InvokeAddinCommand.");
            }

            if (installationResults.Count == 0)
            {
                Console.WriteLine("Installation completed, but no Revit versions were updated.");
                Console.WriteLine("This might indicate missing DLL resources.");
            }
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
            // - Check Windows Registry to find actually installed Revit versions
            // - Verify installations exist in Add/Remove Programs
            // - Install addins to AppData (current user) to avoid requiring admin privileges
            // - Revit loads addins from both ProgramData and AppData locations

            var detectedYears = new HashSet<string>();
            var installedYears = GetInstalledRevitYearsFromUninstall();

            // Check registry for installed Revit versions
            // Revit registers itself in HKEY_LOCAL_MACHINE under Autodesk\Revit
            try
            {
                using (var baseKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Autodesk"))
                {
                    if (baseKey != null)
                    {
                        foreach (string subKeyName in baseKey.GetSubKeyNames())
                        {
                            // Look for keys like "Revit", "Revit 2024", etc.
                            if (subKeyName.StartsWith("Revit", StringComparison.OrdinalIgnoreCase))
                            {
                                using (var revitKey = baseKey.OpenSubKey(subKeyName))
                                {
                                    if (revitKey != null)
                                    {
                                        // Check for version-specific subkeys
                                        foreach (string versionKey in revitKey.GetSubKeyNames())
                                        {
                                            // Extract year from version (e.g., "24.0" -> 2024, "19.0" -> 2019)
                                            string year = ExtractYearFromVersion(versionKey);
                                            if (year != null)
                                            {
                                                detectedYears.Add(year);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            // Also check HKEY_CURRENT_USER for per-user installations
            try
            {
                using (var baseKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Autodesk"))
                {
                    if (baseKey != null)
                    {
                        foreach (string subKeyName in baseKey.GetSubKeyNames())
                        {
                            if (subKeyName.StartsWith("Revit", StringComparison.OrdinalIgnoreCase))
                            {
                                using (var revitKey = baseKey.OpenSubKey(subKeyName))
                                {
                                    if (revitKey != null)
                                    {
                                        foreach (string versionKey in revitKey.GetSubKeyNames())
                                        {
                                            string year = ExtractYearFromVersion(versionKey);
                                            if (year != null)
                                            {
                                                detectedYears.Add(year);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            // Fallback: Check for Revit executable in common installation paths
            // Always perform filesystem check to catch installations on non-C drives
            string[] programFilesPaths = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
            };

            // Also check other common drives (D:, E:, etc.)
            var drivesToCheck = new List<string>();
            try
            {
                var allDrives = DriveInfo.GetDrives()
                    .Where(d => d.DriveType == DriveType.Fixed && d.IsReady)
                    .Select(d => d.Name.TrimEnd('\\'))
                    .ToList();

                foreach (var drive in allDrives)
                {
                    drivesToCheck.Add(Path.Combine(drive, "Program Files"));
                    drivesToCheck.Add(Path.Combine(drive, "Program Files (x86)"));
                }
            }
            catch
            {
                // If drive enumeration fails, fall back to standard paths
                drivesToCheck.AddRange(programFilesPaths);
            }

            foreach (string programFiles in drivesToCheck.Distinct())
            {
                string autodeskPath = Path.Combine(programFiles, "Autodesk");

                if (Directory.Exists(autodeskPath))
                {
                    string[] revitDirs = Directory.GetDirectories(autodeskPath, "Revit*", SearchOption.TopDirectoryOnly);

                    foreach (string revitDir in revitDirs)
                    {
                        string dirName = Path.GetFileName(revitDir);

                        // Extract year from directory name (e.g., "Revit 2024" -> "2024")
                        string[] parts = dirName.Split(' ');
                        if (parts.Length >= 2 && int.TryParse(parts[1], out int yearNum) && yearNum >= 2011 && yearNum <= 2030)
                        {
                            string year = parts[1];
                            // Verify Revit.exe exists
                            if (File.Exists(Path.Combine(revitDir, "Revit.exe")))
                            {
                                detectedYears.Add(year);
                            }
                        }
                    }
                }
            }

            // Create installation entries - ALWAYS use AppData for installation (no admin required)
            // Only create installation entries for years that are actually installed
            var verifiedYears = detectedYears.Where(y => installedYears.Contains(y)).ToHashSet();

            string userAppData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            foreach (string year in verifiedYears.OrderBy(y => y))
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

        private HashSet<string> GetInstalledRevitYearsFromUninstall()
        {
            // Check Windows Uninstall registry to verify which Revit versions are actually installed
            var installedYears = new HashSet<string>();

            // Check both 32-bit and 64-bit uninstall keys
            var uninstallPaths = new[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            foreach (var path in uninstallPaths)
            {
                try
                {
                    using (var baseKey = Registry.LocalMachine.OpenSubKey(path))
                    {
                        if (baseKey == null) continue;

                        foreach (string subKeyName in baseKey.GetSubKeyNames())
                        {
                            try
                            {
                                using (var subKey = baseKey.OpenSubKey(subKeyName))
                                {
                                    if (subKey == null) continue;

                                    string displayName = subKey.GetValue("DisplayName") as string;
                                    if (string.IsNullOrEmpty(displayName)) continue;

                                    // Look for "Autodesk Revit XXXX" or "Revit XXXX"
                                    if (displayName.Contains("Revit") && displayName.Contains("Autodesk"))
                                    {
                                        // Extract year from display name
                                        var words = displayName.Split(new[] { ' ', '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
                                        foreach (var word in words)
                                        {
                                            if (int.TryParse(word, out int year) && year >= 2011 && year <= 2030)
                                            {
                                                installedYears.Add(year.ToString());
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Skip entries we can't read
                            }
                        }
                    }
                }
                catch
                {
                    // Skip if registry path is inaccessible
                }
            }

            // Also check HKEY_CURRENT_USER uninstall (for per-user installs)
            try
            {
                using (var baseKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
                {
                    if (baseKey != null)
                    {
                        foreach (string subKeyName in baseKey.GetSubKeyNames())
                        {
                            try
                            {
                                using (var subKey = baseKey.OpenSubKey(subKeyName))
                                {
                                    if (subKey == null) continue;

                                    string displayName = subKey.GetValue("DisplayName") as string;
                                    if (string.IsNullOrEmpty(displayName)) continue;

                                    if (displayName.Contains("Revit") && displayName.Contains("Autodesk"))
                                    {
                                        var words = displayName.Split(new[] { ' ', '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
                                        foreach (var word in words)
                                        {
                                            if (int.TryParse(word, out int year) && year >= 2011 && year <= 2030)
                                            {
                                                installedYears.Add(year.ToString());
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Skip entries we can't read
                            }
                        }
                    }
                }
            }
            catch
            {
                // Skip if registry path is inaccessible
            }

            return installedYears;
        }

        private string ExtractYearFromVersion(string versionString)
        {
            // Convert version strings to 4-digit years
            // Handles multiple formats:
            // - "24.0", "24.1", "19.0" -> "2024", "2024", "2019"
            // - "2024", "2025" -> "2024", "2025"

            if (string.IsNullOrEmpty(versionString))
                return null;

            // First check if it's already a 4-digit year (2011-2030)
            if (int.TryParse(versionString, out int fullYear) && fullYear >= 2011 && fullYear <= 2030)
            {
                return fullYear.ToString();
            }

            // Otherwise try to extract from version format like "24.0"
            string[] parts = versionString.Split('.');
            if (parts.Length > 0 && int.TryParse(parts[0], out int versionNum))
            {
                // Revit versions are typically 17-30 (for 2017-2030)
                if (versionNum >= 11 && versionNum <= 30)
                {
                    int year = 2000 + versionNum;
                    return year.ToString();
                }
            }

            return null;
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

                // Try to copy all files to main folder
                try
                {
                    // Extract main DLL
                    string mainDllPath = Path.Combine(mainFolder, "revit-ballet.dll");
                    ExtractResourceSafe(dllResource, mainDllPath);

                    // Extract dependencies
                    ExtractDependenciesSafe(installation.Year, mainFolder);

                    // Successfully updated main folder - always update .addin file to ensure command definitions are current
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

                    installationResults.Add(new InstallationResult
                    {
                        Year = installation.Year,
                        State = wasInstalled ? InstallState.UpdatedSuccessfully : InstallState.FreshInstall
                    });
                }
                catch (IOException)
                {
                    // Files are locked (Revit is running) - use update folder strategy
                    Directory.CreateDirectory(updateFolder);

                    // Extract to update folder
                    string updateDllPath = Path.Combine(updateFolder, "revit-ballet.dll");
                    ExtractResource(dllResource, updateDllPath);
                    ExtractDependencies(installation.Year, updateFolder);

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

        private void DeployKeyboardShortcuts(RevitInstallation installation)
        {
            try
            {
                string kbPath = Path.Combine(installation.RevitDataPath, "KeyboardShortcuts.xml");
                bool fileExists = File.Exists(kbPath);

                // Get embedded keyboard shortcuts resource for this year
                var assembly = Assembly.GetExecutingAssembly();
                string kbResourceName = $"KeyboardShortcuts-{installation.Year}.xml";
                var kbResource = assembly.GetManifestResourceNames()
                    .FirstOrDefault(r => r.EndsWith(kbResourceName));

                if (kbResource == null)
                {
                    // No keyboard shortcuts resource for this version
                    return;
                }

                string resourceContent = GetResourceText(kbResource);
                var resourceDoc = XDocument.Parse(resourceContent);

                // Get all addin-specific shortcuts (matching UUIDs from .addin file)
                var addinDoc = XDocument.Parse(GetResourceText(
                    assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("revit-ballet.addin"))));
                var addinUuids = addinDoc.Descendants("AddInId")
                    .Select(e => e.Value.Trim().ToLower())
                    .ToHashSet();

                if (!fileExists)
                {
                    // KeyboardShortcuts.xml doesn't exist - deploy the full merged version
                    // (contains both Revit defaults and revit-ballet customizations)
                    Directory.CreateDirectory(installation.RevitDataPath);
                    resourceDoc.Save(kbPath);
                }
                else
                {
                    // KeyboardShortcuts.xml exists - only add/update addin-specific shortcuts
                    var doc = XDocument.Load(kbPath);
                    var shortcutsRoot = doc.Element("Shortcuts");

                    if (shortcutsRoot == null)
                        return;

                    bool modified = false;

                    // Get addin shortcuts from resource
                    var addinShortcuts = resourceDoc.Descendants("ShortcutItem")
                        .Where(e => {
                            string cmdId = e.Attribute("CommandId")?.Value?.ToLower();
                            return cmdId != null && addinUuids.Contains(cmdId);
                        })
                        .ToList();

                    // Add or update each addin shortcut
                    foreach (var shortcut in addinShortcuts)
                    {
                        string commandId = shortcut.Attribute("CommandId")?.Value?.ToLower();

                        if (string.IsNullOrEmpty(commandId))
                            continue;

                        // Find existing shortcut with matching CommandId
                        var existingShortcut = shortcutsRoot.Descendants("ShortcutItem")
                            .FirstOrDefault(e => e.Attribute("CommandId")?.Value?.ToLower() == commandId);

                        if (existingShortcut != null)
                        {
                            // Update existing shortcut
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
                            // Create new shortcut element
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
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to deploy keyboard shortcuts for Revit {installation.Year}: {ex.Message}",
                    "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        public void Run(bool quietMode = false)
        {
            try
            {
                if (!quietMode)
                {
                    if (MessageBox.Show("Uninstall Revit Ballet?", "Revit Ballet",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                        return;
                }

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

                if (!quietMode)
                {
                    MessageBox.Show("Revit Ballet uninstalled successfully.", "Uninstaller",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Console.WriteLine("Revit Ballet uninstalled successfully.");
                }
            }
            catch (Exception ex)
            {
                if (!quietMode)
                {
                    MessageBox.Show($"Uninstallation failed: {ex.Message}", "Uninstaller",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    Console.Error.WriteLine($"Uninstallation failed: {ex.Message}");
                    Console.Error.WriteLine(ex.StackTrace);
                    Environment.Exit(1);
                }
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

            // Get addin UUIDs to remove
            var assembly = Assembly.GetExecutingAssembly();
            var addinResource = assembly.GetManifestResourceNames().FirstOrDefault(r => r.EndsWith("revit-ballet.addin"));

            HashSet<string> addinUuids = new HashSet<string>();
            if (addinResource != null)
            {
                try
                {
                    string addinContent = GetResourceText(addinResource);
                    var addinDoc = XDocument.Parse(addinContent);
                    addinUuids = addinDoc.Descendants("AddInId")
                        .Select(e => e.Value.Trim().ToLower())
                        .ToHashSet();
                }
                catch { }
            }

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

                    // Remove only revit-ballet shortcuts (matching our addin UUIDs)
                    var addinShortcuts = shortcutsRoot.Descendants("ShortcutItem")
                        .Where(e => {
                            string cmdId = e.Attribute("CommandId")?.Value?.ToLower();
                            return cmdId != null && addinUuids.Contains(cmdId);
                        })
                        .ToList();

                    bool modified = false;
                    foreach (var shortcut in addinShortcuts)
                    {
                        shortcut.Remove();
                        modified = true;
                    }

                    if (modified)
                    {
                        doc.Save(kbPath);
                    }
                }
                catch { }
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
