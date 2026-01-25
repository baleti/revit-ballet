using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace InvokeAddinCommandInNetwork
{
    static class Program
    {
        private static string logFilePath;

        [STAThread]
        static void Main()
        {
            // Setup logging
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDir = Path.Combine(appData, "revit-ballet", "runtime");
            Directory.CreateDirectory(logDir);
            logFilePath = Path.Combine(logDir, "InvokeAddinCommandInNetwork.log");

            Log($"=== InvokeAddinCommandInNetwork Started at {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ===");

            try
            {
                // Create message-only window for hotkey handling
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                var window = new MessageOnlyWindow();
                // For message-only windows, just run the message loop without a form
                Application.Run();
            }
            catch (Exception ex)
            {
                Log($"FATAL ERROR: {ex.Message}");
                Log($"  StackTrace: {ex.StackTrace}");
                MessageBox.Show($"Fatal error:\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void Log(string message)
        {
            try
            {
                File.AppendAllText(logFilePath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
            }
            catch
            {
                // Ignore logging errors
            }
        }
    }

    /// <summary>
    /// Message-only window for handling global hotkeys
    /// </summary>
    class MessageOnlyWindow : NativeWindow, IDisposable
    {
        private const int HWND_MESSAGE = -3;
        private const int WM_USER = 0x0400;
        private const int WM_SHOW_DIALOG = WM_USER + 1;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CreateWindowEx(
            int dwExStyle, string lpClassName, string lpWindowName, int dwStyle,
            int x, int y, int nWidth, int nHeight, IntPtr hWndParent,
            IntPtr hMenu, IntPtr hInstance, IntPtr lpParam);

        [DllImport("user32.dll")]
        private static extern bool DestroyWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        private GlobalHotkey hotkeyAltQ;
        private HotkeySequenceHandler sequenceHandler;

        public MessageOnlyWindow()
        {
            // Create a true message-only window (completely invisible, no window shown)
            CreateParams cp = new CreateParams();
            cp.Parent = new IntPtr(HWND_MESSAGE); // HWND_MESSAGE = message-only window
            this.CreateHandle(cp);

            // Initialize hotkey sequence: Alt+Q, then `
            Program.Log("Initializing hotkey sequence: Alt+Q, then `");

            hotkeyAltQ = new GlobalHotkey(Keys.Q, KeyModifiers.Alt, this.Handle);
            if (!hotkeyAltQ.Register())
            {
                Program.Log("Failed to register Alt+Q hotkey");
                MessageBox.Show("Failed to register Alt+Q hotkey. Another application may be using it.",
                    "Hotkey Registration Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                Program.Log("Alt+Q hotkey registered successfully");
            }

            sequenceHandler = new HotkeySequenceHandler(
                windowHandle: this.Handle,
                sequenceTimeoutMs: 2000
            );

            Program.Log("Hotkey handler initialized. Press Alt+Q then ` to show command dialog.");
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;

            try
            {
                if (m.Msg == WM_HOTKEY)
                {
                    Program.Log("Alt+Q pressed");
                    sequenceHandler.OnFirstKeyPressed();
                }
                else if (m.Msg == WM_SHOW_DIALOG)
                {
                    Program.Log("WM_SHOW_DIALOG received - calling ShowCommandDialog on UI thread");
                    try
                    {
                        ShowCommandDialog();
                    }
                    catch (Exception ex)
                    {
                        Program.Log($"ERROR in ShowCommandDialog: {ex.Message}");
                        Program.Log($"  StackTrace: {ex.StackTrace}");
                        MessageBox.Show($"Error showing command dialog:\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    base.WndProc(ref m);
                }
            }
            catch (Exception ex)
            {
                Program.Log($"ERROR in WndProc: {ex.Message}");
                Program.Log($"  StackTrace: {ex.StackTrace}");
            }
        }

        private void ShowCommandDialog()
        {
            Program.Log("ShowCommandDialog started");

            try
            {
                // Load Revit API DLLs
                if (!LoadRevitAPIs())
                {
                    MessageBox.Show("Could not load Revit API DLLs. Make sure Revit is installed.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Load revit-ballet.dll
                string dllPath = FindRevitBalletDll();
                if (dllPath == null)
                {
                    MessageBox.Show("Could not find revit-ballet.dll",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Assembly assembly = Assembly.LoadFrom(dllPath);
                Program.Log($"Loaded revit-ballet.dll from: {dllPath}");

                // Find CustomGUIs.DataGrid method
                // Note: Some types in revit-ballet.dll reference RevitAPIUI, which we don't have loaded.
                // We need to handle ReflectionTypeLoadException to get only successfully loaded types.
                Type customGuisType = null;
                try
                {
                    customGuisType = assembly.GetTypes().FirstOrDefault(t => t.Name == "CustomGUIs");
                }
                catch (ReflectionTypeLoadException ex)
                {
                    Program.Log($"Some types failed to load (expected - missing RevitAPIUI). Getting successfully loaded types...");
                    // Get the types that did load successfully (will have nulls for failed types)
                    customGuisType = ex.Types.Where(t => t != null && t.Name == "CustomGUIs").FirstOrDefault();

                    // Log which types failed (for debugging)
                    foreach (var loaderEx in ex.LoaderExceptions.Take(5))
                    {
                        Program.Log($"  Loader exception: {loaderEx.Message}");
                    }
                    if (ex.LoaderExceptions.Length > 5)
                    {
                        Program.Log($"  ... and {ex.LoaderExceptions.Length - 5} more");
                    }
                }

                if (customGuisType == null)
                {
                    MessageBox.Show("Could not find CustomGUIs class", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Program.Log($"Found CustomGUIs type");

                // Disable Revit API access for standalone execution
                var setRevitAccessMethod = customGuisType.GetMethod("SetRevitApiAccess",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new[] { typeof(bool) },
                    null);

                if (setRevitAccessMethod != null)
                {
                    setRevitAccessMethod.Invoke(null, new object[] { false });
                    Program.Log("Disabled Revit API access for standalone DataGrid");
                }
                else
                {
                    Program.Log("WARNING: Could not find SetRevitApiAccess method");
                }

                var dataGridMethod = customGuisType.GetMethod("DataGrid",
                    BindingFlags.Public | BindingFlags.Static);

                if (dataGridMethod == null)
                {
                    MessageBox.Show("Could not find DataGrid method", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Build command list - hardcoded for network-compatible commands
                var commandEntries = new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        ["Command"] = "OpenRvtFilesInNewSessions"
                    }
                };

                var columns = new List<string> { "Command" };

                // Call DataGrid
                Program.Log("Calling CustomGUIs.DataGrid...");
                var result = (List<Dictionary<string, object>>)dataGridMethod.Invoke(
                    null,
                    new object[] { commandEntries, columns, false, null, null, false, null, false });

                Program.Log($"DataGrid returned {result?.Count ?? 0} selected items");

                if (result != null && result.Any())
                {
                    var selectedCommand = result.First();
                    string commandName = selectedCommand["Command"]?.ToString();

                    Program.Log($"Selected command: {commandName}");

                    // Execute the command handler
                    ExecuteCommandHandler(assembly, commandName);
                }
            }
            catch (Exception ex)
            {
                Program.Log($"Error in ShowCommandDialog: {ex.Message}");
                Program.Log($"  StackTrace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Program.Log($"  InnerException: {ex.InnerException.Message}");
                    Program.Log($"  InnerException StackTrace: {ex.InnerException.StackTrace}");
                }
                MessageBox.Show($"Error:\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Program.Log("ShowCommandDialog completed - ready for next hotkey");
        }

        /// <summary>
        /// Execute a command handler by calling its static method
        /// </summary>
        private void ExecuteCommandHandler(Assembly assembly, string commandName)
        {
            try
            {
                Program.Log($"Executing command: {commandName}");

                if (commandName == "OpenRvtFilesInNewSessions")
                {
                    // Find RevitFileHelper class (doesn't require RevitAPIUI)
                    Type helperType = null;
                    try
                    {
                        helperType = assembly.GetType("RevitBallet.Commands.RevitFileHelper");

                        if (helperType == null)
                        {
                            helperType = assembly.GetTypes()
                                .FirstOrDefault(t => t.Name == "RevitFileHelper");
                        }
                    }
                    catch (ReflectionTypeLoadException ex)
                    {
                        helperType = ex.Types
                            .Where(t => t != null && t.Name == "RevitFileHelper")
                            .FirstOrDefault();

                        Program.Log($"ReflectionTypeLoadException while loading RevitFileHelper");
                        foreach (var loaderEx in ex.LoaderExceptions.Take(3))
                        {
                            if (loaderEx != null)
                                Program.Log($"  Loader exception: {loaderEx.Message}");
                        }
                    }

                    if (helperType == null)
                    {
                        Program.Log("ERROR: Could not find RevitFileHelper class");
                        MessageBox.Show("Could not find RevitFileHelper class.\n\nCheck the log file for details.", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    Program.Log($"Found helper type: {helperType.FullName}");

                    // Call GetRevitFilesFromDocuments() static method
                    var getFilesMethod = helperType.GetMethod("GetRevitFilesFromDocuments",
                        BindingFlags.Public | BindingFlags.Static);

                    if (getFilesMethod == null)
                    {
                        Program.Log("ERROR: Could not find GetRevitFilesFromDocuments method");
                        MessageBox.Show("Could not find GetRevitFilesFromDocuments method", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    Program.Log("Calling GetRevitFilesFromDocuments...");
                    var rvtFiles = (List<Dictionary<string, object>>)getFilesMethod.Invoke(null, null);
                    Program.Log($"Found {rvtFiles.Count} .rvt files");

                    if (rvtFiles.Count == 0)
                    {
                        MessageBox.Show("No Revit files found in Documents folder.", "No Files",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // Show files in DataGrid
                    Type customGuisTypeForFiles = null;
                    try
                    {
                        customGuisTypeForFiles = assembly.GetTypes()
                            .FirstOrDefault(t => t.Name == "CustomGUIs");
                    }
                    catch (ReflectionTypeLoadException ex)
                    {
                        customGuisTypeForFiles = ex.Types
                            .Where(t => t != null && t.Name == "CustomGUIs")
                            .FirstOrDefault();
                    }

                    if (customGuisTypeForFiles == null)
                    {
                        Program.Log("ERROR: Could not find CustomGUIs class");
                        MessageBox.Show("Could not find CustomGUIs class for file selection", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var dataGridMethod = customGuisTypeForFiles.GetMethod("DataGrid",
                        BindingFlags.Public | BindingFlags.Static);

                    if (dataGridMethod == null)
                    {
                        Program.Log("ERROR: Could not find DataGrid method");
                        MessageBox.Show("Could not find DataGrid method for file selection", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var columns = new List<string>
                    {
                        "File Name",
                        "Revit Version",
                        "Central Model",
                        "Last Modified",
                        "Last Opened"
                    };

                    Program.Log("Showing files in DataGrid...");
                    var selectedFiles = (List<Dictionary<string, object>>)dataGridMethod.Invoke(
                        null,
                        new object[] { rvtFiles, columns, false, null, null, false, null, false });

                    Program.Log($"User selected {selectedFiles?.Count ?? 0} files");

                    if (selectedFiles != null && selectedFiles.Any())
                    {
                        // Open selected files using RevitFileHelper
                        var openFileMethod = helperType.GetMethod("OpenFileInRevit",
                            BindingFlags.Public | BindingFlags.Static);

                        if (openFileMethod == null)
                        {
                            Program.Log("ERROR: Could not find OpenFileInRevit method");
                            MessageBox.Show("Could not find OpenFileInRevit method", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        foreach (var file in selectedFiles)
                        {
                            string filePath = file["Path"].ToString();
                            string revitVersion = file["Revit Version"].ToString();

                            Program.Log($"Opening {filePath} in Revit {revitVersion}");
                            openFileMethod.Invoke(null, new object[] { filePath, revitVersion });
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"Unknown command: {commandName}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Program.Log($"Error executing command {commandName}: {ex.Message}");
                Program.Log($"  StackTrace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Program.Log($"  InnerException: {ex.InnerException.Message}");
                }
                MessageBox.Show($"Error executing command:\n\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        [DllImport("kernel32.dll")]
        private static extern uint GetLastError();

        /// <summary>
        /// Load Revit API DLLs from installed Revit
        /// </summary>
        private bool LoadRevitAPIs()
        {
            try
            {
                // Try to find Revit installation
                string revitPath = FindRevitInstallation();
                if (revitPath == null)
                {
                    Program.Log("Could not find Revit installation");
                    return false;
                }

                Program.Log($"Found Revit at: {revitPath}");

                // CRITICAL: Add Revit directory to DLL search path
                if (!SetDllDirectory(revitPath))
                {
                    uint error = GetLastError();
                    Program.Log($"WARNING: SetDllDirectory failed with error code: {error}");
                }
                else
                {
                    Program.Log($"Set DLL search path to: {revitPath}");
                }

                // Try loading known dependencies first
                string[] dependencies = {
                    "AdWindows.dll",
                    "UIFramework.dll",
                    "AdUIManagedCore.dll",
                    "AdUICore.dll"
                };

                foreach (string depName in dependencies)
                {
                    string depPath = Path.Combine(revitPath, depName);
                    if (File.Exists(depPath))
                    {
                        try
                        {
                            Assembly.LoadFrom(depPath);
                            Program.Log($"Pre-loaded dependency: {depName}");
                        }
                        catch (Exception ex)
                        {
                            Program.Log($"Could not pre-load {depName}: {ex.Message}");
                        }
                    }
                    else
                    {
                        Program.Log($"Dependency not found (may be optional): {depName}");
                    }
                }

                // Load Revit API DLLs
                // Try loading just RevitAPI.dll first - we may not need RevitAPIUI.dll for DataGrid
                string[] requiredDlls = { "RevitAPI.dll" };

                foreach (string dllName in requiredDlls)
                {
                    string dllPath = Path.Combine(revitPath, dllName);
                    if (!File.Exists(dllPath))
                    {
                        Program.Log($"Missing: {dllPath}");
                        return false;
                    }

                    try
                    {
                        Assembly.LoadFrom(dllPath);
                        Program.Log($"Loaded: {dllName}");
                    }
                    catch (FileNotFoundException fnfEx)
                    {
                        Program.Log($"ERROR loading {dllName}: {fnfEx.Message}");
                        Program.Log($"  FusionLog: {fnfEx.FusionLog}");
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Program.Log($"ERROR loading {dllName}: {ex.Message}");
                        if (ex.InnerException != null)
                        {
                            Program.Log($"  Inner: {ex.InnerException.Message}");
                            if (ex.InnerException is FileNotFoundException fnfInner)
                            {
                                Program.Log($"  Missing file: {fnfInner.FileName}");
                                Program.Log($"  FusionLog: {fnfInner.FusionLog}");
                            }
                        }
                        throw;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Program.Log($"Error loading Revit APIs: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Program.Log($"  Inner exception: {ex.InnerException.Message}");
                    if (ex.InnerException is FileNotFoundException fnf)
                    {
                        Program.Log($"  Missing assembly: {fnf.FileName}");
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// Find Revit installation directory from registry
        /// </summary>
        private string FindRevitInstallation()
        {
            var versions = new[] { "2024", "2023", "2022", "2021", "2020" };

            foreach (var version in versions)
            {
                try
                {
                    // Try HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk\Revit\{version}
                    using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Autodesk\Revit\{version}"))
                    {
                        if (key != null)
                        {
                            string installLocation = key.GetValue("InstallLocation") as string;
                            if (!string.IsNullOrEmpty(installLocation) && Directory.Exists(installLocation))
                            {
                                Program.Log($"Found Revit {version} via registry");
                                return installLocation;
                            }
                        }
                    }
                }
                catch
                {
                    // Try next version
                }

                // Try default installation path
                string defaultPath = $@"C:\Program Files\Autodesk\Revit {version}";
                if (Directory.Exists(defaultPath))
                {
                    Program.Log($"Found Revit {version} at default path");
                    return defaultPath;
                }
            }

            return null;
        }

        /// <summary>
        /// Find revit-ballet.dll in AppData, checking .update folders first
        /// </summary>
        private string FindRevitBalletDll()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string binDir = Path.Combine(appData, "revit-ballet", "commands", "bin");

            if (!Directory.Exists(binDir))
                return null;

            // Try to find matching Revit version DLL
            var versions = new[] { "2024", "2023", "2022", "2021", "2020" };

            foreach (var version in versions)
            {
                string versionDir = Path.Combine(binDir, version);

                // Check for update folders first (same mechanism as Startup.cs)
                var updateDirs = Directory.GetDirectories(binDir, version + ".update*");
                if (updateDirs.Length > 0)
                {
                    // Use the first update folder found (sorted by name, most recent timestamp first)
                    string updateDir = updateDirs.OrderByDescending(d => d).First();
                    string updateDllPath = Path.Combine(updateDir, "revit-ballet.dll");

                    if (File.Exists(updateDllPath))
                    {
                        Program.Log($"Found revit-ballet.dll for Revit {version} in update folder");
                        return updateDllPath;
                    }
                }

                // Fall back to main folder
                string dllPath = Path.Combine(versionDir, "revit-ballet.dll");
                if (File.Exists(dllPath))
                {
                    Program.Log($"Found revit-ballet.dll for Revit {version}");
                    return dllPath;
                }
            }

            return null;
        }

        public void Dispose()
        {
            hotkeyAltQ?.Dispose();
            sequenceHandler?.Dispose();
            if (this.Handle != IntPtr.Zero)
            {
                this.DestroyHandle();
            }
        }
    }

    /// <summary>
    /// Handles two-key hotkey sequence (Alt+Q, then `)
    /// </summary>
    class HotkeySequenceHandler : IDisposable
    {
        private readonly IntPtr targetWindowHandle;
        private readonly int sequenceTimeoutMs;
        private System.Threading.Timer sequenceTimer;
        private bool isWaitingForSecondKey;
        private IntPtr keyboardHookHandle = IntPtr.Zero;
        private LowLevelKeyboardProc keyboardHookCallback;

        // Keyboard hook constants
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_USER = 0x0400;
        private const int WM_SHOW_DIALOG = WM_USER + 1;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        public HotkeySequenceHandler(IntPtr windowHandle, int sequenceTimeoutMs)
        {
            this.targetWindowHandle = windowHandle;
            this.sequenceTimeoutMs = sequenceTimeoutMs;
        }

        public void OnFirstKeyPressed()
        {
            if (isWaitingForSecondKey)
            {
                // Reset if already waiting
                ResetSequence();
            }

            isWaitingForSecondKey = true;
            Program.Log("Waiting for backtick (`) key...");

            // Install keyboard hook to detect backtick
            keyboardHookCallback = KeyboardHookProc;
            using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                keyboardHookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardHookCallback,
                    GetModuleHandle(curModule.ModuleName), 0);
            }

            if (keyboardHookHandle != IntPtr.Zero)
            {
                Program.Log("Keyboard hook installed successfully");
            }
            else
            {
                Program.Log("Failed to install keyboard hook");
            }

            // Start sequence timer
            sequenceTimer = new System.Threading.Timer(_ =>
            {
                Program.Log("Sequence timeout - backtick not pressed in time");
                ResetSequence();
            }, null, sequenceTimeoutMs, System.Threading.Timeout.Infinite);
        }

        private IntPtr KeyboardHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);

                // Check for backtick (both US and UK layouts)
                // VK_OEM_3 = 192 (US: `)
                // VK_OEM_8 = 223 (UK: `)
                if (vkCode == 192 || vkCode == 223)
                {
                    Program.Log($"Backtick key detected (vkCode: {vkCode})");

                    if (isWaitingForSecondKey)
                    {
                        ResetSequence();
                        // Post message to UI thread instead of calling directly
                        PostMessage(targetWindowHandle, WM_SHOW_DIALOG, IntPtr.Zero, IntPtr.Zero);
                        Program.Log("Posted WM_SHOW_DIALOG to UI thread");
                        return (IntPtr)1; // Consume the key
                    }
                }
            }

            return CallNextHookEx(keyboardHookHandle, nCode, wParam, lParam);
        }

        private void ResetSequence()
        {
            isWaitingForSecondKey = false;
            sequenceTimer?.Dispose();
            sequenceTimer = null;

            // Remove keyboard hook
            if (keyboardHookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(keyboardHookHandle);
                keyboardHookHandle = IntPtr.Zero;
                Program.Log("Keyboard hook removed");
            }
        }

        public void Dispose()
        {
            ResetSequence();
        }
    }

    /// <summary>
    /// Global hotkey registration
    /// </summary>
    class GlobalHotkey : IDisposable
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private readonly int id;
        private readonly IntPtr handle;
        private bool registered;

        public GlobalHotkey(Keys key, KeyModifiers modifiers, IntPtr windowHandle)
        {
            this.id = this.GetHashCode();
            this.handle = windowHandle;
            this.Key = key;
            this.Modifiers = modifiers;
        }

        public Keys Key { get; private set; }
        public KeyModifiers Modifiers { get; private set; }

        public bool Register()
        {
            registered = RegisterHotKey(handle, id, (uint)Modifiers, (uint)Key);
            return registered;
        }

        public void Dispose()
        {
            if (registered)
            {
                UnregisterHotKey(handle, id);
                registered = false;
            }
        }
    }

    [Flags]
    enum KeyModifiers
    {
        None = 0,
        Alt = 1,
        Control = 2,
        Shift = 4,
        Windows = 8
    }
}
