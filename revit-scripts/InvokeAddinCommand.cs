using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class InvokeAddinCommand : IExternalCommand
{
    private const string FolderName = "revit-scripts";
    private const string ConfigFileName = "InvokeAddinCommand-last-dll-path";
    private const string LastCommandFileName = "InvokeAddinCommand-history";
    private static readonly string ConfigFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName);
    private static readonly string ConfigFilePath = Path.Combine(ConfigFolderPath, ConfigFileName);
    private static readonly string LastCommandFilePath = Path.Combine(ConfigFolderPath, LastCommandFileName);

    // Dictionary to store loaded assemblies
    private Dictionary<string, Assembly> loadedAssemblies = new Dictionary<string, Assembly>();

    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        string dllPath = null;

        if (File.Exists(ConfigFilePath))
        {
            dllPath = File.ReadAllText(ConfigFilePath);
        }

        if (string.IsNullOrEmpty(dllPath) || !File.Exists(dllPath))
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "DLL files (*.dll)|*.dll";
                openFileDialog.Title = "Open C# DLL with Revit Commands";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    dllPath = openFileDialog.FileName;

                    if (!Directory.Exists(ConfigFolderPath))
                    {
                        Directory.CreateDirectory(ConfigFolderPath);
                    }
                    File.WriteAllText(ConfigFilePath, dllPath);
                }
            }
        }

        if (!string.IsNullOrEmpty(dllPath) && File.Exists(dllPath))
        {
            try
            {
                // Register the assembly resolve event handler
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

                // Load the main assembly
                Assembly assembly = LoadAssembly(dllPath);

                var commandEntries = new List<Dictionary<string, object>>();
                var commandTypes = new Dictionary<string, string>();
                foreach (var type in assembly.GetTypes().Where(t => typeof(IExternalCommand).IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract))
                {
                    var typeName = type.Name;
                    commandEntries.Add(new Dictionary<string, object> { { "Command", typeName } });
                    commandTypes[typeName] = type.FullName;
                }

                List<string> propertyNames = new List<string> { "Command" };
                var selectedCommand = CustomGUIs.DataGrid(commandEntries, propertyNames, false).FirstOrDefault();

                if (selectedCommand != null)
                {
                    string commandClassName = selectedCommand["Command"].ToString();
                    string fullCommandClassName = commandTypes[commandClassName];
                    Type commandType = assembly.GetType(fullCommandClassName);

                    if (commandType != null)
                    {
                        if (fullCommandClassName != "InvokeLastAddinCommand")
                        {
                            // Append to history file instead of overwriting
                            AppendToCommandHistory(fullCommandClassName);
                        }

                        object commandInstance = Activator.CreateInstance(commandType);
                        MethodInfo method = commandType.GetMethod("Execute", new Type[] { typeof(ExternalCommandData), typeof(string).MakeByRefType(), typeof(ElementSet) });

                        if (method != null)
                        {
                            object[] parameters = new object[] { commandData, message, elements };
                            return (Result)method.Invoke(commandInstance, parameters);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = $"An error occurred: {ex.Message}";
                if (ex.InnerException != null)
                {
                    message += $"\nInner Exception: {ex.InnerException.Message}";
                }
            }
            finally
            {
                // Unregister the assembly resolve event handler
                AppDomain.CurrentDomain.AssemblyResolve -= CurrentDomain_AssemblyResolve;
            }
        }

        return Result.Failed;
    }

    private void AppendToCommandHistory(string commandClassName)
    {
        try
        {
            // Ensure directory exists
            if (!Directory.Exists(ConfigFolderPath))
            {
                Directory.CreateDirectory(ConfigFolderPath);
            }

            // Append command with timestamp for better tracking
            string historyEntry = $"{commandClassName}";
            
            // Append to file (creates file if it doesn't exist)
            using (StreamWriter sw = File.AppendText(LastCommandFilePath))
            {
                sw.WriteLine(historyEntry);
            }
        }
        catch (Exception ex)
        {
            // Log error but don't fail the command execution
            TaskDialog.Show("Warning", $"Failed to update command history: {ex.Message}");
        }
    }

    private Assembly LoadAssembly(string assemblyPath)
    {
        string assemblyName = Path.GetFileNameWithoutExtension(assemblyPath);
        
        if (loadedAssemblies.ContainsKey(assemblyName))
        {
            return loadedAssemblies[assemblyName];
        }

        byte[] assemblyBytes = File.ReadAllBytes(assemblyPath);
        Assembly assembly = Assembly.Load(assemblyBytes);
        loadedAssemblies[assemblyName] = assembly;
        
        // Load all DLLs in the same directory
        string directory = Path.GetDirectoryName(assemblyPath);
        foreach (string dllFile in Directory.GetFiles(directory, "*.dll"))
        {
            if (dllFile != assemblyPath)
            {
                string dllName = Path.GetFileNameWithoutExtension(dllFile);
                if (!loadedAssemblies.ContainsKey(dllName))
                {
                    try
                    {
                        byte[] dllBytes = File.ReadAllBytes(dllFile);
                        Assembly dllAssembly = Assembly.Load(dllBytes);
                        loadedAssemblies[dllName] = dllAssembly;
                    }
                    catch (BadImageFormatException)
                    {
                        // Skip native DLLs or incompatible assemblies
                        continue;
                    }
                }
            }
        }

        return assembly;
    }

    private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
    {
        // Parse the assembly name
        AssemblyName assemblyName = new AssemblyName(args.Name);
        string shortName = assemblyName.Name;

        // Check if we've already loaded this assembly
        if (loadedAssemblies.ContainsKey(shortName))
        {
            return loadedAssemblies[shortName];
        }

        // Look for the assembly in the same directory as the main DLL
        string dllPath = File.ReadAllText(ConfigFilePath);
        string directory = Path.GetDirectoryName(dllPath);
        string assemblyPath = Path.Combine(directory, shortName + ".dll");

        if (File.Exists(assemblyPath))
        {
            try
            {
                Assembly assembly = LoadAssembly(assemblyPath);
                return assembly;
            }
            catch (Exception)
            {
                return null;
            }
        }

        return null;
    }
}
