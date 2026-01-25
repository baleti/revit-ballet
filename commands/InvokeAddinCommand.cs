using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using RevitBallet.Commands;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

[Transaction(TransactionMode.Manual)]
public class InvokeAddinCommand : IExternalCommand
{
    private const string LastCommandFileName = "InvokeAddinCommand-history";
    private static readonly string LastCommandFilePath = PathHelper.GetRuntimeFilePath(LastCommandFileName);

    // Dictionary to store loaded assemblies
    private Dictionary<string, Assembly> loadedAssemblies = new Dictionary<string, Assembly>();
    private string currentDllPath;

    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        // Detect Revit version
        string revitVersion = commandData.Application.Application.VersionNumber;

        // Load from same location as Revit loads from (Addins folder)
        // This allows us to benefit from the .update folder mechanism
        string addinsBasePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Autodesk",
            "Revit",
            "Addins",
            revitVersion
        );

        string mainFolder = Path.Combine(addinsBasePath, "revit-ballet");
        string mainDllPath = Path.Combine(mainFolder, "revit-ballet.dll");

        // Check for update folders (same mechanism as Startup.cs)
        var updateFolders = Directory.Exists(addinsBasePath)
            ? Directory.GetDirectories(addinsBasePath, "revit-ballet.update*")
            : new string[0];

        // Prefer update folders if they exist (newer version)
        if (updateFolders.Length > 0)
        {
            // Use the first update folder found (sorted by name, most recent timestamp first)
            string updateFolder = updateFolders.OrderByDescending(f => f).First();
            string updateDllPath = Path.Combine(updateFolder, "revit-ballet.dll");

            if (File.Exists(updateDllPath))
            {
                currentDllPath = updateDllPath;
            }
            else
            {
                currentDllPath = mainDllPath;
            }
        }
        else
        {
            currentDllPath = mainDllPath;
        }

        if (!File.Exists(currentDllPath))
        {
            TaskDialog.Show("Error", $"Could not find revit-ballet.dll at expected location:\n\n{currentDllPath}");
            return Result.Failed;
        }

        if (!string.IsNullOrEmpty(currentDllPath) && File.Exists(currentDllPath))
        {
            try
            {
                // Register the assembly resolve event handler
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

                // Load the main assembly
                Assembly assembly = LoadAssembly(currentDllPath);

                var commandEntries = new List<Dictionary<string, object>>();
                var commandTypes = new Dictionary<string, string>(); // Maps class name to full class name
                foreach (var type in assembly.GetTypes().Where(t => typeof(IExternalCommand).IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract))
                {
                    var className = type.Name;

                    commandEntries.Add(new Dictionary<string, object>
                    {
                        { "Command", className }
                    });
                    commandTypes[className] = type.FullName;
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

        // Load from bytes to avoid file locking (enables hot reload during development)
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

        if (string.IsNullOrEmpty(currentDllPath) || !File.Exists(currentDllPath))
        {
            return null;
        }

        // Look for the assembly in the same directory as the main DLL
        string directory = Path.GetDirectoryName(currentDllPath);
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
