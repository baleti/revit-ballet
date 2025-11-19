using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

[Transaction(TransactionMode.Manual)]
public class InvokeLastAddinCommand : IExternalCommand
{
    private const string FolderName = "revit-ballet/runtime";
    private const string LastCommandFileName = "InvokeAddinCommand-history";
    private static readonly string ConfigFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName);
    private static readonly string LastCommandFilePath = Path.Combine(ConfigFolderPath, LastCommandFileName);

    private Dictionary<string, Assembly> loadedAssemblies = new Dictionary<string, Assembly>();
    private string currentDllPath;

    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        try
        {
            // Detect Revit version
            string revitVersion = commandData.Application.Application.VersionNumber;

            // Build path to revit-ballet.dll in the standard installation location
            currentDllPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "revit-ballet",
                "commands",
                "bin",
                revitVersion,
                "revit-ballet.dll"
            );

            if (!File.Exists(currentDllPath))
            {
                message = $"Could not find revit-ballet.dll at expected location:\n{currentDllPath}";
                return Result.Failed;
            }

            if (!File.Exists(LastCommandFilePath))
            {
                message = "No previous command found. Run a command using InvokeAddinCommand first.";
                return Result.Failed;
            }

            // Read the last command from the history file
            string commandClassName = GetLastCommand();

            if (string.IsNullOrEmpty(commandClassName))
            {
                message = "No command history found.";
                return Result.Failed;
            }

            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

            Assembly assembly = LoadAssembly(currentDllPath);
            Type commandType = assembly.GetType(commandClassName);

            if (commandType == null)
            {
                message = $"Could not find type: {commandClassName}";
                return Result.Failed;
            }

            object commandInstance = Activator.CreateInstance(commandType);
            MethodInfo method = commandType.GetMethod("Execute", new Type[] { typeof(ExternalCommandData), typeof(string).MakeByRefType(), typeof(ElementSet) });

            if (method != null)
            {
                object[] parameters = new object[] { commandData, message, elements };
                return (Result)method.Invoke(commandInstance, parameters);
            }
            else
            {
                message = "Execute method not found.";
                return Result.Failed;
            }
        }
        catch (Exception ex)
        {
            message = $"An error occurred: {ex.Message}";
            if (ex.InnerException != null)
            {
                message += $"\nInner Exception: {ex.InnerException.Message}";
            }
            return Result.Failed;
        }
        finally
        {
            AppDomain.CurrentDomain.AssemblyResolve -= CurrentDomain_AssemblyResolve;
        }
    }

    private string GetLastCommand()
    {
        try
        {
            // Read all lines and get the last non-empty line
            string[] lines = File.ReadAllLines(LastCommandFilePath);
            
            // Find the last non-empty line
            for (int i = lines.Length - 1; i >= 0; i--)
            {
                if (!string.IsNullOrWhiteSpace(lines[i]))
                {
                    return lines[i].Trim();
                }
            }
            
            return null;
        }
        catch
        {
            return null;
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
                        continue;
                    }
                }
            }
        }

        return assembly;
    }

    private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
    {
        AssemblyName assemblyName = new AssemblyName(args.Name);
        string shortName = assemblyName.Name;

        if (loadedAssemblies.ContainsKey(shortName))
        {
            return loadedAssemblies[shortName];
        }

        if (string.IsNullOrEmpty(currentDllPath) || !File.Exists(currentDllPath))
        {
            return null;
        }

        string directory = Path.GetDirectoryName(currentDllPath);
        string assemblyPath = Path.Combine(directory, shortName + ".dll");

        if (File.Exists(assemblyPath))
        {
            try
            {
                return LoadAssembly(assemblyPath);
            }
            catch (Exception)
            {
                return null;
            }
        }

        return null;
    }
}
