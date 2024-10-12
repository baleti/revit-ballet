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
    private const string LastCommandFileName = "InvokeAddinCommand-last-command";
    private static readonly string ConfigFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName);
    private static readonly string ConfigFilePath = Path.Combine(ConfigFolderPath, ConfigFileName);
    private static readonly string LastCommandFilePath = Path.Combine(ConfigFolderPath, LastCommandFileName);

    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        string dllPath = null;

        // Check if the config file exists and read the DLL path from it
        if (File.Exists(ConfigFilePath))
        {
            dllPath = File.ReadAllText(ConfigFilePath);
        }

        // If DLL path is not valid, prompt the user to select a DLL file
        if (string.IsNullOrEmpty(dllPath) || !File.Exists(dllPath))
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "DLL files (*.dll)|*.dll";
                openFileDialog.Title = "Open C# DLL with Revit Commands";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    dllPath = openFileDialog.FileName;

                    // Save the selected DLL path to the config file
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
                // Read the DLL into a byte array
                byte[] dllBytes = File.ReadAllBytes(dllPath);

                // Load the DLL from the byte array
                Assembly assembly = Assembly.Load(dllBytes);

                // Find all IExternalCommand implementing types in the assembly
                var commandEntries = new List<Dictionary<string, object>>();
                var commandTypes = new Dictionary<string, string>();
                foreach (var type in assembly.GetTypes().Where(t => typeof(IExternalCommand).IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract))
                {
                    var typeName = type.Name;
                    commandEntries.Add(new Dictionary<string, object> { { "Command", typeName } });
                    commandTypes[typeName] = type.FullName; // Store the full type name
                }

                // Use CustomGUIs.DataGrid to let user select a command
                List<string> propertyNames = new List<string> { "Command" };
                var selectedCommand = CustomGUIs.DataGrid(commandEntries, propertyNames, false).FirstOrDefault();

                if (selectedCommand != null)
                {
                    string commandClassName = selectedCommand["Command"].ToString();
                    string fullCommandClassName = commandTypes[commandClassName]; // Get the full type name
                    Type commandType = assembly.GetType(fullCommandClassName);

                    if (commandType != null)
                    {
                        // Save the selected command to the last-command file if it is not InvokeLastCommand
                        if (fullCommandClassName != "InvokeLastAddinCommand")  // <-- This is the modified check
                        {
                            File.WriteAllText(LastCommandFilePath, fullCommandClassName);
                        }

                        // Create an instance of the command type
                        object commandInstance = Activator.CreateInstance(commandType);

                        // Find the Execute method to invoke
                        MethodInfo method = commandType.GetMethod("Execute", new Type[] { typeof(ExternalCommandData), typeof(string).MakeByRefType(), typeof(ElementSet) });

                        if (method != null)
                        {
                            // Prepare parameters for the Execute method
                            object[] parameters = new object[] { commandData, message, elements };

                            // Invoke the Execute method
                            return (Result)method.Invoke(commandInstance, parameters);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = $"An error occurred: {ex.Message}";
            }
        }

        return Result.Failed;
    }
}

[Transaction(TransactionMode.Manual)]
public class InvokeLastAddinCommand : IExternalCommand
{
    private const string FolderName = "revit-scripts";
    private const string LastCommandFileName = "InvokeAddinCommand-last-command";
    private static readonly string ConfigFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName);
    private static readonly string LastCommandFilePath = Path.Combine(ConfigFolderPath, LastCommandFileName);

    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        // Check if the InvokeAddinCommand-last-command file exists
        if (!File.Exists(LastCommandFilePath))
        {
            // If not, invoke InvokeAddinCommand command
            var invokeAddinCommand = new InvokeAddinCommand();
            return invokeAddinCommand.Execute(commandData, ref message, elements);
        }

        // Read the command type name from the InvokeAddinCommand-last-command file
        string commandTypeName = File.ReadAllText(LastCommandFilePath);

        // Find the assembly that contains the command
        string dllPath = null;
        if (File.Exists(Path.Combine(ConfigFolderPath, "InvokeAddinCommand-last-dll-path")))
        {
            dllPath = File.ReadAllText(Path.Combine(ConfigFolderPath, "InvokeAddinCommand-last-dll-path"));
        }

        if (!string.IsNullOrEmpty(dllPath) && File.Exists(dllPath))
        {
            try
            {
                // Read the DLL into a byte array
                byte[] dllBytes = File.ReadAllBytes(dllPath);

                // Load the DLL from the byte array
                Assembly assembly = Assembly.Load(dllBytes);

                // Get the command type
                Type commandType = assembly.GetType(commandTypeName);

                if (commandType != null)
                {
                    // Create an instance of the command type
                    object commandInstance = Activator.CreateInstance(commandType);

                    // Find the Execute method to invoke
                    MethodInfo method = commandType.GetMethod("Execute", new Type[] { typeof(ExternalCommandData), typeof(string).MakeByRefType(), typeof(ElementSet) });

                    if (method != null)
                    {
                        // Prepare parameters for the Execute method
                        object[] parameters = new object[] { commandData, message, elements };

                        // Invoke the Execute method
                        return (Result)method.Invoke(commandInstance, parameters);
                    }
                }
            }
            catch (Exception ex)
            {
                message = $"An error occurred: {ex.Message}";
            }
        }

        return Result.Failed;
    }
}
