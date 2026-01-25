using Autodesk.Revit.UI;
using System;
using System.IO;

/// <summary>
/// Logs command execution for debugging and analysis
/// </summary>
public static class CommandExecutionLogger
{
    /// <summary>
    /// Start logging a command execution
    /// </summary>
    public static CommandExecutionLog Start(string commandName, ExternalCommandData commandData)
    {
        return new CommandExecutionLog(commandName, commandData);
    }

    /// <summary>
    /// Execution log session - use with 'using' statement
    /// </summary>
    public class CommandExecutionLog : IDisposable
    {
        private readonly string commandName;
        private readonly DateTime startTime;
        private readonly string logFilePath;
        private Autodesk.Revit.UI.Result? result;

        public CommandExecutionLog(string commandName, ExternalCommandData commandData)
        {
            this.commandName = commandName;
            this.startTime = DateTime.Now;

            // Create log file path
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDir = Path.Combine(appData, "revit-ballet", "runtime", "command-logs");
            Directory.CreateDirectory(logDir);

            string safeCommandName = string.Join("_", commandName.Split(Path.GetInvalidFileNameChars()));
            this.logFilePath = Path.Combine(logDir, $"{safeCommandName}_{startTime:yyyyMMdd_HHmmss}.log");

            // Log start
            try
            {
                File.WriteAllText(logFilePath, $"Command: {commandName}\nStart: {startTime:yyyy-MM-dd HH:mm:ss.fff}\n\n");
            }
            catch
            {
                // Ignore logging errors
            }
        }

        /// <summary>
        /// Set the result of the command execution
        /// </summary>
        public void SetResult(Autodesk.Revit.UI.Result result)
        {
            this.result = result;
            try
            {
                File.AppendAllText(logFilePath, $"Result: {result}\n");
            }
            catch
            {
                // Ignore logging errors
            }
        }

        public void Dispose()
        {
            // Log completion
            try
            {
                var endTime = DateTime.Now;
                var duration = endTime - startTime;
                File.AppendAllText(logFilePath, $"\nEnd: {endTime:yyyy-MM-dd HH:mm:ss.fff}\nDuration: {duration.TotalMilliseconds:F2}ms\n");
                if (result.HasValue)
                {
                    File.AppendAllText(logFilePath, $"Final Result: {result.Value}\n");
                }
            }
            catch
            {
                // Ignore logging errors
            }
        }
    }
}
