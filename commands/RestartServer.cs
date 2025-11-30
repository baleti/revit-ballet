using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace RevitBallet.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class RestartServer : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Terminate the existing server
                RevitBalletServer.TerminateServer();
                
                // Wait a moment for cleanup
                System.Threading.Thread.Sleep(500);
                
                // Start a new server instance
                RevitBalletServer.InitializeServer();
                
                TaskDialog.Show("Server Restarted", 
                    "Revit Ballet server has been restarted successfully.\n\n" +
                    "Check runtime/server.log for the new session details.");
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Failed to restart server: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }
    }
}
