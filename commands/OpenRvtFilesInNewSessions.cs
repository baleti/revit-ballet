using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitBallet.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class OpenRvtFilesInNewSessions : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var rvtFiles = RevitFileHelper.GetRevitFilesFromDocuments();

                if (rvtFiles.Count == 0)
                {
                    TaskDialog.Show("No Files", "No Revit files found in Documents folder.");
                    return Result.Cancelled;
                }

                // Show in DataGrid
                var columns = new List<string>
                {
                    "File Name",
                    "Revit Version",
                    "Central Model",
                    "Last Modified",
                    "Last Opened",
                    "Path"
                };

                var selectedFiles = CustomGUIs.DataGrid(rvtFiles, columns, false);

                if (selectedFiles.Count == 0)
                    return Result.Cancelled;

                // Open selected files
                foreach (var file in selectedFiles)
                {
                    string filePath = file["Path"].ToString();
                    string revitVersion = file["Revit Version"].ToString();

                    RevitFileHelper.OpenFileInRevit(filePath, revitVersion);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"Failed to open Revit files:\n\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
