using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ChangeWorksetOfSelectedElements : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                // Get the UIDocument and Document
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                
                // Check if document is workshared
                if (!doc.IsWorkshared)
                {
                    TaskDialog.Show("Error", "This document is not workshared. Worksets are not available.");
                    return Result.Failed;
                }
                
                // Get current selection using custom method
                var selectedIds = uidoc.GetSelectionIds();
                
                if (selectedIds == null || selectedIds.Count == 0)
                {
                    TaskDialog.Show("Warning", "Please select elements before running this command.");
                    return Result.Cancelled;
                }
                
                // Get all user worksets in the document
                var worksets = new FilteredWorksetCollector(doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .ToList();
                
                if (worksets.Count == 0)
                {
                    TaskDialog.Show("Error", "No user worksets found in the document.");
                    return Result.Failed;
                }
                
                // Prepare workset data for DataGrid
                var worksetEntries = new List<Dictionary<string, object>>();
                foreach (Workset ws in worksets)
                {
                    var entry = new Dictionary<string, object>
                    {
                        { "Id", ws.Id.IntegerValue },
                        { "Name", ws.Name },
                        { "IsOpen", ws.IsOpen },
                        { "IsEditable", ws.IsEditable },
                        { "Owner", ws.Owner ?? "None" }
                    };
                    worksetEntries.Add(entry);
                }
                
                // Sort by name for better user experience
                worksetEntries = worksetEntries.OrderBy(e => e["Name"].ToString()).ToList();
                
                // Define properties to display in DataGrid
                var propertyNames = new List<string> { "Name", "IsOpen", "IsEditable", "Owner" };
                
                // Display DataGrid for workset selection
                var selectedWorksets = CustomGUIs.DataGrid(
                    worksetEntries,
                    propertyNames,
                    false,
                    null);
                
                if (selectedWorksets == null || selectedWorksets.Count == 0)
                {
                    return Result.Cancelled;
                }
                
                // Get the selected workset ID
                var selectedWorkset = selectedWorksets.First();
                int worksetId = Convert.ToInt32(selectedWorkset["Id"]);
                string worksetName = selectedWorkset["Name"].ToString();
                
                // Filter elements that can have their workset changed
                var elementsToChange = new List<Element>();
                var skippedElements = new List<string>();
                
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem != null)
                    {
                        // Check if element can have its workset changed
                        Parameter worksetParam = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                        if (worksetParam != null && !worksetParam.IsReadOnly)
                        {
                            elementsToChange.Add(elem);
                        }
                        else
                        {
                            string elemInfo = $"{elem.Category?.Name ?? "Unknown"} - {elem.Name ?? elem.Id.ToString()}";
                            skippedElements.Add(elemInfo);
                        }
                    }
                }
                
                if (elementsToChange.Count == 0)
                {
                    TaskDialog.Show("Warning", "None of the selected elements can have their workset changed.");
                    return Result.Failed;
                }
                
                // Start transaction to change worksets
                using (Transaction trans = new Transaction(doc, "Change Element Worksets"))
                {
                    trans.Start();
                    
                    int successCount = 0;
                    var failedElements = new List<string>();
                    
                    foreach (Element elem in elementsToChange)
                    {
                        try
                        {
                            Parameter worksetParam = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                            if (worksetParam != null)
                            {
                                worksetParam.Set(worksetId);
                                successCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            string elemInfo = $"{elem.Category?.Name ?? "Unknown"} - {elem.Name ?? elem.Id.ToString()}: {ex.Message}";
                            failedElements.Add(elemInfo);
                        }
                    }
                    
                    trans.Commit();
                    
                    // Only show dialog if there were issues or partial success
                    if (skippedElements.Count > 0 || failedElements.Count > 0)
                    {
                        // Prepare result message
                        string resultMessage = $"Successfully changed workset to '{worksetName}' for {successCount} element(s).";
                        
                        if (skippedElements.Count > 0)
                        {
                            resultMessage += $"\n\nSkipped {skippedElements.Count} element(s) (cannot change workset):";
                            resultMessage += "\n" + string.Join("\n", skippedElements.Take(10));
                            if (skippedElements.Count > 10)
                            {
                                resultMessage += $"\n... and {skippedElements.Count - 10} more";
                            }
                        }
                        
                        if (failedElements.Count > 0)
                        {
                            resultMessage += $"\n\nFailed to change {failedElements.Count} element(s):";
                            resultMessage += "\n" + string.Join("\n", failedElements.Take(10));
                            if (failedElements.Count > 10)
                            {
                                resultMessage += $"\n... and {failedElements.Count - 10} more";
                            }
                        }
                        
                        TaskDialog.Show("Change Workset Results", resultMessage);
                    }
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
