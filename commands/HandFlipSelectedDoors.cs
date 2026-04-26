using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;
namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    [CommandMeta("Door")]
    public class HandFlipSelectedDoors : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // Get the current document and UI document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get the current selection
                ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();

                // Filter for door instances
                List<FamilyInstance> doors = new List<FamilyInstance>();
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is FamilyInstance fi && fi.Category?.Id.AsLong() == (int)BuiltInCategory.OST_Doors)
                        doors.Add(fi);
                }

                if (doors.Count == 0)
                {
                    CustomGUIs.SetCurrentUIDocument(uidoc);
                    var allDoors = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType()
                        .Cast<FamilyInstance>().ToList();
                    var gridData = CustomGUIs.ConvertToDataGridFormat(allDoors, new List<string> { "Name" });
                    var chosen = CustomGUIs.DataGrid(gridData, new List<string> { "Name" }, false);
                    if (chosen == null) return Result.Cancelled;
                    doors = CustomGUIs.ExtractOriginalObjects<FamilyInstance>(chosen) ?? new List<FamilyInstance>();
                    if (doors.Count == 0) return Result.Succeeded;
                }

                // Start a transaction to flip the door hands
                using (Transaction trans = new Transaction(doc, "Flip Selected Door Hands"))
                {
                    trans.Start();

                    int flippedCount = 0;
                    List<string> errorMessages = new List<string>();

                    foreach (FamilyInstance door in doors)
                    {
                        try
                        {
                            // Flip the door hand orientation
                            door.flipHand();
                            flippedCount++;
                        }
                        catch (Exception ex)
                        {
                            // Collect error messages for doors that couldn't be flipped
                            errorMessages.Add($"Door {door.Id}: {ex.Message}");
                        }
                    }

                    trans.Commit();

                    // Show results
                    string resultMessage = $"Successfully flipped hand orientation of {flippedCount} door(s).";
                    
                    if (errorMessages.Count > 0)
                    {
                        resultMessage += $"\n\nFailed to flip hand of {errorMessages.Count} door(s):";
                        resultMessage += "\n" + string.Join("\n", errorMessages.Take(5)); // Show first 5 errors
                        
                        if (errorMessages.Count > 5)
                        {
                            resultMessage += $"\n... and {errorMessages.Count - 5} more errors.";
                        }
                    }

                    TaskDialog.Show("Flip Door Hands Complete", resultMessage);
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
