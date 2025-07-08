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
    public class FaceFlipSelectedDoors : IExternalCommand
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
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("Flip Door Facing", "No elements selected. Please select one or more doors.");
                    return Result.Cancelled;
                }

                // Filter for door instances
                List<FamilyInstance> doors = new List<FamilyInstance>();
                
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    
                    // Check if the element is a FamilyInstance and is a door
                    if (elem is FamilyInstance familyInstance)
                    {
                        // Check if it's a door by category
                        if (familyInstance.Category != null && 
                            familyInstance.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
                        {
                            doors.Add(familyInstance);
                        }
                    }
                }

                if (doors.Count == 0)
                {
                    TaskDialog.Show("Flip Door Facing", "No doors found in the selection. Please select one or more doors.");
                    return Result.Cancelled;
                }

                // Start a transaction to flip the door facing
                using (Transaction trans = new Transaction(doc, "Flip Selected Door Facing"))
                {
                    trans.Start();

                    int flippedCount = 0;
                    List<string> errorMessages = new List<string>();

                    foreach (FamilyInstance door in doors)
                    {
                        try
                        {
                            // Flip the door facing orientation
                            door.flipFacing();
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
                    string resultMessage = $"Successfully flipped facing orientation of {flippedCount} door(s).";
                    
                    if (errorMessages.Count > 0)
                    {
                        resultMessage += $"\n\nFailed to flip facing of {errorMessages.Count} door(s):";
                        resultMessage += "\n" + string.Join("\n", errorMessages.Take(5)); // Show first 5 errors
                        
                        if (errorMessages.Count > 5)
                        {
                            resultMessage += $"\n... and {errorMessages.Count - 5} more errors.";
                        }
                    }

                    TaskDialog.Show("Flip Door Facing Complete", resultMessage);
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
