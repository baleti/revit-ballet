using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class RehostToAdjacentWall : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get active document and UIDocument.
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        // Ensure at least one element is selected.
        ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
        if (selectedIds.Count == 0)
        {
            message = "Please select one or more doors or windows before running the command.";
            return Result.Failed;
        }
        
        // Lists for valid elements and candidate walls per element.
        List<FamilyInstance> validHostElements = new List<FamilyInstance>();
        // Dictionary mapping each valid element to its candidate walls (unique by wall type).
        Dictionary<ElementId, List<Wall>> elementCandidates = new Dictionary<ElementId, List<Wall>>();
        // Union of candidate wall types across all elements.
        // Key: wall type id, Value: a candidate wall instance (first encountered) of that type.
        Dictionary<int, Wall> unionCandidateWallTypes = new Dictionary<int, Wall>();
        
        // Process each selected element.
        foreach (ElementId id in selectedIds)
        {
            Element e = doc.GetElement(id);
            FamilyInstance fi = e as FamilyInstance;
            if (fi == null ||
               !(fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows ||
                 fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors))
            {
                // Skip elements that are not doors or windows.
                continue;
            }
            
            // Get the element's bounding box.
            BoundingBoxXYZ bb = fi.get_BoundingBox(null);
            if (bb == null)
            {
                TaskDialog.Show("Warning", $"Element id {fi.Id.IntegerValue} does not have a bounding box. Skipping.");
                continue;
            }
            
            // Collect all walls that intersect the element's bounding box.
            FilteredElementCollector wallCollector = new FilteredElementCollector(doc).OfClass(typeof(Wall));
            List<Wall> cutWalls = new List<Wall>();
            foreach (Element wallElem in wallCollector)
            {
                Wall wall = wallElem as Wall;
                if (wall != null)
                {
                    BoundingBoxXYZ wallBB = wall.get_BoundingBox(null);
                    if (wallBB != null && DoBoundingBoxesIntersect(bb, wallBB))
                    {
                        cutWalls.Add(wall);
                    }
                }
            }
            
            // Determine the current host wall.
            Wall currentHostWall = null;
            if (fi.Host != null)
                currentHostWall = doc.GetElement(fi.Host.Id) as Wall;
            if (currentHostWall == null)
            {
                TaskDialog.Show("Warning", $"Element id {fi.Id.IntegerValue} does not have a valid host wall. Skipping.");
                continue;
            }
            
            // Candidate walls: those cut by the element that are not the current host.
            List<Wall> candidateWalls = cutWalls.Where(w => w.Id != currentHostWall.Id).ToList();
            if (candidateWalls.Count == 0)
            {
                TaskDialog.Show("Warning", $"Element id {fi.Id.IntegerValue} does not have any candidate wall (other than its current host). Skipping.");
                continue;
            }
            
            // For candidate walls, remove duplicates by wall type (choose the first instance for each type).
            List<Wall> uniqueCandidates = candidateWalls.GroupBy(w => w.WallType.Id.IntegerValue)
                                                         .Select(g => g.First())
                                                         .ToList();
            // Save these candidates for the element.
            elementCandidates[fi.Id] = uniqueCandidates;
            // Update the union of candidate wall types.
            foreach (Wall candidate in uniqueCandidates)
            {
                int key = candidate.WallType.Id.IntegerValue;
                if (!unionCandidateWallTypes.ContainsKey(key))
                    unionCandidateWallTypes.Add(key, candidate);
            }
            
            validHostElements.Add(fi);
        }
        
        if (validHostElements.Count == 0)
        {
            TaskDialog.Show("Warning", "None of the selected elements have candidate walls. Operation aborted.");
            return Result.Failed;
        }
        
        // Prompt the user to choose a wall type from the union candidate wall types.
        int chosenWallTypeId = 0;
        if (unionCandidateWallTypes.Count == 1)
        {
            chosenWallTypeId = unionCandidateWallTypes.First().Key;
        }
        else
        {
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            Dictionary<int, int> wallTypeMapping = new Dictionary<int, int>();
            foreach (var kvp in unionCandidateWallTypes)
            {
                Wall wall = kvp.Value;
                Dictionary<string, object> entry = new Dictionary<string, object>
                {
                    { "Wall Type Id", wall.WallType.Id.IntegerValue },
                    { "Wall Type", wall.WallType.Name }
                };
                entries.Add(entry);
                if (!wallTypeMapping.ContainsKey(wall.WallType.Id.IntegerValue))
                    wallTypeMapping.Add(wall.WallType.Id.IntegerValue, wall.WallType.Id.IntegerValue);
            }
            
            List<Dictionary<string, object>> selectedEntries =
                CustomGUIs.DataGrid(entries, new List<string> { "Wall Type Id", "Wall Type" }, spanAllScreens: false);
            
            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                TaskDialog.Show("Warning", "No wall type selected from prompt. Operation aborted.");
                return Result.Failed;
            }
            
            chosenWallTypeId = Convert.ToInt32(selectedEntries.First()["Wall Type Id"]);
            if (!wallTypeMapping.ContainsKey(chosenWallTypeId))
            {
                TaskDialog.Show("Warning", "Selected wall type id not found. Operation aborted.");
                return Result.Failed;
            }
        }
        
        // Rehost each valid element on its candidate wall that matches the chosen wall type.
        List<FamilyInstance> rehostedElements = new List<FamilyInstance>();
        List<FamilyInstance> skippedElements = new List<FamilyInstance>();
        
        using (Transaction trans = new Transaction(doc, "Rehost Elements"))
        {
            trans.Start();
            foreach (FamilyInstance fi in validHostElements)
            {
                List<Wall> candidates = elementCandidates[fi.Id];
                // Look for a candidate wall of the chosen wall type.
                Wall chosenCandidate = candidates.FirstOrDefault(w => w.WallType.Id.IntegerValue == chosenWallTypeId);
                if (chosenCandidate == null)
                {
                    // This element wasn't cut by a wall of the chosen type.
                    skippedElements.Add(fi);
                    continue;
                }
                
                // Activate the FamilySymbol if necessary.
                FamilySymbol symbol = fi.Symbol;
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                    doc.Regenerate();
                }
                
                // Get the insertion point (assumed LocationPoint).
                LocationPoint locPoint = fi.Location as LocationPoint;
                if (locPoint == null)
                {
                    TaskDialog.Show("Warning", $"Element id {fi.Id.IntegerValue} does not have a point location. Skipping.");
                    skippedElements.Add(fi);
                    continue;
                }
                XYZ insertionPoint = locPoint.Point;
                
                // Get the level associated with the element.
                Level level = doc.GetElement(fi.LevelId) as Level;
                if (level == null)
                {
                    TaskDialog.Show("Warning", $"Could not determine the level for element id {fi.Id.IntegerValue}. Skipping.");
                    skippedElements.Add(fi);
                    continue;
                }
                
                // Create a new instance on the candidate wall for this element.
                FamilyInstance newInstance = doc.Create.NewFamilyInstance(
                    insertionPoint,
                    symbol,
                    chosenCandidate,
                    level,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                
                // (Optionally, copy additional parameters from the old element to the new instance here.)
                
                // Delete the original element.
                doc.Delete(fi.Id);
                rehostedElements.Add(fi);
            }
            trans.Commit();
        }
        
        return Result.Succeeded;
    }
    
    // Helper method to check if two bounding boxes intersect.
    private bool DoBoundingBoxesIntersect(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        return (bb1.Max.X >= bb2.Min.X && bb1.Min.X <= bb2.Max.X) &&
               (bb1.Max.Y >= bb2.Min.Y && bb1.Min.Y <= bb2.Max.Y) &&
               (bb1.Max.Z >= bb2.Min.Z && bb1.Min.Z <= bb2.Max.Z);
    }
}
