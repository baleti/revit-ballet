using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class RehostToJoinedWalls : IExternalCommand
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
        
        // Setup geometry options (for higher detail and to compute references).
        Options geomOptions = new Options
        {
            ComputeReferences = true,
            IncludeNonVisibleObjects = true,
            DetailLevel = ViewDetailLevel.Fine
        };
        
        // Lists for valid elements and candidate walls per element.
        List<FamilyInstance> validHostElements = new List<FamilyInstance>();
        // Map each valid element to its candidate walls (walls joined to its current host wall).
        Dictionary<ElementId, List<Wall>> elementCandidates = new Dictionary<ElementId, List<Wall>>();
        // Build a union of candidate wall types across all elements.
        // Key: wall type id; Value: list of candidate walls of that type.
        Dictionary<int, List<Wall>> unionCandidateWallTypes = new Dictionary<int, List<Wall>>();
        
        // Process each selected element.
        foreach (ElementId id in selectedIds)
        {
            Element e = doc.GetElement(id);
            FamilyInstance fi = e as FamilyInstance;
            if (fi == null ||
               !(fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows ||
                 fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors))
            {
                continue;
            }
            
            // Determine the current host wall.
            Wall currentHostWall = fi.Host as Wall;
            if (currentHostWall == null)
            {
                TaskDialog.Show("Warning", $"Element id {fi.Id.IntegerValue} does not have a valid host wall. Skipping.");
                continue;
            }
            
            // Get walls joined to the current host wall.
            ICollection<ElementId> joinedIds = JoinGeometryUtils.GetJoinedElements(doc, currentHostWall);
            List<Wall> candidateWalls = new List<Wall>();
            foreach (ElementId joinedId in joinedIds)
            {
                Element joinedElem = doc.GetElement(joinedId);
                if (joinedElem is Wall wall)
                {
                    if (wall.Id != currentHostWall.Id)
                    {
                        candidateWalls.Add(wall);
                    }
                }
            }
            if (candidateWalls.Count == 0)
            {
                continue;
            }
            
            // Get the element's bounding box (needed for metric calculations).
            BoundingBoxXYZ fiBB = fi.get_BoundingBox(null);
            if (fiBB == null)
            {
                TaskDialog.Show("Warning", $"Element id {fi.Id.IntegerValue} does not have a bounding box. Skipping.");
                continue;
            }
            
            elementCandidates[fi.Id] = candidateWalls;
            
            // Update union candidate wall types.
            foreach (Wall candidate in candidateWalls)
            {
                int wallTypeId = candidate.WallType.Id.IntegerValue;
                if (!unionCandidateWallTypes.ContainsKey(wallTypeId))
                {
                    unionCandidateWallTypes[wallTypeId] = new List<Wall>();
                }
                if (!unionCandidateWallTypes[wallTypeId].Contains(candidate))
                {
                    unionCandidateWallTypes[wallTypeId].Add(candidate);
                }
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
                // Use the first candidate as representative.
                Wall repWall = kvp.Value.First();
                Dictionary<string, object> entry = new Dictionary<string, object>
                {
                    { "Wall Type Id", repWall.WallType.Id.IntegerValue },
                    { "Wall Type", repWall.WallType.Name }
                };
                entries.Add(entry);
                if (!wallTypeMapping.ContainsKey(repWall.WallType.Id.IntegerValue))
                {
                    wallTypeMapping.Add(repWall.WallType.Id.IntegerValue, repWall.WallType.Id.IntegerValue);
                }
            }
            
            List<Dictionary<string, object>> selectedEntries =
                CustomGUIs.DataGrid(entries, new List<string> { "Wall Type Id", "Wall Type" }, spanAllScreens: false);
            
            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                return Result.Failed;
            }
            
            chosenWallTypeId = Convert.ToInt32(selectedEntries.First()["Wall Type Id"]);
            if (!wallTypeMapping.ContainsKey(chosenWallTypeId))
            {
                TaskDialog.Show("Warning", "Selected wall type id not found. Operation aborted.");
                return Result.Failed;
            }
        }
        
        // Rehost each valid element on its candidate wall of the chosen type with the maximum cut metric.
        List<FamilyInstance> rehostedElements = new List<FamilyInstance>();
        List<FamilyInstance> skippedElements = new List<FamilyInstance>();
        
        using (Transaction trans = new Transaction(doc, "Rehost Elements"))
        {
            trans.Start();
            foreach (FamilyInstance fi in validHostElements)
            {
                // Retrieve the element's bounding box (used for projection fallback).
                BoundingBoxXYZ fiBB = fi.get_BoundingBox(null);
                if (fiBB == null)
                {
                    skippedElements.Add(fi);
                    continue;
                }
                
                List<Wall> candidates = elementCandidates[fi.Id];
                // Filter candidates by the chosen wall type.
                List<Wall> matchingCandidates = candidates.Where(w => w.WallType.Id.IntegerValue == chosenWallTypeId).ToList();
                if (matchingCandidates.Count == 0)
                {
                    skippedElements.Add(fi);
                    continue;
                }
                
                // Further filter candidates to those on the same level as the source element, if available.
                List<Wall> sameLevelCandidates = matchingCandidates.Where(w => w.LevelId == fi.LevelId).ToList();
                if (sameLevelCandidates.Count > 0)
                {
                    matchingCandidates = sameLevelCandidates;
                }
                
                // For each candidate, compute the cut metric and choose the one with the highest value.
                Wall chosenCandidate = matchingCandidates.First();
                double maxMetric = ComputeCutMetric(fi, chosenCandidate, geomOptions, fiBB);
                foreach (Wall candidate in matchingCandidates.Skip(1))
                {
                    double metric = ComputeCutMetric(fi, candidate, geomOptions, fiBB);
                    if (metric > maxMetric)
                    {
                        maxMetric = metric;
                        chosenCandidate = candidate;
                    }
                }
                
                // Activate the FamilySymbol if necessary.
                FamilySymbol symbol = fi.Symbol;
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                    doc.Regenerate();
                }
                
                // Get the insertion point (assumed to be a LocationPoint).
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
                
                // Create a new instance on the chosen candidate wall.
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
    
    // Computes the cut metric for a given element (fi) and candidate wall.
    // It uses a boolean intersection to compute volume first;
    // if that is negligible, it falls back to a projection-based area calculation.
    private double ComputeCutMetric(FamilyInstance fi, Wall wall, Options geomOptions, BoundingBoxXYZ fiBB)
    {
        double bestMetric = 0.0;
        GeometryElement geomFi = fi.get_Geometry(geomOptions);
        GeometryElement geomWall = wall.get_Geometry(geomOptions);
        if (geomFi == null || geomWall == null)
            return bestMetric;
        
        foreach (GeometryObject goFi in geomFi)
        {
            Solid solidFi = goFi as Solid;
            if (solidFi == null || solidFi.Volume < 1e-9)
                continue;
            
            foreach (GeometryObject goWall in geomWall)
            {
                Solid solidWall = goWall as Solid;
                if (solidWall == null || solidWall.Volume < 1e-9)
                    continue;
                
                double metric = 0.0;
                try
                {
                    Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(solidFi, solidWall, BooleanOperationsType.Intersect);
                    if (intersection != null)
                    {
                        metric = intersection.Volume;
                    }
                }
                catch
                {
                    // Ignore boolean operation failure.
                }
                // If the volume is very low, fallback to a projection-based approach.
                if (metric < 1e-6)
                {
                    double areaMetric = ComputeProjectedIntersectionArea(fiBB, wall);
                    if (areaMetric > metric)
                        metric = areaMetric;
                }
                if (metric > bestMetric)
                    bestMetric = metric;
            }
        }
        
        return bestMetric;
    }
    
    // A projection-based method to approximate the overlapping area between the element and wall.
    // This implementation projects the element's bounding box and the wall's bounding box
    // into a 2D coordinate system defined by the wall's exterior (normal) and vertical direction.
    private double ComputeProjectedIntersectionArea(BoundingBoxXYZ fiBB, Wall wall)
    {
        // Get the wall's horizontal direction from its location curve.
        LocationCurve locCurve = wall.Location as LocationCurve;
        if (locCurve == null)
            return 0.0;
        XYZ p0 = locCurve.Curve.GetEndPoint(0);
        XYZ p1 = locCurve.Curve.GetEndPoint(1);
        XYZ wallDir = (p1 - p0).Normalize();
        // Vertical direction (Z axis).
        XYZ up = new XYZ(0, 0, 1);
        // Wall's exterior normal is perpendicular to the wall's direction and up.
        XYZ wallNormal = wallDir.CrossProduct(up).Normalize();
        
        // Define a 2D coordinate system for projection:
        // u-axis = wallNormal, v-axis = up.
        // Project the eight corners of the element's bounding box.
        double fiMinU = double.MaxValue, fiMaxU = double.MinValue;
        double fiMinV = double.MaxValue, fiMaxV = double.MinValue;
        foreach (XYZ pt in GetCorners(fiBB))
        {
            double u = pt.DotProduct(wallNormal);
            double v = pt.DotProduct(up);
            fiMinU = Math.Min(fiMinU, u);
            fiMaxU = Math.Max(fiMaxU, u);
            fiMinV = Math.Min(fiMinV, v);
            fiMaxV = Math.Max(fiMaxV, v);
        }
        
        // Do the same for the wall's bounding box.
        BoundingBoxXYZ wallBB = wall.get_BoundingBox(null);
        if (wallBB == null)
            return 0.0;
        double wallMinU = double.MaxValue, wallMaxU = double.MinValue;
        double wallMinV = double.MaxValue, wallMaxV = double.MinValue;
        foreach (XYZ pt in GetCorners(wallBB))
        {
            double u = pt.DotProduct(wallNormal);
            double v = pt.DotProduct(up);
            wallMinU = Math.Min(wallMinU, u);
            wallMaxU = Math.Max(wallMaxU, u);
            wallMinV = Math.Min(wallMinV, v);
            wallMaxV = Math.Max(wallMaxV, v);
        }
        
        // Compute the intersection rectangle in the (u,v) plane.
        double interMinU = Math.Max(fiMinU, wallMinU);
        double interMaxU = Math.Min(fiMaxU, wallMaxU);
        double interMinV = Math.Max(fiMinV, wallMinV);
        double interMaxV = Math.Min(fiMaxV, wallMaxV);
        if (interMaxU > interMinU && interMaxV > interMinV)
        {
            return (interMaxU - interMinU) * (interMaxV - interMinV);
        }
        return 0.0;
    }
    
    // Helper: returns the eight corner points of a BoundingBoxXYZ.
    private IEnumerable<XYZ> GetCorners(BoundingBoxXYZ box)
    {
        yield return new XYZ(box.Min.X, box.Min.Y, box.Min.Z);
        yield return new XYZ(box.Min.X, box.Min.Y, box.Max.Z);
        yield return new XYZ(box.Min.X, box.Max.Y, box.Min.Z);
        yield return new XYZ(box.Min.X, box.Max.Y, box.Max.Z);
        yield return new XYZ(box.Max.X, box.Min.Y, box.Min.Z);
        yield return new XYZ(box.Max.X, box.Min.Y, box.Max.Z);
        yield return new XYZ(box.Max.X, box.Max.Y, box.Min.Z);
        yield return new XYZ(box.Max.X, box.Max.Y, box.Max.Z);
    }
}
