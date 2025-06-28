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
        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
        if (selectedIds.Count == 0)
        {
            message = "Please select one or more doors or windows before running the command.";
            return Result.Failed;
        }
        
        // Setup geometry options.
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
                if (joinedElem is Wall wall && wall.Id != currentHostWall.Id)
                {
                    candidateWalls.Add(wall);
                }
            }
            
            // Only process elements with exactly one or two candidate walls.
            if (candidateWalls.Count == 0 || candidateWalls.Count > 2)
            {
                continue;
            }
            
            elementCandidates[fi.Id] = candidateWalls;
            validHostElements.Add(fi);
        }
        
        if (validHostElements.Count == 0)
        {
            TaskDialog.Show("Warning", "None of the selected elements have valid candidate walls. Operation aborted.");
            return Result.Failed;
        }
        
        // Rehost each valid element on its candidate wall.
        List<FamilyInstance> rehostedElements = new List<FamilyInstance>();
        List<FamilyInstance> skippedElements = new List<FamilyInstance>();
        
        using (Transaction trans = new Transaction(doc, "Rehost Elements"))
        {
            trans.Start();
            foreach (FamilyInstance fi in validHostElements)
            {
                // Get the element's bounding box.
                BoundingBoxXYZ fiBB = fi.get_BoundingBox(null);
                if (fiBB == null)
                {
                    skippedElements.Add(fi);
                    continue;
                }
                
                // Retrieve the candidate walls.
                List<Wall> candidates = elementCandidates[fi.Id];
                
                // Filter candidates to those on the same level as the source element, if any.
                List<Wall> sameLevelCandidates = candidates.Where(w => w.LevelId == fi.LevelId).ToList();
                if (sameLevelCandidates.Count > 0)
                {
                    candidates = sameLevelCandidates;
                }
                
                Wall chosenCandidate = null;
                if (candidates.Count == 1)
                {
                    chosenCandidate = candidates.First();
                }
                else if (candidates.Count == 2)
                {
                    // Compute the cut metric and choose the candidate with the higher value.
                    chosenCandidate = candidates.First();
                    double maxMetric = ComputeCutMetric(fi, chosenCandidate, geomOptions, fiBB);
                    foreach (Wall candidate in candidates.Skip(1))
                    {
                        double metric = ComputeCutMetric(fi, candidate, geomOptions, fiBB);
                        if (metric > maxMetric)
                        {
                            maxMetric = metric;
                            chosenCandidate = candidate;
                        }
                    }
                }
                else
                {
                    // No valid candidate after filtering.
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
                
                // Get the insertion point.
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
                
                // Delete the original element.
                doc.Delete(fi.Id);
                rehostedElements.Add(fi);
            }
            trans.Commit();
        }
        
        return Result.Succeeded;
    }
    
    // Computes the cut metric for a given element and candidate wall.
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
                // Fallback to a projection-based approach if volume is very low.
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
    
    // Computes a projection-based intersection area between the element and wall.
    private double ComputeProjectedIntersectionArea(BoundingBoxXYZ fiBB, Wall wall)
    {
        LocationCurve locCurve = wall.Location as LocationCurve;
        if (locCurve == null)
            return 0.0;
        XYZ p0 = locCurve.Curve.GetEndPoint(0);
        XYZ p1 = locCurve.Curve.GetEndPoint(1);
        XYZ wallDir = (p1 - p0).Normalize();
        XYZ up = new XYZ(0, 0, 1);
        XYZ wallNormal = wallDir.CrossProduct(up).Normalize();
        
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
    
    // Returns the eight corner points of a BoundingBoxXYZ.
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
