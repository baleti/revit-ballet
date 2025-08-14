using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

public partial class CopySelectedElementsAlongContainingGroupsByRooms
{
    // Cache for element test points to avoid recalculation
    private Dictionary<ElementId, List<XYZ>> _elementTestPointsCache = new Dictionary<ElementId, List<XYZ>>();

    // Cache for room data
    private Dictionary<ElementId, RoomData> _roomDataCache = new Dictionary<ElementId, RoomData>();

    // Cache for group bounding boxes
    private Dictionary<ElementId, BoundingBoxXYZ> _groupBoundingBoxCache = new Dictionary<ElementId, BoundingBoxXYZ>();

    // Spatial index for groups
    private Dictionary<int, List<Group>> _spatialIndex = new Dictionary<int, List<Group>>();
    private const double SPATIAL_GRID_SIZE = 50.0; // 50 feet grid cells

    // Performance optimization caches
    private Dictionary<string, double> _floorToFloorHeightCache = new Dictionary<string, double>();
    private Dictionary<string, bool> _roomPointContainmentCache = new Dictionary<string, bool>();
    private Dictionary<ElementId, IList<IList<BoundarySegment>>> _roomBoundaryCache = 
        new Dictionary<ElementId, IList<IList<BoundarySegment>>>();

    // Data structure for room information
    private class RoomData
    {
        public double LevelZ { get; set; }
        public bool IsUnbound { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
        public Level Level { get; set; }
        public double Height { get; set; }
        public double MinZ { get; set; }
        public double MaxZ { get; set; }
    }

    // DELETE THE BatchCopyData CLASS FROM HERE - IT'S ALREADY IN Copying.cs

    // Build room cache for all rooms
    private void BuildRoomCache(Document doc)
    {
        StartTiming("BuildRoomCache");
        
        FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
        List<Room> allRooms = roomCollector
            .OfClass(typeof(SpatialElement))
            .WhereElementIsNotElementType()
            .OfType<Room>()
            .Where(r => r != null && r.Area > 0)
            .ToList();

        foreach (Room room in allRooms)
        {
            Level roomLevel = doc.GetElement(room.LevelId) as Level;
            if (roomLevel == null) continue;

            double roomLevelElevation = roomLevel.Elevation;
            double roomHeight = room.LookupParameter("Height")?.AsDouble()
                ?? room.get_Parameter(BuiltInParameter.ROOM_HEIGHT)?.AsDouble()
                ?? UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Feet);

            RoomData data = new RoomData
            {
                Level = roomLevel,
                LevelZ = roomLevelElevation,
                Height = roomHeight,
                MinZ = roomLevelElevation,
                MaxZ = roomLevelElevation + roomHeight,
                BoundingBox = room.get_BoundingBox(null),
                IsUnbound = room.Volume <= 0.001
            };

            _roomDataCache[room.Id] = data;
        }
        
        EndTiming("BuildRoomCache");
    }

    // Rest of the file continues unchanged...
    // Pre-calculate all group bounding boxes and build spatial index
    private void PreCalculateGroupDataAndSpatialIndex(IList<Group> allGroups, Document doc)
    {
        _spatialIndex.Clear();

        foreach (Group group in allGroups)
        {
            // Calculate and cache bounding box
            BoundingBoxXYZ bb = CalculateGroupBoundingBox(group, doc);
            if (bb != null)
            {
                _groupBoundingBoxCache[group.Id] = bb;

                // Add to spatial index
                AddToSpatialIndex(group, bb);
            }
        }
    }

    // Calculate group bounding box (without checking cache)
    private BoundingBoxXYZ CalculateGroupBoundingBox(Group group, Document doc)
    {
        ICollection<ElementId> memberIds = group.GetMemberIds();
        if (memberIds.Count == 0) return null;

        double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

        bool hasValidBB = false;

        foreach (ElementId id in memberIds)
        {
            Element member = doc.GetElement(id);
            if (member != null)
            {
                BoundingBoxXYZ bb = member.get_BoundingBox(null);
                if (bb != null)
                {
                    hasValidBB = true;
                    minX = Math.Min(minX, bb.Min.X);
                    minY = Math.Min(minY, bb.Min.Y);
                    minZ = Math.Min(minZ, bb.Min.Z);
                    maxX = Math.Max(maxX, bb.Max.X);
                    maxY = Math.Max(maxY, bb.Max.Y);
                    maxZ = Math.Max(maxZ, bb.Max.Z);
                }
            }
        }

        if (!hasValidBB) return null;

        BoundingBoxXYZ groupBB = new BoundingBoxXYZ();
        groupBB.Min = new XYZ(minX, minY, minZ);
        groupBB.Max = new XYZ(maxX, maxY, maxZ);

        return groupBB;
    }

    // Add group to spatial index
    private void AddToSpatialIndex(Group group, BoundingBoxXYZ bb)
    {
        // Get all grid cells that this bounding box overlaps
        int minGridX = (int)Math.Floor(bb.Min.X / SPATIAL_GRID_SIZE);
        int maxGridX = (int)Math.Floor(bb.Max.X / SPATIAL_GRID_SIZE);
        int minGridY = (int)Math.Floor(bb.Min.Y / SPATIAL_GRID_SIZE);
        int maxGridY = (int)Math.Floor(bb.Max.Y / SPATIAL_GRID_SIZE);

        // Add group to all overlapping cells
        for (int x = minGridX; x <= maxGridX; x++)
        {
            for (int y = minGridY; y <= maxGridY; y++)
            {
                int cellKey = GetSpatialCellKey(x, y);
                if (!_spatialIndex.ContainsKey(cellKey))
                {
                    _spatialIndex[cellKey] = new List<Group>();
                }
                _spatialIndex[cellKey].Add(group);
            }
        }
    }

    // Get spatial cell key from grid coordinates
    private int GetSpatialCellKey(int gridX, int gridY)
    {
        // Simple hash combining x and y coordinates
        // Assumes grid coordinates are within reasonable bounds
        return (gridX + 10000) * 20000 + (gridY + 10000);
    }

    // Get groups that might intersect with the given bounding box
    private List<Group> GetSpatiallyRelevantGroups(BoundingBoxXYZ targetBB)
    {
        HashSet<Group> relevantGroups = new HashSet<Group>();

        // Get all grid cells that the target bounding box overlaps
        int minGridX = (int)Math.Floor(targetBB.Min.X / SPATIAL_GRID_SIZE);
        int maxGridX = (int)Math.Floor(targetBB.Max.X / SPATIAL_GRID_SIZE);
        int minGridY = (int)Math.Floor(targetBB.Min.Y / SPATIAL_GRID_SIZE);
        int maxGridY = (int)Math.Floor(targetBB.Max.Y / SPATIAL_GRID_SIZE);

        // Collect all groups from overlapping cells
        for (int x = minGridX; x <= maxGridX; x++)
        {
            for (int y = minGridY; y <= maxGridY; y++)
            {
                int cellKey = GetSpatialCellKey(x, y);
                if (_spatialIndex.ContainsKey(cellKey))
                {
                    foreach (Group group in _spatialIndex[cellKey])
                    {
                        relevantGroups.Add(group);
                    }
                }
            }
        }

        return relevantGroups.ToList();
    }

    // Pre-calculate test points for an element
    private List<XYZ> GetElementTestPoints(Element element)
    {
        List<XYZ> testPoints = new List<XYZ>();

        LocationPoint locPoint = element.Location as LocationPoint;
        LocationCurve locCurve = element.Location as LocationCurve;

        if (locPoint != null)
        {
            testPoints.Add(locPoint.Point);
        }
        else if (locCurve != null)
        {
            // For curves, check endpoints and midpoint
            Curve curve = locCurve.Curve;
            testPoints.Add(curve.GetEndPoint(0));
            testPoints.Add(curve.GetEndPoint(1));
            testPoints.Add(curve.Evaluate(0.5, true));
        }
        else
        {
            // For other elements, use bounding box center
            BoundingBoxXYZ bb = element.get_BoundingBox(null);
            if (bb != null)
            {
                testPoints.Add((bb.Min + bb.Max) * 0.5);
            }
        }

        return testPoints;
    }

    // Get overall bounding box for a list of elements
    private BoundingBoxXYZ GetOverallBoundingBox(List<Element> elements)
    {
        if (elements.Count == 0) return null;

        double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

        bool hasValidBB = false;

        foreach (Element elem in elements)
        {
            BoundingBoxXYZ bb = elem.get_BoundingBox(null);
            if (bb != null)
            {
                hasValidBB = true;
                minX = Math.Min(minX, bb.Min.X);
                minY = Math.Min(minY, bb.Min.Y);
                minZ = Math.Min(minZ, bb.Min.Z);
                maxX = Math.Max(maxX, bb.Max.X);
                maxY = Math.Max(maxY, bb.Max.Y);
                maxZ = Math.Max(maxZ, bb.Max.Z);
            }
        }

        if (!hasValidBB) return null;

        BoundingBoxXYZ overallBB = new BoundingBoxXYZ();
        overallBB.Min = new XYZ(minX, minY, minZ);
        overallBB.Max = new XYZ(maxX, maxY, maxZ);

        return overallBB;
    }

    // Check if two bounding boxes intersect
    private bool BoundingBoxesIntersect(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        return !(bb1.Max.X < bb2.Min.X || bb2.Max.X < bb1.Min.X ||
                 bb1.Max.Y < bb2.Min.Y || bb2.Max.Y < bb1.Min.Y ||
                 bb1.Max.Z < bb2.Min.Z || bb2.Max.Z < bb1.Min.Z);
    }
}
