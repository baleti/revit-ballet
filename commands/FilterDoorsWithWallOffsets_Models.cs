using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace FilterDoorsWithWallOffsets
{
    public class DoorOrientation
    {
        public XYZ DoorPoint { get; set; }
        public XYZ DoorFacing { get; set; }
        public XYZ DoorHand { get; set; }
        public XYZ WallDirection { get; set; }
        public XYZ LeftEdge { get; set; }
        public XYZ RightEdge { get; set; }
        public double DoorWidth { get; set; }
        public Curve HostWallCurve { get; set; }
        public double DoorStartParam { get; set; }
        public double DoorEndParam { get; set; }
    }

    public class WallEndpointSide
    {
        public bool IsInFront { get; set; }
        public bool IsInBack { get; set; }
        public bool IsLeft { get; set; }
        public bool IsRight { get; set; }
        public bool IsCenter { get; set; }
    }

    public class WallData
    {
        public ElementId WallId { get; set; }
        public Wall Wall { get; set; }
        public Curve Curve { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
    }

    public class DoorProcessingResult
    {
        public FamilyInstance Door { get; set; }
        public List<AdjacentWallInfo> AdjacentWalls { get; set; }
        public bool NoHostWall { get; set; }
        public string Error { get; set; }

        public DoorProcessingResult()
        {
            AdjacentWalls = new List<AdjacentWallInfo>();
            // WallCheckDiagnostics and diagnostic strings removed
        }
    }

    public class AdjacentWallInfo
    {
        public ElementId WallId { get; set; }
        public Wall Wall { get; set; }
        public WallPosition Position { get; set; }
        public double Distance { get; set; }
        public List<WallSide> WallSides { get; set; }
        public bool RequiresBothSides { get; set; }
        public Reference ClosestFaceReference { get; set; }
    }

    public class DimensionInfo
    {
        public Dimension Dimension { get; set; }
        public double Value { get; set; }
        public string OrientationLabel { get; set; }
        public ElementId WallId { get; set; }
        public bool IsInFront { get; set; }
        public bool RequiresBothSides { get; set; }
        
        // Additional properties for storing calculation data
        public Reference DoorReference { get; set; }
        public Reference WallReference { get; set; }
        public XYZ DoorPoint { get; set; }
        public XYZ WallPoint { get; set; }
        public XYZ HostWallDirection { get; set; }
        public XYZ DoorFacing { get; set; }
    }

    public enum WallPosition
    {
        Left,
        Right,
        Front
    }

    public enum WallSide
    {
        Front,
        Back
    }

    // WallCheckDiagnostic class removed
    // IsWallInFrontCheckResult class removed
}
