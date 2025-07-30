using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading.Tasks;
using Newtonsoft.Json;

// Aliases to resolve conflicts
using WinForm = System.Windows.Forms.Form;
using WinPoint = System.Drawing.Point;
using WinSize = System.Drawing.Size;
using WinColor = System.Drawing.Color;
using WinFont = System.Drawing.Font;
using WinControl = System.Windows.Forms.Control;
using RevitColor = Autodesk.Revit.DB.Color;

/// <summary>
/// Optimized ClashDetectionSelectedElements with spatial indexing and parallel processing support
/// </summary>
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ClashDetectionSelectedElements : IExternalCommand
{
    // Configuration
    private const double CLASH_TOLERANCE = 0.001; // 1mm tolerance
    private const string CLASH_DS_PREFIX = "CLASH_AREA_";
    private const double MIN_ELEMENT_SIZE = 0.1; // Minimum element size in feet
    private const double BBOX_OFFSET = 0.5; // Offset for bounding box in feet
    private const double SPATIAL_CELL_SIZE = 10.0; // Spatial grid cell size in feet
    private const bool USE_PARALLEL_PROCESSING = false; // Enable subprocess parallelization
    
    // Performance monitoring
    private Stopwatch _stopwatch = new Stopwatch();
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        try
        {
            _stopwatch.Start();
            
            // Get selected elements
            var selectedElements = GetSelectedElements(uidoc, doc);
            
            if (selectedElements.Count < 2)
            {
                TaskDialog.Show("Clash Detection", 
                    $"Please select at least 2 elements for clash detection.\n" +
                    $"Currently selected: {selectedElements.Count} element(s).");
                return Result.Cancelled;
            }
            
            // Extract all geometry data upfront (for potential parallel processing)
            var geometryData = ExtractGeometryData(selectedElements);
            
            List<ClashInfo> clashes;
            bool isCancelled;
            
            if (USE_PARALLEL_PROCESSING && geometryData.All(g => g.SerializedSolid != null))
            {
                // Use parallel subprocess approach
                (clashes, isCancelled) = PerformParallelClashDetection(doc, selectedElements, geometryData);
            }
            else
            {
                // Use optimized single-threaded approach with spatial indexing
                (clashes, isCancelled) = PerformOptimizedClashDetection(doc, selectedElements, geometryData);
            }
            
            if (clashes.Count == 0)
            {
                TaskDialog.Show("Clash Detection", "No clashes found between selected elements.");
                return Result.Succeeded;
            }
            
            // Create DirectShape markers for clashes
            CreateClashMarkers(doc, clashes);
            
            _stopwatch.Stop();
            
            // Report results
            string title = isCancelled ? "Clash Detection Cancelled" : "Clash Detection Complete";
            string finalReport = $"{title}\n" +
                          $"Elements Analyzed: {selectedElements.Count}\n" +
                          $"Clashes Found: {clashes.Count}\n" +
                          $"Processing Time: {_stopwatch.ElapsedMilliseconds}ms";
            
            TaskDialog.Show("Clash Detection Results", finalReport);
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
    
    #region Geometry Data Extraction
    
    /// <summary>
    /// Extract all geometry data upfront to minimize Revit API calls
    /// </summary>
    private List<GeometryData> ExtractGeometryData(List<ElementInfo> elements)
    {
        var geometryDataList = new List<GeometryData>();
        
        foreach (var elemInfo in elements)
        {
            var geomData = new GeometryData
            {
                ElementId = elemInfo.Element.Id,
                ElementInfo = elemInfo,
                BoundingBox = GetTransformedBoundingBox(elemInfo.Element, elemInfo.Transform),
                BoundingSphere = null,
                Solid = null,
                SerializedSolid = null
            };
            
            // Get solid
            var solid = GetSolid(elemInfo.Element);
            if (solid != null)
            {
                // Transform solid if needed
                if (!elemInfo.Transform.IsIdentity)
                {
                    solid = SolidUtils.CreateTransformed(solid, elemInfo.Transform);
                }
                geomData.Solid = solid;
                
                // Calculate bounding sphere for quick checks
                geomData.BoundingSphere = CalculateBoundingSphere(geomData.BoundingBox);
                
                // Try to serialize solid for parallel processing
                try
                {
                    geomData.SerializedSolid = SerializeSolid(solid);
                }
                catch
                {
                    // Serialization failed, will use in-process checking
                }
            }
            
            geometryDataList.Add(geomData);
        }
        
        return geometryDataList;
    }
    
    private BoundingSphere CalculateBoundingSphere(BoundingBoxXYZ bbox)
    {
        if (bbox == null) return null;
        
        var center = (bbox.Min + bbox.Max) * 0.5;
        var radius = center.DistanceTo(bbox.Max);
        
        return new BoundingSphere { Center = center, Radius = radius };
    }
    
    private SerializedSolid SerializeSolid(Solid solid)
    {
        // Extract faces and edges for external processing
        var serialized = new SerializedSolid
        {
            Faces = new List<SerializedFace>(),
            Volume = solid.Volume
        };
        
        foreach (Face face in solid.Faces)
        {
            var mesh = face.Triangulate();
            if (mesh == null) continue;
            
            var sFace = new SerializedFace
            {
                Vertices = new List<double[]>(),
                Triangles = new List<int[]>()
            };
            
            // Store vertices
            for (int i = 0; i < mesh.NumTriangles; i++)
            {
                var triangle = mesh.get_Triangle(i);
                for (int j = 0; j < 3; j++)
                {
                    var vertex = triangle.get_Vertex(j);
                    sFace.Vertices.Add(new[] { vertex.X, vertex.Y, vertex.Z });
                }
                sFace.Triangles.Add(new[] { i * 3, i * 3 + 1, i * 3 + 2 });
            }
            
            serialized.Faces.Add(sFace);
        }
        
        return serialized;
    }
    
    #endregion
    
    #region Spatial Indexing
    
    /// <summary>
    /// Spatial grid for efficient collision detection
    /// </summary>
    private class SpatialGrid
    {
        private Dictionary<long, List<GeometryData>> _grid = new Dictionary<long, List<GeometryData>>();
        private double _cellSize;
        
        public SpatialGrid(double cellSize = SPATIAL_CELL_SIZE)
        {
            _cellSize = cellSize;
        }
        
        public void AddElement(GeometryData geomData)
        {
            if (geomData.BoundingBox == null) return;
            
            var minCell = GetCellCoord(geomData.BoundingBox.Min);
            var maxCell = GetCellCoord(geomData.BoundingBox.Max);
            
            for (int x = minCell.X; x <= maxCell.X; x++)
            for (int y = minCell.Y; y <= maxCell.Y; y++)
            for (int z = minCell.Z; z <= maxCell.Z; z++)
            {
                long key = GetCellKey(x, y, z);
                if (!_grid.ContainsKey(key))
                    _grid[key] = new List<GeometryData>();
                _grid[key].Add(geomData);
            }
        }
        
        public HashSet<(GeometryData, GeometryData)> GetPotentialClashPairs()
        {
            var pairs = new HashSet<(GeometryData, GeometryData)>();
            var processed = new HashSet<string>();
            
            foreach (var cell in _grid.Values)
            {
                // Check all pairs within the cell
                for (int i = 0; i < cell.Count - 1; i++)
                {
                    for (int j = i + 1; j < cell.Count; j++)
                    {
                        var elem1 = cell[i];
                        var elem2 = cell[j];
                        
                        // Create unique key for this pair
                        var key1 = $"{elem1.ElementId.IntegerValue}_{elem2.ElementId.IntegerValue}";
                        var key2 = $"{elem2.ElementId.IntegerValue}_{elem1.ElementId.IntegerValue}";
                        
                        if (!processed.Contains(key1) && !processed.Contains(key2))
                        {
                            processed.Add(key1);
                            
                            // Quick sphere check first
                            if (elem1.BoundingSphere != null && elem2.BoundingSphere != null)
                            {
                                double distance = elem1.BoundingSphere.Center.DistanceTo(elem2.BoundingSphere.Center);
                                if (distance > elem1.BoundingSphere.Radius + elem2.BoundingSphere.Radius)
                                    continue; // No possible intersection
                            }
                            
                            pairs.Add((elem1, elem2));
                        }
                    }
                }
            }
            
            return pairs;
        }
        
        private (int X, int Y, int Z) GetCellCoord(XYZ point)
        {
            return (
                (int)Math.Floor(point.X / _cellSize),
                (int)Math.Floor(point.Y / _cellSize),
                (int)Math.Floor(point.Z / _cellSize)
            );
        }
        
        private long GetCellKey(int x, int y, int z)
        {
            // Simple spatial hashing
            const int offset = 1000000;
            return ((long)(x + offset) << 40) | ((long)(y + offset) << 20) | (long)(z + offset);
        }
    }
    
    #endregion
    
    #region Optimized Clash Detection
    
    private (List<ClashInfo> Clashes, bool IsCancelled) PerformOptimizedClashDetection(
        Document doc, List<ElementInfo> allElements, List<GeometryData> geometryData)
    {
        var clashes = new List<ClashInfo>();
        var progressForm = new ClashDetectionProgressForm();
        bool isCancelled = false;
        
        try
        {
            progressForm.Show();
            Application.DoEvents();
            
            // Build spatial index
            progressForm.UpdateProgress(0, 100, "Building spatial index...", "");
            var spatialGrid = new SpatialGrid();
            
            foreach (var geomData in geometryData)
            {
                spatialGrid.AddElement(geomData);
            }
            
            // Get potential clash pairs from spatial index
            var potentialPairs = spatialGrid.GetPotentialClashPairs();
            int totalComparisons = potentialPairs.Count;
            
            progressForm.UpdateProgress(0, totalComparisons, 
                $"Checking {totalComparisons} potential clashes (reduced from {(allElements.Count * (allElements.Count - 1)) / 2})...", "");
            
            int currentComparison = 0;
            
            foreach (var (geom1, geom2) in potentialPairs)
            {
                currentComparison++;
                
                // Update UI periodically
                if (currentComparison % 5 == 0 || clashes.Count > 0)
                {
                    Application.DoEvents();
                    Thread.Yield();
                    
                    if (progressForm.IsCancelled)
                    {
                        isCancelled = true;
                        break;
                    }
                }
                
                var elem1Info = geom1.ElementInfo;
                var elem2Info = geom2.ElementInfo;
                
                string currentDesc = $"{GetElementDescription(elem1Info.Element, false)} vs {GetElementDescription(elem2Info.Element, false)}";
                progressForm.UpdateProgress(currentComparison, totalComparisons,
                    $"Checking potential clash {currentComparison} of {totalComparisons}", 
                    currentDesc);
                
                // Detailed solid intersection check
                if (geom1.Solid != null && geom2.Solid != null && 
                    CheckElementsClash(geom1.Solid, geom2.Solid))
                {
                    var clashInfo = new ClashInfo
                    {
                        Element1 = elem1Info.Element,
                        Element2 = elem2Info.Element,
                        Link1 = elem1Info.LinkInstance,
                        Link2 = elem2Info.LinkInstance,
                        ClashPoint = CalculateClashPoint(geom1.BoundingBox, geom2.BoundingBox),
                        OverlapBBox = CalculateOverlapBBox(geom1.BoundingBox, geom2.BoundingBox)
                    };
                    
                    clashes.Add(clashInfo);
                    
                    // Report clash immediately
                    string source1 = elem1Info.IsLinked ? elem1Info.SourceName : "Host";
                    string source2 = elem2Info.IsLinked ? elem2Info.SourceName : "Host";
                    
                    string clashDesc = $"Clash {clashes.Count}: [{source1}] {GetElementDescription(elem1Info.Element, false)} " +
                                      $"<-> [{source2}] {GetElementDescription(elem2Info.Element, false)}";
                    progressForm.AddClash(clashDesc);
                }
            }
            
            progressForm.SetComplete(clashes.Count, _stopwatch.ElapsedMilliseconds, isCancelled);
        }
        finally
        {
            if (progressForm != null && !progressForm.IsDisposed && !isCancelled)
            {
                // Keep form open if not cancelled
            }
        }
        
        return (clashes, isCancelled);
    }
    
    #endregion
    
    #region Parallel Subprocess Detection
    
    private (List<ClashInfo> Clashes, bool IsCancelled) PerformParallelClashDetection(
        Document doc, List<ElementInfo> allElements, List<GeometryData> geometryData)
    {
        var clashes = new List<ClashInfo>();
        var progressForm = new ClashDetectionProgressForm();
        bool isCancelled = false;
        
        try
        {
            progressForm.Show();
            Application.DoEvents();
            
            // Prepare data for external processing
            var clashCheckData = new ClashCheckData
            {
                Elements = geometryData.Select(g => new ExternalGeometryData
                {
                    ElementId = g.ElementId.IntegerValue,
                    BoundingBox = new[] 
                    { 
                        g.BoundingBox.Min.X, g.BoundingBox.Min.Y, g.BoundingBox.Min.Z,
                        g.BoundingBox.Max.X, g.BoundingBox.Max.Y, g.BoundingBox.Max.Z
                    },
                    SerializedSolid = g.SerializedSolid
                }).ToList()
            };
            
            // Save data to temp file
            string tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, JsonConvert.SerializeObject(clashCheckData));
            
            // Start subprocess (would need separate executable)
            progressForm.UpdateProgress(0, 100, "Starting parallel clash detection...", "");
            
            // This is where you'd spawn subprocesses
            // For demo purposes, showing the structure:
            int numProcesses = Environment.ProcessorCount;
            var tasks = new Task<List<ClashResult>>[numProcesses];
            
            for (int i = 0; i < numProcesses; i++)
            {
                int processIndex = i;
                tasks[i] = Task.Run(() => 
                {
                    // In real implementation, this would start a separate process
                    // Process.Start("ClashDetector.exe", $"{tempFile} {processIndex} {numProcesses}");
                    // For now, just return empty results
                    return new List<ClashResult>();
                });
            }
            
            // Wait for all tasks and collect results
            Task.WaitAll(tasks);
            
            // Convert results back to ClashInfo
            foreach (var task in tasks)
            {
                foreach (var result in task.Result)
                {
                    var elem1 = allElements.First(e => e.Element.Id.IntegerValue == result.ElementId1);
                    var elem2 = allElements.First(e => e.Element.Id.IntegerValue == result.ElementId2);
                    
                    clashes.Add(new ClashInfo
                    {
                        Element1 = elem1.Element,
                        Element2 = elem2.Element,
                        Link1 = elem1.LinkInstance,
                        Link2 = elem2.LinkInstance,
                        ClashPoint = new XYZ(result.ClashPoint[0], result.ClashPoint[1], result.ClashPoint[2]),
                        OverlapBBox = null // Would need to reconstruct
                    });
                }
            }
            
            // Clean up
            File.Delete(tempFile);
            
            progressForm.SetComplete(clashes.Count, _stopwatch.ElapsedMilliseconds, isCancelled);
        }
        finally
        {
            if (progressForm != null && !progressForm.IsDisposed && !isCancelled)
            {
                // Keep form open
            }
        }
        
        return (clashes, isCancelled);
    }
    
    #endregion
    
    #region Original Methods (unchanged)
    
    private List<ElementInfo> GetSelectedElements(UIDocument uidoc, Document doc)
    {
        var elementInfos = new List<ElementInfo>();
        
        // Try references first
        var references = uidoc.GetReferences();
        if (references != null && references.Count > 0)
        {
            foreach (var reference in references)
            {
                if (reference.LinkedElementId != ElementId.InvalidElementId)
                {
                    // Linked element
                    var linkInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkInstance != null)
                    {
                        var linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            var linkedElement = linkedDoc.GetElement(reference.LinkedElementId);
                            if (linkedElement != null && HasSolidGeometry(linkedElement))
                            {
                                elementInfos.Add(new ElementInfo
                                {
                                    Element = linkedElement,
                                    Document = linkedDoc,
                                    Transform = linkInstance.GetTotalTransform(),
                                    SourceName = linkInstance.Name,
                                    IsLinked = true,
                                    LinkInstance = linkInstance
                                });
                            }
                        }
                    }
                }
                else
                {
                    // Regular element in host model
                    var element = doc.GetElement(reference);
                    if (element != null && HasSolidGeometry(element))
                    {
                        elementInfos.Add(new ElementInfo
                        {
                            Element = element,
                            Document = doc,
                            Transform = Transform.Identity,
                            SourceName = "Host Model",
                            IsLinked = false,
                            LinkInstance = null
                        });
                    }
                }
            }
        }
        
        // If no references, try element IDs
        if (elementInfos.Count == 0)
        {
            var selectedIds = uidoc.GetSelectionIds();
            foreach (var id in selectedIds)
            {
                var element = doc.GetElement(id);
                if (element != null && HasSolidGeometry(element))
                {
                    elementInfos.Add(new ElementInfo
                    {
                        Element = element,
                        Document = doc,
                        Transform = Transform.Identity,
                        SourceName = "Host Model",
                        IsLinked = false,
                        LinkInstance = null
                    });
                }
            }
        }
        
        return elementInfos;
    }
    
    private bool HasSolidGeometry(Element element)
    {
        if (element.Category == null)
            return false;
        return element.get_Geometry(new Options()) != null;
    }
    
    private bool CheckElementsClash(Solid solid1, Solid solid2)
    {
        if (solid1 == null || solid2 == null)
            return false;
        
        try
        {
            var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                solid1, solid2, BooleanOperationsType.Intersect);
            
            return intersection != null && intersection.Volume > CLASH_TOLERANCE;
        }
        catch
        {
            return false;
        }
    }
    
    private XYZ CalculateClashPoint(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        if (bb1 == null || bb2 == null)
            return XYZ.Zero;
        
        var minX = Math.Max(bb1.Min.X, bb2.Min.X);
        var minY = Math.Max(bb1.Min.Y, bb2.Min.Y);
        var minZ = Math.Max(bb1.Min.Z, bb2.Min.Z);
        
        var maxX = Math.Min(bb1.Max.X, bb2.Max.X);
        var maxY = Math.Min(bb1.Max.Y, bb2.Max.Y);
        var maxZ = Math.Min(bb1.Max.Z, bb2.Max.Z);
        
        return new XYZ(
            (minX + maxX) / 2,
            (minY + maxY) / 2,
            (minZ + maxZ) / 2);
    }
    
    private BoundingBoxXYZ CalculateOverlapBBox(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        if (bb1 == null || bb2 == null)
            return null;
        
        var minX = Math.Max(bb1.Min.X, bb2.Min.X);
        var minY = Math.Max(bb1.Min.Y, bb2.Min.Y);
        var minZ = Math.Max(bb1.Min.Z, bb2.Min.Z);
        
        var maxX = Math.Min(bb1.Max.X, bb2.Max.X);
        var maxY = Math.Min(bb1.Max.Y, bb2.Max.Y);
        var maxZ = Math.Min(bb1.Max.Z, bb2.Max.Z);
        
        if (minX >= maxX || minY >= maxY || minZ >= maxZ)
            return null;
        
        return new BoundingBoxXYZ
        {
            Min = new XYZ(minX, minY, minZ),
            Max = new XYZ(maxX, maxY, maxZ)
        };
    }
    
    private void CreateClashMarkers(Document doc, List<ClashInfo> clashes)
    {
        using (Transaction trans = new Transaction(doc, "Create Clash Markers"))
        {
            trans.Start();
            
            // Get or create clash subcategory
            var genModelCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel);
            Category clashSubCat = null;
            
            if (!genModelCat.SubCategories.Contains("Clashes"))
            {
                clashSubCat = doc.Settings.Categories.NewSubcategory(genModelCat, "Clashes");
                clashSubCat.LineColor = new RevitColor(255, 0, 0); // Red
            }
            else
            {
                clashSubCat = genModelCat.SubCategories.get_Item("Clashes");
            }
            
            // Get or create transparent material
            Material clashMat = new FilteredElementCollector(doc)
                .OfClass(typeof(Material))
                .Cast<Material>()
                .FirstOrDefault(m => m.Name == "ClashTransparentRed");
            
            if (clashMat == null)
            {
                ElementId matId = Material.Create(doc, "ClashTransparentRed");
                clashMat = doc.GetElement(matId) as Material;
                clashMat.Color = new RevitColor(255, 0, 0);
                clashMat.Transparency = 70;
            }
            
            int clashNumber = 1;
            foreach (var clash in clashes)
            {
                if (clash.OverlapBBox != null)
                {
                    CreateClashDirectShape(doc, clash, clashNumber++, clashSubCat.Id, clashMat.Id);
                }
            }
            
            trans.Commit();
        }
    }
    
    private void CreateClashDirectShape(Document doc, ClashInfo clash, int clashNumber, ElementId subCategoryId, ElementId materialId)
    {
        var bbox = clash.OverlapBBox;
        
        // Apply offset
        var min = bbox.Min - new XYZ(BBOX_OFFSET, BBOX_OFFSET, BBOX_OFFSET);
        var max = bbox.Max + new XYZ(BBOX_OFFSET, BBOX_OFFSET, BBOX_OFFSET);
        
        // Create box geometry
        var bottomPts = new List<XYZ>
        {
            new XYZ(min.X, min.Y, min.Z),
            new XYZ(max.X, min.Y, min.Z),
            new XYZ(max.X, max.Y, min.Z),
            new XYZ(min.X, max.Y, min.Z)
        };
        
        var topPts = new List<XYZ>
        {
            new XYZ(min.X, min.Y, max.Z),
            new XYZ(max.X, min.Y, max.Z),
            new XYZ(max.X, max.Y, max.Z),
            new XYZ(min.X, max.Y, max.Z)
        };
        
        var builder = new TessellatedShapeBuilder();
        builder.OpenConnectedFaceSet(true);
        
        // Bottom face
        builder.AddFace(new TessellatedFace(bottomPts, materialId));
        
        // Top face
        builder.AddFace(new TessellatedFace(topPts, materialId));
        
        // Side faces
        for (int i = 0; i < 4; i++)
        {
            var j = (i + 1) % 4;
            var facePts = new List<XYZ> 
            { 
                bottomPts[i], 
                bottomPts[j], 
                topPts[j], 
                topPts[i] 
            };
            builder.AddFace(new TessellatedFace(facePts, materialId));
        }
        
        builder.CloseConnectedFaceSet();
        builder.Build();
        
        var result = builder.GetBuildResult();
        
        // Create DirectShape
        var ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
        ds.SetShape(result.GetGeometricalObjects());
        ds.Name = $"{CLASH_DS_PREFIX}{clashNumber:D4}";
        
        // Set subcategory
        var param = ds.get_Parameter(BuiltInParameter.FAMILY_ELEM_SUBCATEGORY);
        if (param != null && !param.IsReadOnly)
        {
            param.Set(subCategoryId);
        }
        
        // Add clash information to comments
        var clashInfo = $"Clash #{clashNumber}: {GetElementDescription(clash.Element1, true)}, {GetElementDescription(clash.Element2, true)}";
        
        var commentsParam = ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        if (commentsParam != null && !commentsParam.IsReadOnly)
        {
            commentsParam.Set(clashInfo);
        }
    }
    
    private string GetElementDescription(Element elem, bool detailed = true)
    {
        var category = elem.Category?.Name ?? "Unknown";
        var family = "";
        var type = "";
        
        if (elem is FamilyInstance fi)
        {
            family = fi.Symbol?.Family?.Name ?? "";
            type = fi.Symbol?.Name ?? "";
        }
        else
        {
            type = elem.Name;
        }
        
        if (!detailed)
        {
            if (!string.IsNullOrEmpty(type))
                return $"{category}: {type}";
            return category;
        }
        
        var desc = $"{category}";
        if (!string.IsNullOrEmpty(family))
            desc += $" - {family}";
        if (!string.IsNullOrEmpty(type))
            desc += $" - {type}";
        
        desc += $" [Id: {elem.Id.IntegerValue}]";
        
        return desc;
    }
    
    private BoundingBoxXYZ GetTransformedBoundingBox(Element elem, Transform transform)
    {
        var bb = elem.get_BoundingBox(null);
        if (bb == null) return null;
        
        var transformedBB = new BoundingBoxXYZ();
        transformedBB.Min = transform.OfPoint(bb.Min);
        transformedBB.Max = transform.OfPoint(bb.Max);
        
        // Ensure min/max are properly ordered
        var min = new XYZ(
            Math.Min(transformedBB.Min.X, transformedBB.Max.X),
            Math.Min(transformedBB.Min.Y, transformedBB.Max.Y),
            Math.Min(transformedBB.Min.Z, transformedBB.Max.Z));
        
        var max = new XYZ(
            Math.Max(transformedBB.Min.X, transformedBB.Max.X),
            Math.Max(transformedBB.Min.Y, transformedBB.Max.Y),
            Math.Max(transformedBB.Min.Z, transformedBB.Max.Z));
        
        transformedBB.Min = min;
        transformedBB.Max = max;
        
        return transformedBB;
    }
    
    private Solid GetSolid(Element elem)
    {
        var options = new Options
        {
            ComputeReferences = false,
            DetailLevel = ViewDetailLevel.Coarse
        };
        
        var geomElem = elem.get_Geometry(options);
        if (geomElem == null) return null;
        
        Solid largestSolid = null;
        double largestVolume = 0;
        
        foreach (var geomObj in geomElem)
        {
            var solid = geomObj as Solid;
            if (solid != null && solid.Volume > MIN_ELEMENT_SIZE && solid.Volume > largestVolume)
            {
                largestSolid = solid;
                largestVolume = solid.Volume;
            }
            
            var instance = geomObj as GeometryInstance;
            if (instance != null)
            {
                var instGeom = instance.GetInstanceGeometry();
                foreach (var instObj in instGeom)
                {
                    solid = instObj as Solid;
                    if (solid != null && solid.Volume > MIN_ELEMENT_SIZE && solid.Volume > largestVolume)
                    {
                        largestSolid = solid;
                        largestVolume = solid.Volume;
                    }
                }
            }
        }
        
        return largestSolid;
    }
    
    #endregion
    
    #region Helper Classes
    
    private class ElementInfo
    {
        public Element Element { get; set; }
        public Document Document { get; set; }
        public Transform Transform { get; set; }
        public string SourceName { get; set; }
        public bool IsLinked { get; set; }
        public RevitLinkInstance LinkInstance { get; set; }
    }
    
    private class ClashInfo
    {
        public Element Element1 { get; set; }
        public Element Element2 { get; set; }
        public RevitLinkInstance Link1 { get; set; }
        public RevitLinkInstance Link2 { get; set; }
        public XYZ ClashPoint { get; set; }
        public BoundingBoxXYZ OverlapBBox { get; set; }
    }
    
    private class GeometryData
    {
        public ElementId ElementId { get; set; }
        public ElementInfo ElementInfo { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
        public BoundingSphere BoundingSphere { get; set; }
        public Solid Solid { get; set; }
        public SerializedSolid SerializedSolid { get; set; }
    }
    
    private class BoundingSphere
    {
        public XYZ Center { get; set; }
        public double Radius { get; set; }
    }
    
    // Serializable classes for external processing
    [Serializable]
    private class SerializedSolid
    {
        public List<SerializedFace> Faces { get; set; }
        public double Volume { get; set; }
    }
    
    [Serializable]
    private class SerializedFace
    {
        public List<double[]> Vertices { get; set; }
        public List<int[]> Triangles { get; set; }
    }
    
    private class ClashCheckData
    {
        public List<ExternalGeometryData> Elements { get; set; }
    }
    
    private class ExternalGeometryData
    {
        public int ElementId { get; set; }
        public double[] BoundingBox { get; set; }
        public SerializedSolid SerializedSolid { get; set; }
    }
    
    private class ClashResult
    {
        public int ElementId1 { get; set; }
        public int ElementId2 { get; set; }
        public double[] ClashPoint { get; set; }
    }
    
    #endregion
}

// Progress form remains the same
public class ClashDetectionProgressForm : WinForm
{
    private ProgressBar progressBar;
    private Label statusLabel;
    private Label currentElementLabel;
    private ListBox clashListBox;
    private Button cancelButton;
    private Label clashCountLabel;
    private Label timeLabel;
    private Stopwatch stopwatch;
    private System.Windows.Forms.Timer updateTimer;
    
    public bool IsCancelled { get; private set; }
    
    public ClashDetectionProgressForm()
    {
        InitializeComponent();
        stopwatch = new Stopwatch();
        stopwatch.Start();
        
        updateTimer = new System.Windows.Forms.Timer();
        updateTimer.Interval = 100;
        updateTimer.Tick += (s, e) => UpdateElapsedTime();
        updateTimer.Start();
    }
    
    private void InitializeComponent()
    {
        this.Text = "Clash Detection Progress";
        this.Size = new WinSize(600, 520);
        this.MinimumSize = new WinSize(400, 300);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        
        progressBar = new ProgressBar
        {
            Location = new WinPoint(12, 12),
            Size = new WinSize(this.ClientSize.Width - 24, 23),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Style = ProgressBarStyle.Continuous
        };
        
        statusLabel = new Label
        {
            Location = new WinPoint(12, 40),
            Size = new WinSize(this.ClientSize.Width - 24, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "Initializing..."
        };
        
        currentElementLabel = new Label
        {
            Location = new WinPoint(12, 65),
            Size = new WinSize(this.ClientSize.Width - 24, 40),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "Processing: ",
            AutoSize = false
        };
        
        timeLabel = new Label
        {
            Location = new WinPoint(this.ClientSize.Width - 172 - 12, 110),
            Size = new WinSize(172, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Text = "Elapsed: 00:00",
            TextAlign = ContentAlignment.TopRight
        };
        
        clashCountLabel = new Label
        {
            Location = new WinPoint(12, 110),
            Size = new WinSize(200, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Left,
            Text = "Clashes found: 0",
            Font = new WinFont(this.Font, FontStyle.Bold)
        };
        
        clashListBox = new ListBox
        {
            Location = new WinPoint(12, 135),
            Size = new WinSize(this.ClientSize.Width - 24, this.ClientSize.Height - 135 - 35 - 12),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            HorizontalScrollbar = true,
            Font = new WinFont("Consolas", 9),
            SelectionMode = SelectionMode.MultiExtended
        };
        
        cancelButton = new Button
        {
            Location = new WinPoint(this.ClientSize.Width - 75 - 12, this.ClientSize.Height - 23 - 12),
            Size = new WinSize(75, 23),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Text = "Cancel",
            DialogResult = DialogResult.Cancel
        };
        cancelButton.Click += (sender, e) => 
        {
            IsCancelled = true;
            this.DialogResult = DialogResult.Cancel;
        };
        
        this.Controls.AddRange(new WinControl[] 
        {
            progressBar,
            statusLabel,
            currentElementLabel,
            timeLabel,
            clashCountLabel,
            clashListBox,
            cancelButton
        });
        
        this.KeyPreview = true;
        this.KeyDown += new KeyEventHandler(Form_KeyDown);
    }
    
    private void Form_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.C)
        {
            if (clashListBox.SelectedItems.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                foreach (var item in clashListBox.SelectedItems)
                {
                    sb.AppendLine(item.ToString());
                }
                Clipboard.SetText(sb.ToString());
            }
        }
    }
    
    private void UpdateElapsedTime()
    {
        if (!this.IsDisposed)
        {
            var elapsed = stopwatch.Elapsed;
            timeLabel.Text = $"Elapsed: {elapsed:mm\\:ss}";
        }
    }
    
    public void UpdateProgress(int current, int total, string status, string currentElement)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => UpdateProgress(current, total, status, currentElement)));
            return;
        }
        
        progressBar.Maximum = total;
        progressBar.Value = Math.Min(current, total);
        statusLabel.Text = status;
        currentElementLabel.Text = $"Processing: {currentElement}";
        
        int percentage = total > 0 ? (current * 100 / total) : 0;
        this.Text = $"Clash Detection Progress - {percentage}%";
        
        if (current > 0 && stopwatch.ElapsedMilliseconds > 1000)
        {
            var avgTimePerItem = stopwatch.ElapsedMilliseconds / (double)current;
            var remainingItems = total - current;
            var estimatedMs = avgTimePerItem * remainingItems;
            var estimatedTime = TimeSpan.FromMilliseconds(estimatedMs);
            
            if (estimatedTime.TotalSeconds > 5)
            {
                statusLabel.Text += $" (Est. remaining: {estimatedTime:mm\\:ss})";
            }
        }
    }
    
    public void AddClash(string clashDescription)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => AddClash(clashDescription)));
            return;
        }
        
        string timestamp = DateTime.Now.ToString("HH:mm:ss");
        clashListBox.Items.Add($"[{timestamp}] {clashDescription}");
        clashListBox.TopIndex = clashListBox.Items.Count - 1;
        clashCountLabel.Text = $"Clashes found: {clashListBox.Items.Count}";
        
        var originalColor = clashCountLabel.ForeColor;
        clashCountLabel.ForeColor = System.Drawing.Color.Red;
        Application.DoEvents();
        System.Threading.Tasks.Task.Delay(100).ContinueWith(t => 
        {
            if (!this.IsDisposed)
            {
                this.Invoke(new Action(() => clashCountLabel.ForeColor = originalColor));
            }
        });
    }
    
    public void SetComplete(int totalClashes, long elapsedMs, bool isCancelled = false)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => SetComplete(totalClashes, elapsedMs, isCancelled)));
            return;
        }
        
        stopwatch.Stop();
        updateTimer.Stop();
        
        progressBar.Value = progressBar.Maximum;
        if (isCancelled)
        {
            this.Text = "Clash Detection Cancelled";
            statusLabel.Text = $"Cancelled. Found {totalClashes} clashes in {elapsedMs}ms";
        }
        else
        {
            statusLabel.Text = $"Complete! Found {totalClashes} clashes in {elapsedMs}ms";
        }
        cancelButton.Text = "Close";
        IsCancelled = false;
        
        this.Controls.Remove(cancelButton);
        cancelButton = new Button
        {
            Location = new WinPoint(this.ClientSize.Width - 75 - 12, this.ClientSize.Height - 23 - 12),
            Size = new WinSize(75, 23),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Text = "Close"
        };
        cancelButton.Click += (sender, e) => { this.Close(); };
        this.Controls.Add(cancelButton);
    }
    
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            updateTimer?.Stop();
            updateTimer?.Dispose();
            stopwatch?.Stop();
        }
        base.Dispose(disposing);
    }
}
