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

// Aliases to resolve conflicts between Revit API and Windows Forms namespaces
using WinForm = System.Windows.Forms.Form;
using WinPoint = System.Drawing.Point;
using WinSize = System.Drawing.Size;
using WinColor = System.Drawing.Color;
using WinFont = System.Drawing.Font;
using WinControl = System.Windows.Forms.Control;
using RevitColor = Autodesk.Revit.DB.Color;

/// <summary>
/// ClashDetectionSelectedElements - Single file implementation
/// Detects clashes between currently selected elements (from project or linked models)
/// Creates DirectShape markers at clash locations with detailed information
/// </summary>
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ClashDetectionSelectedElements : IExternalCommand
{
    // Configuration
    private const double CLASH_TOLERANCE = 0.001; // 1mm tolerance
    private const string CLASH_DS_PREFIX = "CLASH_AREA_";
    private const double MIN_ELEMENT_SIZE = 0.1; // Minimum element size in feet to consider
    private const double BBOX_OFFSET = 0.5; // Offset for bounding box in feet
    
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
            
            // Get selected elements using SelectionModeManager
            var selectedElements = GetSelectedElements(uidoc, doc);
            
            if (selectedElements.Count < 2)
            {
                TaskDialog.Show("Clash Detection", 
                    $"Please select at least 2 elements for clash detection.\n" +
                    $"Currently selected: {selectedElements.Count} element(s).");
                return Result.Cancelled;
            }
            
            // Group elements by their source (document or link)
            var elementGroups = GroupElementsBySource(selectedElements);
            
            if (elementGroups.Count < 2 && elementGroups.First().Value.Count < 2)
            {
                TaskDialog.Show("Clash Detection", 
                    "Selected elements must be from at least 2 different sources (models/links) " +
                    "or select multiple elements from the same model to check for self-clashes.");
                return Result.Cancelled;
            }
            
            // Perform clash detection
            var (clashes, isCancelled) = PerformClashDetection(doc, selectedElements, elementGroups);
            
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
        // Quick check for element types that typically have solid geometry
        if (element.Category == null)
            return false;
            
        // Check if it's a type that typically has geometry
        return element.get_Geometry(new Options()) != null;
    }
    
    private Dictionary<string, List<ElementInfo>> GroupElementsBySource(List<ElementInfo> elements)
    {
        return elements.GroupBy(e => e.IsLinked ? e.LinkInstance.UniqueId : "HOST")
                      .ToDictionary(g => g.Key, g => g.ToList());
    }
    
    private (List<ClashInfo> Clashes, bool IsCancelled) PerformClashDetection(Document doc, List<ElementInfo> allElements, 
                                                  Dictionary<string, List<ElementInfo>> elementGroups)
    {
        var clashes = new List<ClashInfo>();
        var progressForm = new ClashDetectionProgressForm();
        bool isCancelled = false;
        
        try
        {
            progressForm.Show();
            Application.DoEvents();
            
            // Precompute transformed bounding boxes and solids for all elements (optimization to avoid repeated API calls in loop)
            var elementData = new Dictionary<ElementId, (BoundingBoxXYZ BB, Solid Solid)>();
            foreach (var elemInfo in allElements)
            {
                var bb = GetTransformedBoundingBox(elemInfo.Element, elemInfo.Transform);
                var solid = GetSolid(elemInfo.Element);
                if (solid != null && !elemInfo.Transform.IsIdentity)
                {
                    solid = SolidUtils.CreateTransformed(solid, elemInfo.Transform);
                }
                elementData[elemInfo.Element.Id] = (bb, solid);
            }
            
            // Calculate total comparisons - always all unique pairs
            int totalComparisons = (allElements.Count * (allElements.Count - 1)) / 2;
            
            progressForm.UpdateProgress(0, totalComparisons, 
                $"Checking {allElements.Count} elements for clashes...", "");
            
            int currentComparison = 0;
            
            // Check each pair of elements
            for (int i = 0; i < allElements.Count - 1; i++)
            {
                var elem1Info = allElements[i];
                var (bb1, solid1) = elementData[elem1Info.Element.Id];
                
                for (int j = i + 1; j < allElements.Count; j++)
                {
                    var elem2Info = allElements[j];
                    var (bb2, solid2) = elementData[elem2Info.Element.Id];
                    
                    // Skip if same element
                    if (elem1Info.Element.Id == elem2Info.Element.Id && 
                        elem1Info.Document.Title == elem2Info.Document.Title)
                        continue;
                    
                    currentComparison++;
                    
                    // Update UI more frequently for better responsiveness
                    if (currentComparison % 5 == 0 || clashes.Count > 0)
                    {
                        Application.DoEvents();
                        Thread.Yield(); // Allow other threads to run
                        
                        if (progressForm.IsCancelled)
                        {
                            isCancelled = true;
                            goto DetectionEnd;
                        }
                    }
                    
                    string currentDesc = $"{GetElementDescription(elem1Info.Element, false)} [ID: {elem1Info.Element.Id}] vs {GetElementDescription(elem2Info.Element, false)} [ID: {elem2Info.Element.Id}]";
                    progressForm.UpdateProgress(currentComparison, totalComparisons,
                        $"Checking element {i+1} vs {j+1} of {allElements.Count}",
                        currentDesc);
                    
                    // Quick bounding box check first
                    if (!DoBoundingBoxesOverlap(bb1, bb2))
                        continue;
                    
                    // Detailed solid intersection check
                    if (CheckElementsClash(solid1, solid2))
                    {
                        var clashInfo = new ClashInfo
                        {
                            Element1 = elem1Info.Element,
                            Element2 = elem2Info.Element,
                            Link1 = elem1Info.LinkInstance,
                            Link2 = elem2Info.LinkInstance,
                            ClashPoint = CalculateClashPoint(bb1, bb2),
                            OverlapBBox = CalculateOverlapBBox(bb1, bb2)
                        };
                        
                        clashes.Add(clashInfo);
                        
                        // Report clash immediately
                        string source1 = elem1Info.IsLinked ? elem1Info.SourceName : "Host";
                        string source2 = elem2Info.IsLinked ? elem2Info.SourceName : "Host";
                        
                        string clashDesc = $"Clash {clashes.Count}: [{source1}] {GetElementDescription(elem1Info.Element, false)} " +
                                          $"<-> [{source2}] {GetElementDescription(elem2Info.Element, false)}";
                        progressForm.AddClash(clashDesc);
                        
                        Application.DoEvents();
                    }
                }
            }
            
DetectionEnd:
            progressForm.SetComplete(clashes.Count, _stopwatch.ElapsedMilliseconds, isCancelled);
        }
        finally
        {
            if (progressForm != null && !progressForm.IsDisposed)
            {
                if (isCancelled)
                {
                    progressForm.Close();
                }
            }
        }
        
        return (clashes, isCancelled);
    }
    
    private bool DoBoundingBoxesOverlap(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
    {
        if (bb1 == null || bb2 == null)
            return false;
        
        return bb1.Min.X <= bb2.Max.X && bb1.Max.X >= bb2.Min.X &&
               bb1.Min.Y <= bb2.Max.Y && bb1.Max.Y >= bb2.Min.Y &&
               bb1.Min.Z <= bb2.Max.Z && bb1.Max.Z >= bb2.Min.Z;
    }
    
    private bool CheckElementsClash(Solid solid1, Solid solid2)
    {
        if (solid1 == null || solid2 == null)
            return false;
        
        // Check intersection
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
        
        // Calculate intersection center
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
                clashMat.Transparency = 70; // 0-100 scale
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
        
        // Set comments parameter
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
            // Short version for progress display
            if (!string.IsNullOrEmpty(type))
                return $"{category}: {type}";
            return category;
        }
        
        // Detailed version for clash report
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
        
        // Get largest solid (skip very small solids)
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
            
            // Check geometry instances
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
    
    // Helper class to store element information with transform
    private class ElementInfo
    {
        public Element Element { get; set; }
        public Document Document { get; set; }
        public Transform Transform { get; set; }
        public string SourceName { get; set; }
        public bool IsLinked { get; set; }
        public RevitLinkInstance LinkInstance { get; set; }
    }
    
    // Helper class to store clash information
    private class ClashInfo
    {
        public Element Element1 { get; set; }
        public Element Element2 { get; set; }
        public RevitLinkInstance Link1 { get; set; }
        public RevitLinkInstance Link2 { get; set; }
        public XYZ ClashPoint { get; set; }
        public BoundingBoxXYZ OverlapBBox { get; set; }
    }
}

/// <summary>
/// Progress dialog for clash detection operations
/// Shows real-time progress, found clashes, elapsed time, and provides cancellation
/// </summary>
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
        
        // Timer to update elapsed time
        updateTimer = new System.Windows.Forms.Timer();
        updateTimer.Interval = 100; // Update every 100ms
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
        
        // Progress bar
        progressBar = new ProgressBar
        {
            Location = new WinPoint(12, 12),
            Size = new WinSize(this.ClientSize.Width - 24, 23),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Style = ProgressBarStyle.Continuous
        };
        
        // Status label
        statusLabel = new Label
        {
            Location = new WinPoint(12, 40),
            Size = new WinSize(this.ClientSize.Width - 24, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "Initializing..."
        };
        
        // Current element label
        currentElementLabel = new Label
        {
            Location = new WinPoint(12, 65),
            Size = new WinSize(this.ClientSize.Width - 24, 40),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "Processing: ",
            AutoSize = false
        };
        
        // Time label
        timeLabel = new Label
        {
            Location = new WinPoint(this.ClientSize.Width - 172 - 12, 110),
            Size = new WinSize(172, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Text = "Elapsed: 00:00",
            TextAlign = ContentAlignment.TopRight
        };
        
        // Clash count label
        clashCountLabel = new Label
        {
            Location = new WinPoint(12, 110),
            Size = new WinSize(200, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Left,
            Text = "Clashes found: 0",
            Font = new WinFont(this.Font, FontStyle.Bold)
        };
        
        // Clash list box
        clashListBox = new ListBox
        {
            Location = new WinPoint(12, 135),
            Size = new WinSize(this.ClientSize.Width - 24, this.ClientSize.Height - 135 - 35 - 12),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            HorizontalScrollbar = true,
            Font = new WinFont("Consolas", 9),
            SelectionMode = SelectionMode.MultiExtended
        };
        
        // Cancel button
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
        
        // Add controls
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
        
        // Update percentage in title
        int percentage = total > 0 ? (current * 100 / total) : 0;
        this.Text = $"Clash Detection Progress - {percentage}%";
        
        // Estimate time remaining
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
        clashListBox.TopIndex = clashListBox.Items.Count - 1; // Scroll to latest
        clashCountLabel.Text = $"Clashes found: {clashListBox.Items.Count}";
        
        // Flash the clash count label briefly
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
        IsCancelled = false; // Reset the flag
        
        // Recreate button to clear all handlers
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
