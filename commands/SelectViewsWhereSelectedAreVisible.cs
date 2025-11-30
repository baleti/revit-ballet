using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using WinForms = System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class SelectViewsWhereSelectedAreVisible : IExternalCommand
{
    // View info class for DataGrid display
    public class ViewInfo
    {
        public string ViewName { get; set; }
        public string ViewType { get; set; }
        public string Level { get; set; }
        public string Sheet { get; set; }
        public string SheetNumber { get; set; }
        public string Scale { get; set; }
        public int ElementsFound { get; set; }
        public int ViewId { get; set; }
        public Autodesk.Revit.DB.View ViewObject { get; set; }
    }
    
    // Progress dialog with live DataGrid results
    private class ProgressDialog : WinForms.Form
    {
        private WinForms.Label statusLabel;
        private WinForms.ProgressBar progressBar;
        private WinForms.Button cancelButton;
        private WinForms.Button selectButton;
        private WinForms.DataGridView dataGrid;
        private WinForms.Label countLabel;
        private System.Windows.Forms.BindingSource bindingSource;
        private volatile bool cancelRequested = false;
        
        public List<ViewInfo> FoundViews { get; private set; }
        public List<ViewInfo> SelectedViews { get; private set; }
        
        private Document doc;
        private List<ElementId> selectedElementIds;
        
        public ProgressDialog(Document document, List<ElementId> elementIds)
        {
            doc = document;
            selectedElementIds = elementIds;
            FoundViews = new List<ViewInfo>();
            SelectedViews = new List<ViewInfo>();
            
            InitializeComponents();
        }
        
        private void InitializeComponents()
        {
            // Form settings
            this.Text = "Finding Views with Selected Elements";
            this.Width = 900;
            this.Height = 600;
            this.FormBorderStyle = WinForms.FormBorderStyle.Sizable;
            this.MinimumSize = new System.Drawing.Size(800, 400);
            this.StartPosition = WinForms.FormStartPosition.CenterScreen;
            
            // Enable ESC key handling
            this.KeyPreview = true;
            this.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == WinForms.Keys.Escape && !cancelRequested)
                {
                    cancelRequested = true;
                    cancelButton.Text = "Cancelling...";
                    cancelButton.Enabled = false;
                    cancelButton.Refresh();
                }
            };
            
            // Status label
            statusLabel = new WinForms.Label
            {
                Text = $"Searching views for {selectedElementIds.Count} selected element(s)...",
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(860, 20),
                AutoSize = false,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            
            // Progress bar
            progressBar = new WinForms.ProgressBar
            {
                Location = new System.Drawing.Point(10, 35),
                Size = new System.Drawing.Size(860, 23),
                Minimum = 0,
                Maximum = 100, // Will be updated when views are counted
                Style = WinForms.ProgressBarStyle.Continuous,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            
            // Count label
            countLabel = new WinForms.Label
            {
                Text = "Found: 0 views",
                Location = new System.Drawing.Point(10, 65),
                Size = new System.Drawing.Size(200, 20),
                AutoSize = false
            };
            
            // DataGrid for results
            dataGrid = new WinForms.DataGridView
            {
                Location = new System.Drawing.Point(10, 90),
                Size = new System.Drawing.Size(860, 430),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | 
                         WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToOrderColumns = true,
                SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                AutoGenerateColumns = true,
                ReadOnly = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.Fill
            };
            
            // Setup data binding
            bindingSource = new WinForms.BindingSource();
            bindingSource.DataSource = FoundViews;
            dataGrid.DataSource = bindingSource;
            
            // Cancel button
            cancelButton = new WinForms.Button
            {
                Text = "Cancel Search",
                Location = new System.Drawing.Point(10, 530),
                Size = new System.Drawing.Size(120, 30),
                Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left
            };
            cancelButton.Click += (sender, e) =>
            {
                if (!cancelRequested)
                {
                    cancelRequested = true;
                    cancelButton.Text = "Cancelling...";
                    cancelButton.Enabled = false;
                    cancelButton.Refresh();
                    WinForms.Application.DoEvents();
                }
            };
            
            // Select button
            selectButton = new WinForms.Button
            {
                Text = "Select Views",
                Location = new System.Drawing.Point(750, 530),
                Size = new System.Drawing.Size(120, 30),
                Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right,
                Enabled = false
            };
            selectButton.Click += (sender, e) =>
            {
                // Get selected rows
                SelectedViews.Clear();
                foreach (WinForms.DataGridViewRow row in dataGrid.SelectedRows)
                {
                    var viewInfo = row.DataBoundItem as ViewInfo;
                    if (viewInfo != null)
                    {
                        SelectedViews.Add(viewInfo);
                    }
                }
                
                // Cancel any ongoing search
                cancelRequested = true;
                
                this.DialogResult = WinForms.DialogResult.OK;
                this.Close();
            };
            
            // Configure DataGrid columns after binding
            dataGrid.DataBindingComplete += (sender, e) =>
            {
                if (dataGrid.Columns.Count > 0)
                {
                    // Hide the ViewObject column
                    if (dataGrid.Columns["ViewObject"] != null)
                        dataGrid.Columns["ViewObject"].Visible = false;
                    
                    // Set column widths
                    if (dataGrid.Columns["ViewName"] != null)
                        dataGrid.Columns["ViewName"].FillWeight = 200;
                    if (dataGrid.Columns["ViewType"] != null)
                        dataGrid.Columns["ViewType"].FillWeight = 100;
                    if (dataGrid.Columns["Level"] != null)
                        dataGrid.Columns["Level"].FillWeight = 80;
                    if (dataGrid.Columns["Sheet"] != null)
                        dataGrid.Columns["Sheet"].FillWeight = 150;
                    if (dataGrid.Columns["SheetNumber"] != null)
                        dataGrid.Columns["SheetNumber"].FillWeight = 80;
                    if (dataGrid.Columns["Scale"] != null)
                        dataGrid.Columns["Scale"].FillWeight = 60;
                    if (dataGrid.Columns["ElementsFound"] != null)
                    {
                        dataGrid.Columns["ElementsFound"].FillWeight = 60;
                        dataGrid.Columns["ElementsFound"].HeaderText = "Elements";
                    }
                    if (dataGrid.Columns["ViewId"] != null)
                        dataGrid.Columns["ViewId"].FillWeight = 60;
                }
            };
            
            // Enable select button when rows are selected
            dataGrid.SelectionChanged += (sender, e) =>
            {
                selectButton.Enabled = dataGrid.SelectedRows.Count > 0;
            };
            
            // Add controls
            this.Controls.AddRange(new WinForms.Control[] { 
                statusLabel, progressBar, countLabel, dataGrid, cancelButton, selectButton 
            });
        }
        
        private void SearchForViews()
        {
            int viewIndex = 0;
            int updateCounter = 0;
            
            try
            {
                // Update status
                statusLabel.Text = "Collecting views from document...";
                WinForms.Application.DoEvents();
                
                // Get all views
                var allViews = new FilteredElementCollector(doc)
                    .OfClass(typeof(Autodesk.Revit.DB.View))
                    .Cast<Autodesk.Revit.DB.View>()
                    .Where(v => !v.IsTemplate && IsGraphicalView(v))
                    .OrderBy(v => v.ViewType)
                    .ThenBy(v => v.Name)
                    .ToList();
                
                if (cancelRequested) return;
                
                // Set progress bar maximum
                progressBar.Maximum = allViews.Count;
                progressBar.Value = 0;
                
                // Build view to sheet mapping
                statusLabel.Text = "Building sheet references...";
                WinForms.Application.DoEvents();
                
                var viewSheetMap = BuildViewSheetMapping();
                if (cancelRequested) return;
                
                // Pre-compute bounding boxes for selected elements
                statusLabel.Text = "Analyzing selected elements...";
                WinForms.Application.DoEvents();
                
                var elementBoundingBoxes = new Dictionary<ElementId, BoundingBoxXYZ>();
                foreach (var elementId in selectedElementIds)
                {
                    if (cancelRequested) return;
                    
                    var element = doc.GetElement(elementId);
                    if (element != null)
                    {
                        var bbox = element.get_BoundingBox(null); // Get in model space
                        if (bbox != null)
                        {
                            elementBoundingBoxes[elementId] = bbox;
                        }
                    }
                }
                
                statusLabel.Text = $"Searching {allViews.Count} views for {selectedElementIds.Count} element(s)...";
                WinForms.Application.DoEvents();
                
                // Now search through views
                foreach (var view in allViews)
                {
                    // Check for cancellation
                    if (cancelRequested) break;
                    
                    viewIndex++;
                    progressBar.Value = viewIndex;
                    updateCounter++;
                    
                    var visibleElementIds = new List<ElementId>();
                    
                    try
                    {
                        // First, check if any selected elements could possibly be visible in this view
                        if (!CouldElementsBeVisibleInView(view, elementBoundingBoxes))
                        {
                            statusLabel.Text = $"Skipped: {view.Name} (bounds check)";
                            
                            // Update UI periodically
                            if (updateCounter % 20 == 0)
                            {
                                WinForms.Application.DoEvents();
                                System.Threading.Thread.Sleep(1);
                            }
                            continue;
                        }
                        
                        // Now do the expensive check - get ALL elements visible in this view
                        var viewCollector = new FilteredElementCollector(doc, view.Id);
                        var elementsInView = viewCollector.ToElementIds();
                        
                        // Check which selected elements are in this view
                        foreach (var selectedId in selectedElementIds)
                        {
                            if (elementsInView.Contains(selectedId))
                            {
                                visibleElementIds.Add(selectedId);
                            }
                        }
                        
                        if (visibleElementIds.Count > 0)
                        {
                            // Create ViewInfo object
                            var viewInfo = new ViewInfo
                            {
                                ViewName = view.Name,
                                ViewType = view.ViewType.ToString(),
                                Level = GetViewLevel(view),
                                Scale = GetViewScale(view),
                                ElementsFound = visibleElementIds.Count,
                                ViewId = (int)view.Id.AsLong(),
                                ViewObject = view
                            };
                            
                            // Get sheet info if view is on a sheet
                            if (viewSheetMap.ContainsKey(view.Id))
                            {
                                var sheet = viewSheetMap[view.Id];
                                viewInfo.Sheet = sheet.Name;
                                viewInfo.SheetNumber = sheet.SheetNumber;
                            }
                            else
                            {
                                viewInfo.Sheet = "Not on sheet";
                                viewInfo.SheetNumber = "-";
                            }
                            
                            // Add to found views
                            FoundViews.Add(viewInfo);
                            bindingSource.ResetBindings(false);
                            
                            // Update status
                            countLabel.Text = $"Found: {FoundViews.Count} view(s)";
                            statusLabel.Text = $"Found view: {viewInfo.ViewName} ({viewInfo.ElementsFound} elements visible)";
                            
                            // Auto-scroll to the last added item
                            if (dataGrid.Rows.Count > 0)
                            {
                                dataGrid.FirstDisplayedScrollingRowIndex = dataGrid.Rows.Count - 1;
                            }
                        }
                        else
                        {
                            statusLabel.Text = $"Checked: {view.Name}";
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log the error but continue
                        statusLabel.Text = $"Error checking: {view.Name} - {ex.Message}";
                    }
                    
                    // Update UI more frequently and ensure cancel button events are processed
                    if (updateCounter % 5 == 0 || visibleElementIds.Count > 0)
                    {
                        // Force processing of all pending Windows messages
                        WinForms.Application.DoEvents();
                        System.Threading.Thread.Sleep(1); // Give UI thread a chance to process
                        
                        // Double-check cancellation
                        if (cancelRequested) break;
                    }
                }
                
                // Search completed
                if (cancelRequested)
                {
                    statusLabel.Text = $"Search cancelled. Found {FoundViews.Count} view(s).";
                }
                else
                {
                    statusLabel.Text = $"Search completed. Found {FoundViews.Count} view(s) containing the selected elements.";
                }
                
                cancelButton.Text = "Close";
                cancelButton.Enabled = true;
                cancelButton.Click -= null;
                cancelButton.Click += (s, args) => this.Close();
                
                // Enable select button if we have results
                selectButton.Enabled = FoundViews.Count > 0 && dataGrid.SelectedRows.Count > 0;
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"Error: {ex.Message}";
                WinForms.MessageBox.Show($"Error during search: {ex.Message}", "Error", 
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                
                cancelButton.Text = "Close";
                cancelButton.Enabled = true;
                cancelButton.Click -= null;
                cancelButton.Click += (s, args) => this.Close();
            }
        }
        
        private bool CouldElementsBeVisibleInView(Autodesk.Revit.DB.View view, Dictionary<ElementId, BoundingBoxXYZ> elementBoundingBoxes)
        {
            try
            {
                // Skip checking for views without crop boxes or where crop is not active
                if (!view.CropBoxActive && !(view is View3D view3d && view3d.IsSectionBoxActive))
                    return true; // Assume elements might be visible if no crop box
                
                BoundingBoxXYZ viewBounds = null;
                
                // Handle 3D views with section boxes
                if (view is View3D view3D && view3D.IsSectionBoxActive)
                {
                    viewBounds = view3D.GetSectionBox();
                }
                // Handle plan views (floor plans, ceiling plans, etc.)
                else if (view.ViewType == ViewType.FloorPlan ||
                         view.ViewType == ViewType.CeilingPlan ||
                         view.ViewType == ViewType.EngineeringPlan ||
                         view.ViewType == ViewType.AreaPlan)
                {
                    // Get the crop box for X and Y extents
                    var cropBox = view.CropBox;
                    
                    // Get the X and Y extents from the crop box in world coordinates
                    var worldPointsXY = new List<XYZ>();
                    foreach (double x in new double[] { cropBox.Min.X, cropBox.Max.X })
                    {
                        foreach (double y in new double[] { cropBox.Min.Y, cropBox.Max.Y })
                        {
                            XYZ localPt = new XYZ(x, y, 0);
                            XYZ worldPt = cropBox.Transform.OfPoint(localPt);
                            worldPointsXY.Add(worldPt);
                        }
                    }
                    
                    double minX = worldPointsXY.Min(pt => pt.X);
                    double minY = worldPointsXY.Min(pt => pt.Y);
                    double maxX = worldPointsXY.Max(pt => pt.X);
                    double maxY = worldPointsXY.Max(pt => pt.Y);
                    
                    // Get the view range for Z extents
                    ViewPlan viewPlan = view as ViewPlan;
                    PlanViewRange viewRange = viewPlan.GetViewRange();
                    
                    // Get the levels associated with the view range
                    ElementId topLevelId = viewRange.GetLevelId(PlanViewPlane.TopClipPlane);
                    ElementId bottomLevelId = viewRange.GetLevelId(PlanViewPlane.BottomClipPlane);
                    ElementId viewDepthLevelId = viewRange.GetLevelId(PlanViewPlane.ViewDepthPlane);
                    
                    Level topLevel = doc.GetElement(topLevelId) as Level;
                    Level bottomLevel = doc.GetElement(bottomLevelId) as Level;
                    Level viewDepthLevel = doc.GetElement(viewDepthLevelId) as Level;
                    
                    // Get the offsets
                    double topOffset = viewRange.GetOffset(PlanViewPlane.TopClipPlane);
                    double bottomOffset = viewRange.GetOffset(PlanViewPlane.BottomClipPlane);
                    double viewDepthOffset = viewRange.GetOffset(PlanViewPlane.ViewDepthPlane);
                    
                    // Calculate actual elevations
                    double topElevation = topLevel.ProjectElevation + topOffset;
                    double bottomElevation = bottomLevel.ProjectElevation + bottomOffset;
                    double viewDepthElevation = viewDepthLevel.ProjectElevation + viewDepthOffset;
                    
                    // Use view depth as the bottom and top as the top (as shown in reference code)
                    double minZ = viewDepthElevation;
                    double maxZ = topElevation;
                    
                    viewBounds = new BoundingBoxXYZ
                    {
                        Min = new XYZ(minX, minY, minZ),
                        Max = new XYZ(maxX, maxY, maxZ),
                        Transform = Transform.Identity // World coordinates
                    };
                }
                // Handle sections, elevations, and other view types
                else if (view.CropBoxActive)
                {
                    var cropBox = view.CropBox;
                    
                    // For non-plan views, transform all 8 corners of the crop box
                    var worldPoints = new List<XYZ>();
                    foreach (double x in new double[] { cropBox.Min.X, cropBox.Max.X })
                    {
                        foreach (double y in new double[] { cropBox.Min.Y, cropBox.Max.Y })
                        {
                            foreach (double z in new double[] { cropBox.Min.Z, cropBox.Max.Z })
                            {
                                XYZ localPt = new XYZ(x, y, z);
                                XYZ worldPt = cropBox.Transform.OfPoint(localPt);
                                worldPoints.Add(worldPt);
                            }
                        }
                    }
                    
                    double minX = worldPoints.Min(pt => pt.X);
                    double minY = worldPoints.Min(pt => pt.Y);
                    double minZ = worldPoints.Min(pt => pt.Z);
                    double maxX = worldPoints.Max(pt => pt.X);
                    double maxY = worldPoints.Max(pt => pt.Y);
                    double maxZ = worldPoints.Max(pt => pt.Z);
                    
                    viewBounds = new BoundingBoxXYZ
                    {
                        Min = new XYZ(minX, minY, minZ),
                        Max = new XYZ(maxX, maxY, maxZ),
                        Transform = Transform.Identity // World coordinates
                    };
                }
                
                // If we couldn't get view bounds, assume elements might be visible
                if (viewBounds == null)
                    return true;
                
                // Check if any element's bounding box intersects with view bounds
                foreach (var kvp in elementBoundingBoxes)
                {
                    var elementBounds = kvp.Value;
                    
                    // Check for intersection
                    if (BoundingBoxesIntersect(viewBounds, elementBounds))
                        return true;
                }
                
                return false;
            }
            catch
            {
                // If there's any error, assume elements might be visible
                return true;
            }
        }
        
        private bool BoundingBoxesIntersect(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            // Check if boxes intersect in all three dimensions
            return !(box1.Max.X < box2.Min.X || box1.Min.X > box2.Max.X ||
                     box1.Max.Y < box2.Min.Y || box1.Min.Y > box2.Max.Y ||
                     box1.Max.Z < box2.Min.Z || box1.Min.Z > box2.Max.Z);
        }
        
        private Dictionary<ElementId, ViewSheet> BuildViewSheetMapping()
        {
            var viewSheetMap = new Dictionary<ElementId, ViewSheet>();
            
            try
            {
                // Get all sheets
                var sheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .ToList();
                
                if (cancelRequested) return viewSheetMap;
                
                foreach (var sheet in sheets)
                {
                    if (cancelRequested) return viewSheetMap;
                    
                    // Get all views placed on this sheet
                    try
                    {
                        var viewports = new FilteredElementCollector(doc, sheet.Id)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>();
                        
                        foreach (var viewport in viewports)
                        {
                            var viewId = viewport.ViewId;
                            if (!viewSheetMap.ContainsKey(viewId))
                            {
                                viewSheetMap[viewId] = sheet;
                            }
                        }
                    }
                    catch { /* Skip problematic sheets */ }
                }
            }
            catch { /* Continue without sheet info if there's an error */ }
            
            return viewSheetMap;
        }
        
        // Check if a view is a graphical view (not a schedule, legend, etc.)
        private bool IsGraphicalView(Autodesk.Revit.DB.View view)
        {
            // Exclude non-graphical view types
            if (view.ViewType == ViewType.Schedule ||
                view.ViewType == ViewType.ColumnSchedule ||
                view.ViewType == ViewType.PanelSchedule ||
                view.ViewType == ViewType.Legend ||
                view.ViewType == ViewType.Undefined ||
                view.ViewType == ViewType.Internal ||
                view.ViewType == ViewType.Report ||
#if !REVIT2017 && !REVIT2018 && !REVIT2019
                view.ViewType == ViewType.SystemsAnalysisReport ||
#endif
                view.ViewType == ViewType.CostReport ||
                view.ViewType == ViewType.LoadsReport ||
                view is ViewSheet ||
                view is ViewSchedule ||
                view is TableView)
            {
                return false;
            }
            
            return true;
        }
        
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            // Start the search on the UI thread
            SearchForViews();
        }
        
        private string GetViewLevel(Autodesk.Revit.DB.View view)
        {
            try
            {
                // Try to get associated level
                if (view.GenLevel != null)
                    return view.GenLevel.Name;
                
                // For plan views
                var param = view.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL);
                if (param != null && param.HasValue)
                {
                    var levelId = param.AsElementId();
                    if (levelId != ElementId.InvalidElementId)
                    {
                        var level = doc.GetElement(levelId) as Level;
                        if (level != null)
                            return level.Name;
                    }
                }
                
                // Try to get level from view name (often contains level info)
                if (view is ViewPlan || view is ViewSection)
                {
                    // Check if view has an associated level through its properties
                    var levelParam = view.LookupParameter("Associated Level");
                    if (levelParam == null)
                        levelParam = view.LookupParameter("Reference Level");
                    
                    if (levelParam != null && levelParam.StorageType == StorageType.ElementId)
                    {
                        var levelId = levelParam.AsElementId();
                        if (levelId != ElementId.InvalidElementId)
                        {
                            var level = doc.GetElement(levelId) as Level;
                            if (level != null)
                                return level.Name;
                        }
                    }
                }
            }
            catch { }
            
            return "-";
        }
        
        private string GetViewScale(Autodesk.Revit.DB.View view)
        {
            try
            {
                if (view.Scale > 0)
                    return $"1:{view.Scale}";
            }
            catch { }
            
            return "-";
        }
    }
    
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = commandData.Application.ActiveUIDocument;
        var doc = uidoc.Document;
        
        try
        {
            // Get currently selected elements
            var selectedIds = uidoc.GetSelectionIds();
            
            if (selectedIds == null || selectedIds.Count == 0)
            {
                TaskDialog.Show("No Selection", "Please select one or more elements before running this command.");
                return Result.Cancelled;
            }
            
            // Show progress dialog with live results - dialog will collect views internally
            List<ViewInfo> selectedViews;
            List<ViewInfo> foundViews;
            using (var progressDialog = new ProgressDialog(doc, selectedIds.ToList()))
            {
                var dialogResult = progressDialog.ShowDialog();
                selectedViews = progressDialog.SelectedViews;
                foundViews = progressDialog.FoundViews;
                
                // If user cancelled or closed without selecting, and no views were found
                if (dialogResult != WinForms.DialogResult.OK && foundViews.Count == 0)
                {
                    // Check if any views were found at all
                    if (foundViews.Count == 0)
                    {
                        TaskDialog.Show("No Views Found", 
                            "Selected elements are not visible in any views.\n\n" +
                            "Note: Elements might be hidden, outside view range, or in a hidden category.");
                    }
                    return Result.Cancelled;
                }
                
                // If user didn't explicitly select views, use all found views
                if (dialogResult != WinForms.DialogResult.OK || selectedViews.Count == 0)
                {
                    selectedViews = foundViews;
                }
            }
            
            // Only proceed if we have views to work with
            if (selectedViews != null && selectedViews.Count > 0)
            {
                // Use the CustomGUIs.DataGrid to show final selection
                var viewData = new List<Dictionary<string, object>>();
                foreach (var viewInfo in selectedViews)
                {
                    var dict = new Dictionary<string, object>
                    {
                        { "ViewName", viewInfo.ViewName },
                        { "ViewType", viewInfo.ViewType },
                        { "Level", viewInfo.Level },
                        { "Sheet", viewInfo.Sheet },
                        { "SheetNumber", viewInfo.SheetNumber },
                        { "Scale", viewInfo.Scale },
                        { "Elements", viewInfo.ElementsFound },
                        { "ViewId", viewInfo.ViewId },
                        { "ViewObject", viewInfo.ViewObject }
                    };
                    viewData.Add(dict);
                }
                
                // Show in CustomGUIs.DataGrid for final selection
                var propertyNames = new List<string> { 
                    "ViewName", "ViewType", "Level", "Sheet", "SheetNumber", "Scale", "Elements", "ViewId" 
                };
                
                var finalSelection = CustomGUIs.DataGrid(viewData, propertyNames, false);
                
                if (finalSelection != null && finalSelection.Count > 0)
                {
                    // Get the selected views and convert to ElementIds
                    var viewElementIds = finalSelection
                        .Select(dict => dict["ViewObject"] as Autodesk.Revit.DB.View)
                        .Where(v => v != null)
                        .Select(v => v.Id)
                        .ToList();
                    
                    if (viewElementIds.Count > 0)
                    {
                        // Select the views
                        uidoc.SetSelectionIds(viewElementIds);
                        
                        // Show confirmation
                        TaskDialog.Show("Views Selected", 
                            $"Selected {viewElementIds.Count} view(s) where the elements are visible.");
                    }
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
