using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;
using System;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// Filters selected elements across all open documents in the session.
/// Uses SelectionStorage for cross-document selection tracking.
/// </summary>
[Transaction(TransactionMode.Manual)]
public class FilterSelectedInSession : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document activeDoc = uidoc.Document;

            // Load selection from SelectionStorage
            var storedSelection = SelectionStorage.LoadSelection();

            if (!storedSelection.Any())
            {
                TaskDialog.Show("Info", "No elements in session selection.\n\nUse session-scope commands like SelectByCategoriesInSession to build a cross-document selection first.");
                return Result.Cancelled;
            }

            // Collect element data from all open documents
            List<Dictionary<string, object>> elementData = new List<Dictionary<string, object>>();

            using (var progress = new CancellableProgressDialog("Collecting element data"))
            {
                progress.Start();
                progress.SetTotal(storedSelection.Count);

                // Group by document for efficient processing
                var byDocument = storedSelection.GroupBy(s => s.DocumentTitle);

                foreach (var docGroup in byDocument)
                {
                    // Find matching open document
                    Document doc = null;
                    foreach (Document d in app.Documents)
                    {
                        if (d.IsLinked) continue;
                        if (d.Title == docGroup.Key)
                        {
                            doc = d;
                            break;
                        }
                    }

                    if (doc == null)
                    {
                        // Document not open, skip these elements
                        foreach (var _ in docGroup)
                        {
                            progress.IncrementProgress();
                        }
                        continue;
                    }

                    // Process elements in this document
                    foreach (var selItem in docGroup)
                    {
                        progress.CheckAndShow();

                        if (progress.IsCancelled)
                            throw new OperationCanceledException("Operation cancelled by user.");

                        try
                        {
                            // Get element by UniqueId
                            Element elem = doc.GetElement(selItem.UniqueId);
                            if (elem != null)
                            {
                                // Build element data dictionary using ElementDataHelper
                                // Note: We can't use ElementDataHelper.GetElementData directly because it works on UIDocument
                                // Instead, we'll build the data dictionary manually using similar logic
                                var data = BuildElementDataDictionary(elem, doc);
                                elementData.Add(data);
                            }
                        }
                        catch { /* Skip problematic elements */ }

                        progress.IncrementProgress();
                    }
                }
            }

            if (!elementData.Any())
            {
                TaskDialog.Show("Info", "None of the selected elements are available in currently open documents.");
                return Result.Cancelled;
            }

            // Get ALL property names from ALL elements (union, not intersection)
            var allPropertyNames = new HashSet<string>();
            foreach (var data in elementData)
            {
                foreach (var key in data.Keys)
                {
                    if (!key.EndsWith("Object"))  // Exclude internal object fields
                    {
                        allPropertyNames.Add(key);
                    }
                }
            }

            // Build ordered list with standard columns in preferred order
            var orderedProps = new List<string> { "Name" };
            if (allPropertyNames.Contains("Document")) orderedProps.Add("Document");  // Show document name for session scope
            if (allPropertyNames.Contains("Type Name")) orderedProps.Add("Type Name");  // Editable
            if (allPropertyNames.Contains("Family")) orderedProps.Add("Family");  // Editable
            if (allPropertyNames.Contains("Scope Box")) orderedProps.Add("Scope Box");  // View scope box (editable)
            if (allPropertyNames.Contains("ScopeBoxes")) orderedProps.Add("ScopeBoxes");  // Element containment (read-only)
            orderedProps.Add("Category");
            if (allPropertyNames.Contains("Group")) orderedProps.Add("Group");
            if (allPropertyNames.Contains("OwnerView")) orderedProps.Add("OwnerView");

            // Add crop region columns (editable)
            if (allPropertyNames.Contains("Crop Region Top")) orderedProps.Add("Crop Region Top");
            if (allPropertyNames.Contains("Crop Region Bottom")) orderedProps.Add("Crop Region Bottom");
            if (allPropertyNames.Contains("Crop Region Left")) orderedProps.Add("Crop Region Left");
            if (allPropertyNames.Contains("Crop Region Right")) orderedProps.Add("Crop Region Right");

            // Add centroid columns
            if (allPropertyNames.Contains("X Centroid")) orderedProps.Add("X Centroid");
            if (allPropertyNames.Contains("Y Centroid")) orderedProps.Add("Y Centroid");
            if (allPropertyNames.Contains("Z Centroid")) orderedProps.Add("Z Centroid");

            orderedProps.Add("Id");

            var remainingProps = allPropertyNames.Except(orderedProps).OrderBy(p => p);
            var propertyNames = orderedProps.Where(p => allPropertyNames.Contains(p))
                .Concat(remainingProps)
                .ToList();

            // Set the current UIDocument for edit operations
            CustomGUIs.SetCurrentUIDocument(uidoc);

            var chosenRows = CustomGUIs.DataGrid(elementData, propertyNames, false);

            // Apply any pending edits to Revit elements
            if (CustomGUIs.HasPendingEdits())
            {
                CustomGUIs.ApplyCellEditsToEntities();
            }

            if (chosenRows.Count == 0)
            {
                // Clear selection storage
                SelectionStorage.ClearSelection();
                return Result.Cancelled;
            }

            // Update SelectionStorage with filtered selection
            var newSelection = new List<SelectionItem>();

            foreach (var row in chosenRows)
            {
                if (row.TryGetValue("UniqueId", out var uniqueIdObj) && uniqueIdObj is string uniqueId &&
                    row.TryGetValue("DocumentTitle", out var docTitleObj) && docTitleObj is string docTitle &&
                    row.TryGetValue("DocumentPath", out var docPathObj) && docPathObj is string docPath &&
                    row.TryGetValue("Id", out var idObj))
                {
                    long idValue = 0;
                    if (idObj is int intId)
                        idValue = intId;
                    else if (idObj is long longId)
                        idValue = longId;

                    newSelection.Add(new SelectionItem
                    {
                        DocumentTitle = docTitle,
                        DocumentPath = docPath,
                        UniqueId = uniqueId,
                        ElementIdValue = (int)idValue
                    });
                }
            }

            // Save filtered selection
            SelectionStorage.SaveSelection(newSelection);

            return Result.Succeeded;
        }
        catch (OperationCanceledException)
        {
            message = "Operation cancelled by user.";
            return Result.Cancelled;
        }
        catch (Exception ex)
        {
            message = $"Unexpected error: {ex.Message}";
            return Result.Failed;
        }
    }

    /// <summary>
    /// Build element data dictionary similar to ElementDataHelper but for session scope
    /// </summary>
    private Dictionary<string, object> BuildElementDataDictionary(Element element, Document doc)
    {
        string groupName = string.Empty;
        if (element.GroupId != null && element.GroupId != ElementId.InvalidElementId && element.GroupId.AsLong() != -1)
        {
            if (doc.GetElement(element.GroupId) is Group g)
                groupName = g.Name;
        }

        string ownerViewName = string.Empty;
        if (element.OwnerViewId != null && element.OwnerViewId != ElementId.InvalidElementId)
        {
            if (doc.GetElement(element.OwnerViewId) is View v)
                ownerViewName = v.Name;
        }

        var data = new Dictionary<string, object>
        {
            ["Name"] = element.Name,
            ["Category"] = element.Category?.Name ?? string.Empty,
            ["Group"] = groupName,
            ["OwnerView"] = ownerViewName,
            ["Id"] = element.Id.AsLong(),
            ["ElementIdObject"] = element.Id,
            ["UniqueId"] = element.UniqueId,
            ["Document"] = doc.Title,
            ["DocumentTitle"] = doc.Title,
            ["DocumentPath"] = doc.PathName ?? ""
        };

        // Add Type Name and Family columns
        ElementId typeId = element.GetTypeId();
        if (typeId != null && typeId != ElementId.InvalidElementId)
        {
            Element typeElement = doc.GetElement(typeId);
            if (typeElement != null)
            {
                data["Type Name"] = typeElement.Name;

                if (typeElement is FamilySymbol familySymbol)
                {
                    data["Family"] = familySymbol.FamilyName;
                }
                else
                {
                    Parameter familyParam = typeElement.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                    data["Family"] = (familyParam != null && !string.IsNullOrEmpty(familyParam.AsString()))
                        ? familyParam.AsString()
                        : "System Type";
                }
            }
        }

        // Add view-specific scope box property
        if (element is View view)
        {
            try
            {
                Parameter scopeBoxParam = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP);
                if (scopeBoxParam != null && scopeBoxParam.HasValue)
                {
                    ElementId scopeBoxId = scopeBoxParam.AsElementId();
                    if (scopeBoxId != null && scopeBoxId != ElementId.InvalidElementId)
                    {
                        Element assignedScopeBox = doc.GetElement(scopeBoxId);
                        data["Scope Box"] = assignedScopeBox?.Name ?? "";
                    }
                    else
                    {
                        data["Scope Box"] = "";
                    }
                }
            }
            catch { }
        }

        // Add crop region columns for views and viewports (only if rectangular)
        try
        {
            Autodesk.Revit.DB.View viewForCrop = null;
            if (element is Autodesk.Revit.DB.View v)
            {
                viewForCrop = v;
            }
            else if (element is Viewport viewport)
            {
                viewForCrop = doc.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;
            }

            if (viewForCrop != null && viewForCrop.CropBoxActive && viewForCrop.CropBox != null)
            {
                // Check if crop region is rectangular
                bool isRectangular = true;
                if (viewForCrop is ViewPlan plan)
                {
                    try
                    {
                        var managerProperty = plan.GetType().GetProperty("CropRegionShapeManager");
                        if (managerProperty != null)
                        {
                            object manager = managerProperty.GetValue(plan, null);
                            if (manager != null)
                            {
                                var method = manager.GetType().GetMethod("GetCropRegionShape");
                                if (method != null)
                                {
                                    CurveLoop cropLoop = method.Invoke(manager, null) as CurveLoop;
                                    if (cropLoop != null)
                                    {
                                        isRectangular = cropLoop.ToList().Count == 4;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        isRectangular = false;
                    }
                }

                if (isRectangular)
                {
                    BoundingBoxXYZ bbox = viewForCrop.CropBox;
                    Transform transform = bbox.Transform;

                    // Get survey point elevation
                    double surveyElevation = 0.0;
                    try
                    {
                        FilteredElementCollector surveyCollector = new FilteredElementCollector(doc);
                        var surveyPoint = surveyCollector.OfCategory(BuiltInCategory.OST_SharedBasePoint)
                            .FirstOrDefault() as BasePoint;
                        if (surveyPoint != null)
                        {
                            surveyElevation = surveyPoint.get_BoundingBox(null).Min.Z;
                        }
                    }
                    catch { }

                    // Transform crop box corners to project coordinates
                    XYZ minCorner = transform.OfPoint(bbox.Min);
                    XYZ maxCorner = transform.OfPoint(bbox.Max);

                    // Determine coordinate interpretation based on view type
                    bool isElevationOrSection = viewForCrop.ViewType == ViewType.Elevation || viewForCrop.ViewType == ViewType.Section;

                    double top, bottom, left, right;
                    if (isElevationOrSection)
                    {
                        // For elevation/section: Top/Bottom are elevations (Z) relative to survey point
                        top = maxCorner.Z - surveyElevation;
                        bottom = minCorner.Z - surveyElevation;
                        left = minCorner.X;
                        right = maxCorner.X;
                    }
                    else
                    {
                        // For plan: Top/Bottom are Y, Left/Right are X
                        top = maxCorner.Y;
                        bottom = minCorner.Y;
                        left = minCorner.X;
                        right = maxCorner.X;
                    }

                    // Convert from internal units (feet) to display units
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                    Units projectUnits = doc.GetUnits();
                    FormatOptions lengthOpts = projectUnits.GetFormatOptions(SpecTypeId.Length);
                    ForgeTypeId unitTypeId = lengthOpts.GetUnitTypeId();

                    data["Crop Region Top"] = UnitUtils.ConvertFromInternalUnits(top, unitTypeId);
                    data["Crop Region Bottom"] = UnitUtils.ConvertFromInternalUnits(bottom, unitTypeId);
                    data["Crop Region Left"] = UnitUtils.ConvertFromInternalUnits(left, unitTypeId);
                    data["Crop Region Right"] = UnitUtils.ConvertFromInternalUnits(right, unitTypeId);
#else
                    // Revit 2017-2020: Use DisplayUnitType
                    Units projectUnits = doc.GetUnits();
                    FormatOptions lengthOpts = projectUnits.GetFormatOptions(UnitType.UT_Length);
                    DisplayUnitType unitType = lengthOpts.DisplayUnits;

                    data["Crop Region Top"] = UnitUtils.ConvertFromInternalUnits(top, unitType);
                    data["Crop Region Bottom"] = UnitUtils.ConvertFromInternalUnits(bottom, unitType);
                    data["Crop Region Left"] = UnitUtils.ConvertFromInternalUnits(left, unitType);
                    data["Crop Region Right"] = UnitUtils.ConvertFromInternalUnits(right, unitType);
#endif
                }
            }
        }
        catch { }

        return data;
    }
}
