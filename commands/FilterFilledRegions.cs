using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// Extended helper for FilterFilledRegions command that adds sheet info and perimeter
/// </summary>
public static class ExtendedElementDataHelper
{
    public static List<Dictionary<string, object>> GetElementDataWithSheetInfo(UIDocument uiDoc, bool selectedOnly = false, bool includeParameters = false)
    {
        var elementData = ElementDataHelper.GetElementData(uiDoc, selectedOnly, includeParameters);

        var doc = uiDoc.Document;

        // Precollect scope boxes
        var scopeBoxes = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
            .ToElements()
            .Cast<Element>()
            .ToList();

        // Precompute view sheet info
        var sheets = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet))
            .Cast<ViewSheet>()
            .ToList();

        var viewSheetInfo = new Dictionary<ElementId, Tuple<string, string, string>>();

        foreach (var sheet in sheets)
        {
            foreach (var vpId in sheet.GetAllViewports())
            {
                var vp = doc.GetElement(vpId) as Viewport;
                if (vp != null)
                {
                    var viewId = vp.ViewId;
                    string viewTitleOnSheet = string.Empty;
                    var titleParam = vp.get_Parameter(BuiltInParameter.VIEWPORT_VIEW_NAME);
                    if (titleParam == null || !titleParam.HasValue || string.IsNullOrEmpty(titleParam.AsString()))
                    {
                        titleParam = vp.LookupParameter("Title on Sheet");
                    }
                    if (titleParam != null && titleParam.HasValue)
                    {
                        viewTitleOnSheet = titleParam.AsString() ?? string.Empty;
                    }
                    viewSheetInfo[viewId] = Tuple.Create(sheet.SheetNumber, sheet.Name, viewTitleOnSheet);
                }
            }
        }

        // Enhance each element's data
        foreach (var data in elementData)
        {
            // Rename original ScopeBoxes key if present
            if (data.ContainsKey("ScopeBoxes"))
            {
                data["Scope Boxes Outline"] = data["ScopeBoxes"];
                data.Remove("ScopeBoxes");
            }

            // Get the element to add sheet info
            Element element = null;
            Document elementDoc = uiDoc.Document;
            
            // Try to get the element from the data
            if (data.ContainsKey("ElementIdObject") && data["ElementIdObject"] is ElementId elemId)
            {
                element = elementDoc.GetElement(elemId);
            }
            else if (data.ContainsKey("Id") && data["Id"] is int id)
            {
                element = elementDoc.GetElement(new ElementId(id));
            }
            
            if (element != null)
            {
                // Add sheet information
                AddSheetInfo(element, doc, data, viewSheetInfo);
                
                // Add area if not already present
                AddArea(element, data);
                
                // Add perimeter for elements that support it
                AddPerimeter(element, data);

                // Add scope boxes center for filled regions
                AddScopeBoxesCenter(element, doc, data, scopeBoxes);
            }
        }
        
        return elementData;
    }
    
    private static void AddSheetInfo(Element element, Document doc, Dictionary<string, object> data, Dictionary<ElementId, Tuple<string, string, string>> viewSheetInfo)
    {
        string sheetNumber = string.Empty;
        string sheetTitle = string.Empty;
        string viewTitleOnSheet = string.Empty;
        
        if (element.OwnerViewId != null && element.OwnerViewId != ElementId.InvalidElementId)
        {
            if (viewSheetInfo.TryGetValue(element.OwnerViewId, out var info))
            {
                sheetNumber = info.Item1;
                sheetTitle = info.Item2;
                viewTitleOnSheet = info.Item3;
            }
        }
        
        data["SheetNumber"] = sheetNumber;
        data["SheetTitle"] = sheetTitle;
        data["ViewTitleOnSheet"] = viewTitleOnSheet;
    }
    
    private static void AddArea(Element element, Dictionary<string, object> data)
    {
        // Skip if area already exists
        if (data.ContainsKey("Area"))
            return;
            
        try
        {
            // Try to get area parameter
            var areaParam = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
            if (areaParam != null && areaParam.HasValue)
            {
                double areaInSqFt = areaParam.AsDouble();
                if (areaInSqFt > 0)
                {
                    Units projectUnits = element.Document.GetUnits();
                    FormatOptions areaOptions = projectUnits.GetFormatOptions(SpecTypeId.Area);
                    ForgeTypeId unitTypeId = areaOptions.GetUnitTypeId();
                    double convertedValue = UnitUtils.ConvertFromInternalUnits(areaInSqFt, unitTypeId);
                    string unitSymbol = LabelUtils.GetLabelForUnit(unitTypeId);
                    data["Area"] = $"{convertedValue:F2} {unitSymbol}";
                    return;
                }
            }
        }
        catch { }
        
        // If no area found or zero, set empty string
        data["Area"] = string.Empty;
    }
    
    private static void AddPerimeter(Element element, Dictionary<string, object> data)
    {
            try
            {
                double perimeterInFeet = 0.0;

                // Try to get perimeter for elements that support it
                var perimeterParam = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                if (perimeterParam != null && perimeterParam.HasValue)
                {
                    perimeterInFeet = perimeterParam.AsDouble();
                }
                else if (element is FilledRegion filledRegion)
                {
                    var boundaries = filledRegion.GetBoundaries();
                    foreach (var boundary in boundaries)
                    {
                        foreach (var curve in boundary)
                        {
                            perimeterInFeet += curve.Length;
                        }
                    }
                }

                AddFormattedPerimeter(element.Document, data, perimeterInFeet);
            }
            catch { }
            
            // If no perimeter found, set empty string
            if (!data.ContainsKey("Perimeter"))
            {
                data["Perimeter"] = string.Empty;
            }
        }

    private static void AddFormattedPerimeter(Document doc, Dictionary<string, object> data, double perimeterInFeet)
    {
        if (perimeterInFeet <= 0)
            {
                data["Perimeter"] = string.Empty;
                return;
            }

        ForgeTypeId unitTypeId = UnitTypeId.Meters;
        double convertedValue = UnitUtils.ConvertFromInternalUnits(perimeterInFeet, unitTypeId);
        data["Perimeter"] = $"{convertedValue:F2} m";
    }

    private static void AddScopeBoxesCenter(Element element, Document doc, Dictionary<string, object> data, List<Element> scopeBoxes)
    {
        if (!(element is FilledRegion))
        {
            data["Scope Boxes Center"] = string.Empty;
            return;
        }

        View ownerView = doc.GetElement(element.OwnerViewId) as View;
        if (ownerView == null)
        {
            data["Scope Boxes Center"] = string.Empty;
            return;
        }

        BoundingBoxXYZ bb = element.get_BoundingBox(ownerView);
        if (bb == null)
        {
            data["Scope Boxes Center"] = string.Empty;
            return;
        }

        XYZ center = (bb.Min + bb.Max) / 2.0;

        var containingScopeBoxes = scopeBoxes
            .Where(sb => IsPointInScopeBox(sb, center))
            .Select(sb => sb.Name)
            .OrderBy(name => name)
            .ToList();

        data["Scope Boxes Center"] = string.Join(", ", containingScopeBoxes);
    }

    private static bool IsPointInScopeBox(Element scopeBox, XYZ point)
    {
        BoundingBoxXYZ bb = scopeBox.get_BoundingBox(null);
        if (bb == null) return false;

        Transform inverse = bb.Transform.Inverse;
        XYZ localPt = inverse.OfPoint(point);

        return (bb.Min.X <= localPt.X && localPt.X <= bb.Max.X) &&
               (bb.Min.Y <= localPt.Y && localPt.Y <= bb.Max.Y) &&
               (bb.Min.Z <= localPt.Z && localPt.Z <= bb.Max.Z);
    }
}

/// <summary>
/// Filter filled regions (and other selected elements) with sheet information and improved scope box detection
/// </summary>
[Transaction(TransactionMode.Manual)]
public class FilterFilledRegions : IExternalCommand
{
    public Result Execute(ExternalCommandData cData, ref string message, ElementSet elements)
    {
        try
        {
            var uiDoc = cData.Application.ActiveUIDocument;
            
            // Use the extended helper to get element data with sheet info
            var elementData = ExtendedElementDataHelper.GetElementDataWithSheetInfo(uiDoc, selectedOnly: true, includeParameters: true);

            if (!elementData.Any())
            {
                TaskDialog.Show("Info", "No elements found.");
                return Result.Cancelled;
            }

            // Create a unique key for each element to map back to full data
            var elementDataMap = new Dictionary<string, Dictionary<string, object>>();
            var displayData = new List<Dictionary<string, object>>();
            
            for (int i = 0; i < elementData.Count; i++)
            {
                var data = elementData[i];
                // Create a unique key combining multiple properties
                string linkName = data.ContainsKey("LinkName") ? data["LinkName"].ToString() : "";
                string uniqueKey = $"{data["Id"]}_{data["Name"]}_{data["Category"]}_{linkName}_{i}";
                
                // Store the full data in our map
                elementDataMap[uniqueKey] = data;
                
                // Create display data with the unique key
                var display = new Dictionary<string, object>(data);
                display["UniqueKey"] = uniqueKey;
                displayData.Add(display);
            }

            // Get property names, including UniqueKey but excluding internal object fields
            var propertyNames = displayData.First().Keys
                .Where(k => !k.EndsWith("Object"))
                .ToList();

            // Reorder to put most useful columns first, including sheet info
            var orderedProps = new List<string> { 
                "DisplayName", 
                "Scope Boxes Outline", 
                "Scope Boxes Center",
                "Category",
                "SheetNumber",
                "SheetTitle",
                "OwnerView",
                "ViewTitleOnSheet",
                "LinkName", 
                "Group", 
                "Area",
                "Perimeter",
                "Id" 
            };
            
            var remainingProps = propertyNames.Except(orderedProps).OrderBy(p => p);
            propertyNames = orderedProps.Where(p => propertyNames.Contains(p))
                .Concat(remainingProps)
                .ToList();

            var chosenRows = CustomGUIs.DataGrid(displayData, propertyNames, spanAllScreens: false);
            if (chosenRows.Count == 0)
                return Result.Cancelled;

            // Handle selection (same logic as original)
            var regularIds = new List<ElementId>();
            var linkedReferences = new List<Reference>();

            foreach (var row in chosenRows)
            {
                // Get the unique key to look up full data
                if (!row.TryGetValue("UniqueKey", out var keyObj) || !(keyObj is string uniqueKey))
                    continue;
                
                // Get the full data from our map
                if (!elementDataMap.TryGetValue(uniqueKey, out var fullData))
                    continue;
                
                // Check if this is a linked element
                if (fullData.TryGetValue("LinkInstanceObject", out var linkObj) && linkObj is RevitLinkInstance linkInstance &&
                    fullData.TryGetValue("LinkedElementIdObject", out var linkedIdObj) && linkedIdObj is ElementId linkedElementId)
                {
                    // This is a linked element - create reference
                    try
                    {
                        Document linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            Element linkedElement = linkedDoc.GetElement(linkedElementId);
                            if (linkedElement != null)
                            {
                                Reference elemRef = new Reference(linkedElement);
                                Reference linkedRef = elemRef.CreateLinkReference(linkInstance);
                                
                                if (linkedRef != null)
                                {
                                    linkedReferences.Add(linkedRef);
                                }
                            }
                        }
                    }
                    catch { }
                }
                else if (fullData.TryGetValue("ElementIdObject", out var idObj) && idObj is ElementId elemId)
                {
                    regularIds.Add(elemId);
                }
                else if (fullData.TryGetValue("Id", out var intId) && intId is int id)
                {
                    regularIds.Add(new ElementId(id));
                }
            }

            // Set selection based on what we have
            if (linkedReferences.Any() && !regularIds.Any())
            {
                uiDoc.SetReferences(linkedReferences);
            }
            else if (!linkedReferences.Any() && regularIds.Any())
            {
                uiDoc.SetSelectionIds(regularIds);
            }
            else if (linkedReferences.Any() && regularIds.Any())
            {
                // Mixed selection - prioritize what we have more of
                if (linkedReferences.Count >= regularIds.Count)
                {
                    var allReferences = new List<Reference>(linkedReferences);
                    
                    foreach (var elemId in regularIds)
                    {
                        try
                        {
                            Element elem = cData.Application.ActiveUIDocument.Document.GetElement(elemId);
                            if (elem != null)
                            {
                                Reference elemRef = new Reference(elem);
                                if (elemRef != null)
                                {
                                    allReferences.Add(elemRef);
                                }
                            }
                        }
                        catch { }
                    }
                    
                    uiDoc.SetReferences(allReferences);
                }
                else
                {
                    uiDoc.SetSelectionIds(regularIds);
                    
                    TaskDialog.Show("Mixed Selection",
                        $"Selected {regularIds.Count} regular elements.\n" +
                        $"Note: {linkedReferences.Count} linked elements could not be included in the selection " +
                        "because Revit doesn't support mixing regular and linked selections.");
                }
            }

            return Result.Succeeded;
        }
        catch (InvalidOperationException ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
        catch (Exception ex)
        {
            message = $"Unexpected error: {ex.Message}";
            return Result.Failed;
        }
    }
}
