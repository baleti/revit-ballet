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
        // First get the standard element data using the existing helper
        var elementData = ElementDataHelper.GetElementData(uiDoc, selectedOnly, includeParameters);
        
        // Enhance each element's data
        foreach (var data in elementData)
        {
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
                AddSheetInfo(element, elementDoc, data);
                
                // Add area if not already present
                AddArea(element, data);
                
                // Add perimeter for elements that support it
                AddPerimeter(element, data);
            }
        }
        
        return elementData;
    }
    
    private static void AddSheetInfo(Element element, Document doc, Dictionary<string, object> data)
    {
        string sheetNumber = string.Empty;
        string sheetTitle = string.Empty;
        string viewTitleOnSheet = string.Empty;
        
        if (element.OwnerViewId != null && element.OwnerViewId != ElementId.InvalidElementId)
        {
            View ownerView = doc.GetElement(element.OwnerViewId) as View;
            if (ownerView != null)
            {
                // Check if view is on a sheet
                var viewSheetId = ownerView.get_Parameter(BuiltInParameter.VIEWER_SHEET_NUMBER);
                if (viewSheetId != null && !string.IsNullOrEmpty(viewSheetId.AsString()))
                {
                    // Find the sheet that contains this view
                    var sheets = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>();
                    
                    foreach (var sheet in sheets)
                    {
                        var viewportId = sheet.GetAllViewports().FirstOrDefault(vpId => 
                        {
                            var vp = doc.GetElement(vpId) as Viewport;
                            return vp != null && vp.ViewId == ownerView.Id;
                        });
                        
                        if (viewportId != null && viewportId != ElementId.InvalidElementId)
                        {
                            sheetNumber = sheet.SheetNumber;
                            sheetTitle = sheet.Name;
                            
                            // Get the viewport to access "Title on Sheet"
                            var viewport = doc.GetElement(viewportId) as Viewport;
                            if (viewport != null)
                            {
                                // Try to get the title on sheet from viewport
                                var titleParam = viewport.get_Parameter(BuiltInParameter.VIEWPORT_VIEW_NAME);
                                if (titleParam == null || !titleParam.HasValue || string.IsNullOrEmpty(titleParam.AsString()))
                                {
                                    // Try alternative parameter
                                    titleParam = viewport.LookupParameter("Title on Sheet");
                                }
                                
                                if (titleParam != null && titleParam.HasValue)
                                {
                                    viewTitleOnSheet = titleParam.AsString() ?? string.Empty;
                                }
                            }
                            
                            break;
                        }
                    }
                }
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
                "ScopeBoxes", 
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
