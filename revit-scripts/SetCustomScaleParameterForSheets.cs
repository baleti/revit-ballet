using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class SetCustomScaleParameterForSheets : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData, 
        ref string message, 
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get all global and shared parameters from the entire project
        List<BindingMap> bindingMaps = new List<BindingMap>
        {
            doc.ParameterBindings
        };

        HashSet<string> globalAndSharedParameterNames = new HashSet<string>();

        foreach (BindingMap bindingMap in bindingMaps)
        {
            DefinitionBindingMapIterator iter = bindingMap.ForwardIterator();
            iter.Reset();
            while (iter.MoveNext())
            {
                Definition def = iter.Key as Definition;
                if (def != null)
                {
                    // Add the parameter name to the set
                    globalAndSharedParameterNames.Add(def.Name);
                }
            }
        }

        // Convert global and shared parameters to a list of dictionaries for display
        List<Dictionary<string, object>> parameterEntries = globalAndSharedParameterNames
            .Select(paramName => new Dictionary<string, object>
            {
                { "Parameter Name", paramName }
            }).ToList();

        // Show parameters selection dialog
        List<string> parameterProperties = new List<string> { "Parameter Name" };
        List<Dictionary<string, object>> selectedParameter = CustomGUIs.DataGrid(parameterEntries, parameterProperties, false);

        if (selectedParameter == null || selectedParameter.Count == 0)
        {
            return Result.Cancelled;
        }

        // Get all sheets in the project and store both Title and ElementId
        FilteredElementCollector collector = new FilteredElementCollector(doc);
        ICollection<Element> allSheets = collector.OfClass(typeof(ViewSheet)).ToElements();

        List<Dictionary<string, object>> sheetEntries = allSheets
            .Select(sheet => new Dictionary<string, object>
            {
                { "Title", $"{(sheet as ViewSheet).SheetNumber} - {sheet.Name}" },
                { "ElementId", sheet.Id }  // Store the actual ElementId
            }).ToList();

        // Show sheets selection dialog
        List<string> sheetProperties = new List<string> { "Title", "ElementId" };
        List<Dictionary<string, object>> selectedSheets = CustomGUIs.DataGrid(sheetEntries, sheetProperties, false);

        if (selectedSheets == null || selectedSheets.Count == 0)
        {
            return Result.Cancelled;
        }

        string parameterName = selectedParameter[0]["Parameter Name"].ToString();
        List<string> inconsistentSheets = new List<string>();

        using (Transaction trans = new Transaction(doc, "Set Custom Scale Parameter for Sheets"))
        {
            trans.Start();

            foreach (var sheetDict in selectedSheets)
            {
                if (sheetDict.TryGetValue("ElementId", out object elementIdObj) && elementIdObj is ElementId elementId)
                {
                    ViewSheet sheet = doc.GetElement(elementId) as ViewSheet;

                    if (sheet != null)
                    {
                        // Get the scales of the views on the sheet
                        var viewports = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .WhereElementIsNotElementType()
                            .Cast<Viewport>()
                            .Where(vp => vp.SheetId == sheet.Id);

                        var viewScales = viewports
                            .Select(vp => doc.GetElement(vp.ViewId) as View)
                            .Where(v => v != null && v.ViewType != ViewType.Legend) // Exclude legends
                            .Select(v => v.Scale)
                            .Distinct()
                            .ToList();

                        if (viewScales.Count == 1) // All views have the same scale
                        {
                            string scaleValue = $"1:{viewScales.First()}";
                            Parameter param = sheet.LookupParameter(parameterName);
                            if (param != null && param.StorageType == StorageType.String)
                            {
                                param.Set(scaleValue);
                            }
                        }
                        else
                        {
                            // Log the sheet as having inconsistent scales
                            inconsistentSheets.Add($"{sheet.SheetNumber} - {sheet.Name}");
                        }
                    }
                }
                else
                {
                    message = "Error: Unable to find the ElementId key in the selected sheet dictionary.";
                    return Result.Failed;
                }
            }

            trans.Commit();
        }

        if (inconsistentSheets.Count > 0)
        {
            // Display a message showing the sheets with inconsistent scales
            TaskDialog.Show("Inconsistent Scales", 
                "The following sheets have views with different scales and were not updated:\n" + 
                string.Join("\n", inconsistentSheets));
        }

        return Result.Succeeded;
    }
}
