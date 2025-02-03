using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class SetInstanceParameterToNumberSequenceAlongX : IExternalCommand
{
    public class ParameterInfo
    {
        public string Name { get; set; }
    }
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get selected elements
        ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "No elements selected.");
            return Result.Cancelled;
        }

        // Collect elements with valid bounding boxes
        var elementsToProcess = new List<Element>();
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            BoundingBoxXYZ bb = elem.get_BoundingBox(doc.ActiveView);
            if (bb != null) elementsToProcess.Add(elem);
        }

        // Collect all writable string instance parameters
        var paramNames = new HashSet<string>();
        foreach (Element elem in elementsToProcess)
        {
            foreach (Parameter param in elem.Parameters)
            {
                if (!param.IsShared && 
                    !param.IsReadOnly && 
                    param.StorageType == StorageType.String)
                {
                    paramNames.Add(param.Definition.Name);
                }
            }
        }

        if (paramNames.Count == 0)
        {
            TaskDialog.Show("Error", "No writable string instance parameters found.");
            return Result.Cancelled;
        }

        // Let user select parameters
        var paramList = paramNames.Select(n => new ParameterInfo { Name = n }).ToList();
        var selectedParams = CustomGUIs.DataGrid(
            paramList, 
            new List<string> { "Name" }, 
            null, 
            "Select Parameters to Number"
        );

        if (selectedParams.Count == 0)
        {
            TaskDialog.Show("Info", "No parameters selected.");
            return Result.Cancelled;
        }
        var selectedParamNames = selectedParams.Select(p => p.Name).ToList();

        // Sort elements by X position
        var sortedElements = elementsToProcess
            .OrderBy(e => GetBoundingBoxCenterX(e, doc))
            .ToList();

        // Set parameters in transaction
        using (Transaction t = new Transaction(doc, "Set Number Sequence"))
        {
            t.Start();

            int counter = 1;
            foreach (Element elem in sortedElements)
            {
                foreach (string paramName in selectedParamNames)
                {
                    Parameter param = elem.Parameters.Cast<Parameter>()
                        .FirstOrDefault(p => p.Definition.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase) 
                            && !p.IsShared 
                            && !p.IsReadOnly 
                            && p.StorageType == StorageType.String);

                    if (param != null)
                    {
                        param.Set(counter.ToString());
                    }
                }
                counter++;
            }

            t.Commit();
        }

        return Result.Succeeded;
    }

    private double GetBoundingBoxCenterX(Element element, Document doc)
    {
        BoundingBoxXYZ bb = element.get_BoundingBox(doc.ActiveView);
        return (bb.Min.X + bb.Max.X) / 2.0;
    }
}
