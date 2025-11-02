using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class SetInstanceParameterToNumberSequenceClockwise : IExternalCommand
{
    public class ParameterInfo
    {
        public string Name { get; set; }
    }

    private enum EdgeType { Top = 0, Right = 1, Bottom = 2, Left = 3 }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "No elements selected.");
            return Result.Cancelled;
        }

        var elementsToProcess = new List<Element>();
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem.get_BoundingBox(doc.ActiveView) != null)
                elementsToProcess.Add(elem);
        }

        if (elementsToProcess.Count == 0)
        {
            TaskDialog.Show("Error", "No elements with valid bounding boxes selected.");
            return Result.Cancelled;
        }

        // Compute overall bounding box
        BoundingBoxXYZ overallBB = GetOverallBoundingBox(elementsToProcess, doc);

        // Prepare list for sorting
        var elementSortInfos = new List<ElementSortInfo>();
        foreach (Element elem in elementsToProcess)
        {
            XYZ center = GetBoundingBoxCenter(elem, doc);
            double ex = center.X;
            double ey = center.Y;

            // Calculate distances to each edge
            double topDistance = overallBB.Max.Y - ey;
            double rightDistance = overallBB.Max.X - ex;
            double bottomDistance = ey - overallBB.Min.Y;
            double leftDistance = ex - overallBB.Min.X;

            double minDistance = Math.Min(Math.Min(topDistance, rightDistance), Math.Min(bottomDistance, leftDistance));

            EdgeType edge;
            // First, check if it's closest to top or left edge for the first element
            if (Math.Abs(topDistance - minDistance) < 1e-6 && Math.Abs(leftDistance - minDistance) < 1e-6)
            {
                edge = EdgeType.Top; // This will be our starting point
            }
            else if (minDistance == topDistance)
                edge = EdgeType.Top;
            else if (minDistance == rightDistance)
                edge = EdgeType.Right;
            else if (minDistance == bottomDistance)
                edge = EdgeType.Bottom;
            else
                edge = EdgeType.Left;

            double parameter = 0.0;
            double width = overallBB.Max.X - overallBB.Min.X;
            double height = overallBB.Max.Y - overallBB.Min.Y;

            switch (edge)
            {
                case EdgeType.Top:
                    parameter = width == 0 ? 0 : (ex - overallBB.Min.X) / width;
                    break;
                case EdgeType.Right:
                    parameter = height == 0 ? 0 : (overallBB.Max.Y - ey) / height;
                    break;
                case EdgeType.Bottom:
                    parameter = width == 0 ? 0 : (overallBB.Max.X - ex) / width;
                    break;
                case EdgeType.Left:
                    parameter = height == 0 ? 0 : (ey - overallBB.Min.Y) / height;
                    break;
            }

            // For top edge, we want left-to-right ordering
            if (edge == EdgeType.Top)
                parameter = ex - overallBB.Min.X;
            
            elementSortInfos.Add(new ElementSortInfo
            {
                Element = elem,
                EdgeOrder = (int)edge,
                Parameter = parameter,
                DistanceToTopLeft = Math.Sqrt(
                    Math.Pow(ex - overallBB.Min.X, 2) + 
                    Math.Pow(overallBB.Max.Y - ey, 2))
            });
        }

        // Modified sorting logic:
        // 1. First, find the element closest to top-left corner
        var topLeftElement = elementSortInfos
            .OrderBy(esi => esi.DistanceToTopLeft)
            .First();

        // 2. Remove it from the list and add it back as the first element
        elementSortInfos.Remove(topLeftElement);
        
        // 3. Sort remaining elements
        var remainingSortedElements = elementSortInfos
            .OrderBy(esi => esi.EdgeOrder)
            .ThenBy(esi => esi.Parameter)
            .ToList();

        // 4. Combine the lists
        var finalSortedElements = new List<Element> { topLeftElement.Element };
        finalSortedElements.AddRange(remainingSortedElements.Select(esi => esi.Element));

        // Collect parameters and set values
        var paramNames = CollectStringParameters(elementsToProcess);
        if (paramNames.Count == 0)
            return Result.Cancelled;

        var selectedParams = GetUserSelectedParameters(paramNames);
        if (selectedParams.Count == 0)
            return Result.Cancelled;

        int paddingLength = finalSortedElements.Count.ToString().Length;

        using (Transaction t = new Transaction(doc, "Set Clockwise Number Sequence"))
        {
            t.Start();
            SetParameters(finalSortedElements, selectedParams, paddingLength);
            t.Commit();
        }

        return Result.Succeeded;
    }

    private BoundingBoxXYZ GetOverallBoundingBox(List<Element> elements, Document doc)
    {
        BoundingBoxXYZ overallBB = null;
        foreach (Element elem in elements)
        {
            BoundingBoxXYZ elemBB = elem.get_BoundingBox(doc.ActiveView);
            if (overallBB == null)
            {
                overallBB = elemBB;
            }
            else
            {
                overallBB.Min = new XYZ(
                    Math.Min(overallBB.Min.X, elemBB.Min.X),
                    Math.Min(overallBB.Min.Y, elemBB.Min.Y),
                    Math.Min(overallBB.Min.Z, elemBB.Min.Z));
                overallBB.Max = new XYZ(
                    Math.Max(overallBB.Max.X, elemBB.Max.X),
                    Math.Max(overallBB.Max.Y, elemBB.Max.Y),
                    Math.Max(overallBB.Max.Z, elemBB.Max.Z));
            }
        }
        return overallBB;
    }

    private XYZ GetBoundingBoxCenter(Element element, Document doc)
    {
        BoundingBoxXYZ bb = element.get_BoundingBox(doc.ActiveView);
        return (bb.Min + bb.Max) / 2;
    }

    private HashSet<string> CollectStringParameters(List<Element> elements)
    {
        var paramNames = new HashSet<string>();
        foreach (Element elem in elements)
        {
            foreach (Parameter param in elem.Parameters)
            {
                if (!param.IsShared && !param.IsReadOnly && param.StorageType == StorageType.String)
                {
                    paramNames.Add(param.Definition.Name);
                }
            }
        }
        return paramNames;
    }

    private List<string> GetUserSelectedParameters(HashSet<string> paramNames)
    {
        var paramList = paramNames.Select(n => new ParameterInfo { Name = n }).ToList();
        var selectedParams = CustomGUIs.DataGrid(
            paramList,
            new List<string> { "Name" },
            null,
            "Select Parameters to Number"
        );
        return selectedParams.Select(p => p.Name).ToList();
    }

    private void SetParameters(List<Element> sortedElements, List<string> selectedParamNames, int paddingLength)
    {
        int counter = 1;
        foreach (Element elem in sortedElements)
        {
            string formattedNumber = counter.ToString().PadLeft(paddingLength, '0');
            foreach (string paramName in selectedParamNames)
            {
                Parameter param = elem.Parameters.Cast<Parameter>()
                    .FirstOrDefault(p => p.Definition.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase)
                        && !p.IsShared
                        && !p.IsReadOnly
                        && p.StorageType == StorageType.String);
                param?.Set(formattedNumber);
            }
            counter++;
        }
    }

    private class ElementSortInfo
    {
        public Element Element { get; set; }
        public int EdgeOrder { get; set; }
        public double Parameter { get; set; }
        public double DistanceToTopLeft { get; set; }
    }
}
