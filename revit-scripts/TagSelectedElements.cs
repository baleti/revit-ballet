using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class TagSelectedElements : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData, 
      ref string message, 
      ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // Get the selected elements
        ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Error", "Please select elements to tag.");
            return Result.Failed;
        }

        using (Transaction trans = new Transaction(doc, "Place Tags"))
        {
            trans.Start();
            try
            {
                int tagNumber = 1;

                foreach (ElementId id in selectedIds)
                {
                    if (tagNumber > 100)
                    {
                        break;
                    }

                    Element element = doc.GetElement(id);
                    LocationPoint location = element.Location as LocationPoint;
                    XYZ originalPosition = location.Point;
                    BoundingBoxXYZ boundingBox = element.get_BoundingBox(null);

                    if (boundingBox != null)
                    {
                        // Get the current view scale
                        View activeView = doc.ActiveView;
                        int viewScale = activeView.Scale;

                        // Calculate the offset based on the view scale
                        double offsetY = 0.007 * viewScale; // Arbitrary scaling factor

                        // Calculate the tag position with offset
                        XYZ minPoint = boundingBox.Min;
                        XYZ tagPosition = new XYZ(originalPosition.X, minPoint.Y - offsetY, originalPosition.Z);

                        IndependentTag newTag = IndependentTag.Create(doc, doc.ActiveView.Id, new Reference(element), false, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, tagPosition);
                        Parameter typeMarkParam = newTag.LookupParameter("Type Mark");

                        if (typeMarkParam != null && typeMarkParam.IsReadOnly == false)
                        {
                            typeMarkParam.Set(tagNumber.ToString());
                        }

                        tagNumber++;
                    }
                }

                trans.Commit();
            }
            catch (Exception ex)
            {
                message = ex.Message;
                trans.RollBack();
                return Result.Failed;
            }
        }

        return Result.Succeeded;
    }
}
