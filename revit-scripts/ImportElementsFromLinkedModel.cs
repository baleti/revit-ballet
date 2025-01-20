using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class ImportElementsFromLinkedModel : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Get the Revit application and document
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        try
        {
            // Get the linked Revit model (instance)
            ElementId linkedInstanceId = uiDoc.Selection.GetElementIds().FirstOrDefault();
            if (linkedInstanceId == null)
            {
                TaskDialog.Show("Error", "Please select a linked Revit model instance.");
                return Result.Failed;
            }

            Element linkedElement = doc.GetElement(linkedInstanceId);
            RevitLinkInstance linkedInstance = linkedElement as RevitLinkInstance;
            if (linkedInstance == null)
            {
                TaskDialog.Show("Error", "Selected element is not a linked model instance.");
                return Result.Failed;
            }

            Document linkedDoc = linkedInstance.GetLinkDocument();

            // Collect all family types (FamilySymbol) in the linked document
            FilteredElementCollector typeCollector = new FilteredElementCollector(linkedDoc)
                .OfClass(typeof(FamilySymbol));

            // Create a list of available family types and a dictionary to map them to their ElementIds
            List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
            Dictionary<string, ElementId> typeIdMap = new Dictionary<string, ElementId>();

            foreach (FamilySymbol familySymbol in typeCollector)
            {
                string uniqueKey = $"{familySymbol.Name}|{familySymbol.Family.Name}|{familySymbol.Category.Name}";

                // Store the properties to display in the selection
                typeEntries.Add(new Dictionary<string, object>
                {
                    {"Name", familySymbol.Name},
                    {"Family", familySymbol.Family.Name},
                    {"Category", familySymbol.Category.Name},
                });

                // Store the corresponding ElementId (FamilySymbol) with a unique key
                typeIdMap[uniqueKey] = familySymbol.Id;
            }

            // Display a GUI to let the user select which family types to import
            List<string> columnHeaders = new List<string> { "Name", "Family", "Category" };
            List<Dictionary<string, object>> selectedTypes = CustomGUIs.DataGrid(typeEntries, columnHeaders, false);

            // Prepare a list of selected FamilySymbol ElementIds
            List<ElementId> selectedTypeIds = new List<ElementId>();

            foreach (var selectedEntry in selectedTypes)
            {
                string selectedKey = $"{selectedEntry["Name"]}|{selectedEntry["Family"]}|{selectedEntry["Category"]}";

                if (typeIdMap.ContainsKey(selectedKey))
                {
                    selectedTypeIds.Add(typeIdMap[selectedKey]);
                }
            }

            // Now collect element instances (FamilyInstance) based on selected types
            FilteredElementCollector elementCollector = new FilteredElementCollector(linkedDoc)
                .OfClass(typeof(FamilyInstance));

            // List of element IDs (instances) to be copied
            List<ElementId> elementInstanceIds = new List<ElementId>();

            foreach (FamilyInstance familyInstance in elementCollector)
            {
                if (selectedTypeIds.Contains(familyInstance.Symbol.Id)) // Check if this instance's type is selected
                {
                    elementInstanceIds.Add(familyInstance.Id); // Add instance to the list
                }
            }

            // Copy the selected element instances from the linked model into the current model
            using (Transaction transaction = new Transaction(doc, "Import Selected Elements from Linked Model"))
            {
                transaction.Start();

                // Copy element instances into the current model
                ElementTransformUtils.CopyElements(linkedDoc, elementInstanceIds, doc, null, new CopyPasteOptions());

                transaction.Commit();
            }
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", ex.Message);
            return Result.Failed;
        }

        return Result.Succeeded;
    }
}
