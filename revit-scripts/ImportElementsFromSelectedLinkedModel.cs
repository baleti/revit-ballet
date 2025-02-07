using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class ImportElementsFromSelectedLinkedModel : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        try
        {
            // Get the linked Revit model
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

            // Collect all types that we want to import
            List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
            Dictionary<string, ElementId> typeIdMap = new Dictionary<string, ElementId>();

            // List of element classes we want to collect
            Type[] typesList = new Type[]
            {
                typeof(FamilySymbol),
                typeof(WallType),
                typeof(FloorType),
                typeof(CeilingType),
                typeof(RoofType),
                typeof(PipeType)
            };

            foreach (Type elementType in typesList)
            {
                FilteredElementCollector typeCollector = new FilteredElementCollector(linkedDoc)
                    .OfClass(elementType);

                foreach (Element element in typeCollector)
                {
                    string name = element.Name;
                    string category = element.Category?.Name ?? "Uncategorized";
                    string family = "";

                    // Handle different type of elements
                    if (element is FamilySymbol familySymbol)
                    {
                        family = familySymbol.Family.Name;
                    }
                    else
                    {
                        // For system families, use the class name as family name
                        family = element.GetType().Name.Replace("Type", "");
                    }

                    string uniqueKey = $"{name}|{family}|{category}";

                    typeEntries.Add(new Dictionary<string, object>
                    {
                        {"Name", name},
                        {"Family", family},
                        {"Category", category},
                    });

                    typeIdMap[uniqueKey] = element.Id;
                }
            }

            // Display GUI for type selection
            List<string> columnHeaders = new List<string> { "Name", "Family", "Category" };
            List<Dictionary<string, object>> selectedTypes = CustomGUIs.DataGrid(typeEntries, columnHeaders, false);

            // Check if any types were selected
            if (selectedTypes == null || selectedTypes.Count == 0)
            {
                TaskDialog.Show("Error", "No types were selected. Please select at least one type to import.");
                return Result.Failed;
            }

            // Get selected type IDs
            List<ElementId> selectedTypeIds = new List<ElementId>();
            foreach (var selectedEntry in selectedTypes)
            {
                string selectedKey = $"{selectedEntry["Name"]}|{selectedEntry["Family"]}|{selectedEntry["Category"]}";
                if (typeIdMap.ContainsKey(selectedKey))
                {
                    selectedTypeIds.Add(typeIdMap[selectedKey]);
                }
            }

            // Collect all instances that use the selected types
            List<ElementId> elementInstanceIds = new List<ElementId>();

            // Collect FamilyInstances
            FilteredElementCollector familyInstanceCollector = new FilteredElementCollector(linkedDoc)
                .OfClass(typeof(FamilyInstance));
            foreach (FamilyInstance instance in familyInstanceCollector)
            {
                if (selectedTypeIds.Contains(instance.Symbol.Id))
                {
                    elementInstanceIds.Add(instance.Id);
                }
            }

            // Collect Walls
            FilteredElementCollector wallCollector = new FilteredElementCollector(linkedDoc)
                .OfClass(typeof(Wall));
            foreach (Wall wall in wallCollector)
            {
                if (selectedTypeIds.Contains(wall.WallType.Id))
                {
                    elementInstanceIds.Add(wall.Id);
                }
            }

            // Collect Floors
            FilteredElementCollector floorCollector = new FilteredElementCollector(linkedDoc)
                .OfClass(typeof(Floor));
            foreach (Floor floor in floorCollector)
            {
                if (selectedTypeIds.Contains(floor.FloorType.Id))
                {
                    elementInstanceIds.Add(floor.Id);
                }
            }

            // Collect Ceilings
            FilteredElementCollector ceilingCollector = new FilteredElementCollector(linkedDoc)
                .OfClass(typeof(Ceiling));
            foreach (Ceiling ceiling in ceilingCollector)
            {
                if (selectedTypeIds.Contains(ceiling.GetTypeId()))
                {
                    elementInstanceIds.Add(ceiling.Id);
                }
            }

            // Collect Roofs
            FilteredElementCollector roofCollector = new FilteredElementCollector(linkedDoc)
                .OfClass(typeof(RoofBase));
            foreach (RoofBase roof in roofCollector)
            {
                if (selectedTypeIds.Contains(roof.RoofType.Id))
                {
                    elementInstanceIds.Add(roof.Id);
                }
            }

            // Collect Pipes
            FilteredElementCollector pipeCollector = new FilteredElementCollector(linkedDoc)
                .OfClass(typeof(Pipe));
            foreach (Pipe pipe in pipeCollector)
            {
                if (selectedTypeIds.Contains(pipe.PipeType.Id))
                {
                    elementInstanceIds.Add(pipe.Id);
                }
            }

            // Check if any instances were found
            if (elementInstanceIds.Count == 0)
            {
                TaskDialog.Show("Error", "No instances found using the selected types.");
                return Result.Failed;
            }

            // Copy the selected elements
            using (Transaction transaction = new Transaction(doc, "Import Selected Elements from Linked Model"))
            {
                transaction.Start();
                ElementTransformUtils.CopyElements(linkedDoc, elementInstanceIds, doc, null, new CopyPasteOptions());
                transaction.Commit();
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", ex.Message);
            return Result.Failed;
        }
    }
}
