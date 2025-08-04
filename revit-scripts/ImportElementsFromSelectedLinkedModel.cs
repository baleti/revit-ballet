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
            // Get all selected elements
            ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("Error", "Please select one or more linked Revit model instances.");
                return Result.Failed;
            }

            // Filter to get only RevitLinkInstances
            List<RevitLinkInstance> linkedInstances = new List<RevitLinkInstance>();
            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element is RevitLinkInstance linkInstance)
                {
                    linkedInstances.Add(linkInstance);
                }
            }

            if (linkedInstances.Count == 0)
            {
                TaskDialog.Show("Error", "No linked model instances found in selection.");
                return Result.Failed;
            }

            // Collect all types from all selected linked models
            List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
            Dictionary<string, LinkedTypeInfo> typeInfoMap = new Dictionary<string, LinkedTypeInfo>();

            // List of element classes we want to collect
            Type[] typesList = new Type[]
            {
                typeof(FamilySymbol),
                typeof(WallType),
                typeof(FloorType),
                typeof(CeilingType),
                typeof(RoofType),
                typeof(PipeType),
                typeof(DirectShape)
            };

            // Process each linked model
            foreach (RevitLinkInstance linkedInstance in linkedInstances)
            {
                Document linkedDoc = linkedInstance.GetLinkDocument();
                if (linkedDoc == null) continue;

                string linkedModelName = linkedDoc.Title;

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
                            // For system families and others, use the class name as family name
                            family = element.GetType().Name.Replace("Type", "");
                        }

                        string uniqueKey = $"{linkedModelName}|{name}|{family}|{category}";

                        typeEntries.Add(new Dictionary<string, object>
                        {
                            {"Name", name},
                            {"Family", family},
                            {"Category", category},
                            {"Linked Model", linkedModelName}
                        });

                        typeInfoMap[uniqueKey] = new LinkedTypeInfo
                        {
                            TypeId = element.Id,
                            LinkedDoc = linkedDoc,
                            LinkedInstance = linkedInstance
                        };
                    }
                }
            }

            // Display GUI for type selection with linked model info
            List<string> columnHeaders = new List<string> { "Name", "Family", "Category", "Linked Model" };
            List<Dictionary<string, object>> selectedTypes = CustomGUIs.DataGrid(typeEntries, columnHeaders, false);

            // Check if any types were selected
            if (selectedTypes == null || selectedTypes.Count == 0)
            {
                TaskDialog.Show("Error", "No types were selected. Please select at least one type to import.");
                return Result.Failed;
            }

            // Group selected types by linked model
            Dictionary<Document, List<ElementId>> typesByLinkedDoc = new Dictionary<Document, List<ElementId>>();

            foreach (var selectedEntry in selectedTypes)
            {
                string selectedKey = $"{selectedEntry["Linked Model"]}|{selectedEntry["Name"]}|{selectedEntry["Family"]}|{selectedEntry["Category"]}";
                if (typeInfoMap.ContainsKey(selectedKey))
                {
                    LinkedTypeInfo info = typeInfoMap[selectedKey];
                    if (!typesByLinkedDoc.ContainsKey(info.LinkedDoc))
                    {
                        typesByLinkedDoc[info.LinkedDoc] = new List<ElementId>();
                    }
                    typesByLinkedDoc[info.LinkedDoc].Add(info.TypeId);
                }
            }

            // Collect all instances from each linked model
            Dictionary<Document, List<ElementId>> instancesByLinkedDoc = new Dictionary<Document, List<ElementId>>();

            foreach (var kvp in typesByLinkedDoc)
            {
                Document linkedDoc = kvp.Key;
                List<ElementId> selectedTypeIds = kvp.Value;
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

                // Collect DirectShapes
                FilteredElementCollector directShapeCollector = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(DirectShape));
                foreach (DirectShape directShape in directShapeCollector)
                {
                    // For DirectShape, the element itself is used as the "type"
                    if (selectedTypeIds.Contains(directShape.Id))
                    {
                        elementInstanceIds.Add(directShape.Id);
                    }
                }

                if (elementInstanceIds.Count > 0)
                {
                    instancesByLinkedDoc[linkedDoc] = elementInstanceIds;
                }
            }

            // Check if any instances were found
            if (instancesByLinkedDoc.Count == 0 || instancesByLinkedDoc.Values.All(list => list.Count == 0))
            {
                TaskDialog.Show("Error", "No instances found using the selected types.");
                return Result.Failed;
            }

            // Copy the selected elements from all linked models
            using (Transaction transaction = new Transaction(doc, "Import Selected Elements from Linked Models"))
            {
                transaction.Start();
                
                int totalCopied = 0;
                foreach (var kvp in instancesByLinkedDoc)
                {
                    Document sourceDoc = kvp.Key;
                    List<ElementId> instanceIds = kvp.Value;
                    
                    ElementTransformUtils.CopyElements(sourceDoc, instanceIds, doc, null, new CopyPasteOptions());
                    totalCopied += instanceIds.Count;
                }
                
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
    
    // Helper class to store linked type information
    private class LinkedTypeInfo
    {
        public ElementId TypeId { get; set; }
        public Document LinkedDoc { get; set; }
        public RevitLinkInstance LinkedInstance { get; set; }
    }
}
