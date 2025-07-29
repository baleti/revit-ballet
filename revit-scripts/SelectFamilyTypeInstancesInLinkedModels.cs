using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;
using System;

[Transaction(TransactionMode.Manual)]
public class SelectFamilyTypeInstancesInLinkedModels : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get currently selected RevitLinkInstances
        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
        List<RevitLinkInstance> selectedLinks = new List<RevitLinkInstance>();
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem is RevitLinkInstance linkInstance && linkInstance.GetLinkDocument() != null)
            {
                selectedLinks.Add(linkInstance);
            }
        }

        if (selectedLinks.Count == 0)
        {
            message = "No linked models selected.";
            return Result.Failed;
        }

        // Step 1: Prepare a list of types across all categories in the selected linked models
        List<Dictionary<string, object>> typeEntries = new List<Dictionary<string, object>>();
        Dictionary<string, Tuple<RevitLinkInstance, FamilySymbol>> typeElementMap = new Dictionary<string, Tuple<RevitLinkInstance, FamilySymbol>>();

        // Collect FamilySymbols from each selected linked model
        foreach (var link in selectedLinks)
        {
            Document linkedDoc = link.GetLinkDocument();
            string linkName = link.Name; // Or linkedDoc.Title if preferred

            var familySymbols = new FilteredElementCollector(linkedDoc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            // Iterate through the FamilySymbols, and collect their info
            foreach (var familySymbol in familySymbols)
            {
                // Get the family for the current symbol
                Family family = familySymbol.Family;

                var entry = new Dictionary<string, object>
                {
                    { "Type Name", familySymbol.Name },
                    { "Family", family.Name },
                    { "Category", familySymbol.Category.Name },
                    { "Link", linkName }
                };

                // Store the FamilySymbol with a unique key (Link:Family:Type)
                string uniqueKey = $"{linkName}:{family.Name}:{familySymbol.Name}";
                typeElementMap[uniqueKey] = new Tuple<RevitLinkInstance, FamilySymbol>(link, familySymbol);

                typeEntries.Add(entry);
            }
        }

        // Step 2: Display the list of types using CustomGUIs.DataGrid
        var propertyNames = new List<string> { "Type Name", "Family", "Category", "Link" };
        var selectedEntries = CustomGUIs.DataGrid(typeEntries, propertyNames, spanAllScreens: false);

        if (selectedEntries.Count == 0)
        {
            return Result.Cancelled; // No selection made
        }

        // Step 3: Collect References of the instances of the selected types
        List<Reference> selectedRefs = new List<Reference>();

        foreach (var entry in selectedEntries)
        {
            string uniqueKey = $"{entry["Link"]}:{entry["Family"]}:{entry["Type Name"]}";
            var tuple = typeElementMap[uniqueKey];
            RevitLinkInstance link = tuple.Item1;
            FamilySymbol symbol = tuple.Item2;
            Document linkedDoc = link.GetLinkDocument();

            // Collect all instances of the selected type in the linked model
            var instances = new FilteredElementCollector(linkedDoc)
                .WherePasses(new ElementMulticlassFilter(new List<System.Type> { typeof(FamilyInstance), typeof(ElementType) }))
                .Where(x => x.GetTypeId() == symbol.Id)
                .ToList();

            // Create References for each instance
            foreach (var instance in instances)
            {
                Reference elemRef = new Reference(instance);
                Reference reference = elemRef.CreateLinkReference(link);
                selectedRefs.Add(reference);
            }
        }

        // Step 4: Add the new selection to the existing selection
        IList<Reference> currentRefs = uidoc.GetReferences();
        List<Reference> combinedRefs = new List<Reference>(currentRefs);
        HashSet<string> currentStable = new HashSet<string>();
        foreach (var r in currentRefs)
        {
            currentStable.Add(r.ConvertToStableRepresentation(doc));
        }

        foreach (var newRef in selectedRefs)
        {
            string stable = newRef.ConvertToStableRepresentation(doc);
            if (!currentStable.Contains(stable))
            {
                combinedRefs.Add(newRef);
            }
        }

        // Update the selection with both previous and newly selected references
        uidoc.SetReferences(combinedRefs);

        return Result.Succeeded;
    }
}
