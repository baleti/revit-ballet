using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace YourNamespace
{
    [Transaction(TransactionMode.Manual)]
    public class SelectHostsOfSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the active document and selection.
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
            if (selIds == null || selIds.Count == 0)
            {
                TaskDialog.Show("Select Hosts", "Please select some elements first.");
                return Result.Failed;
            }

            // Map each selected family instance to its host.
            // This dictionary maps a host element Id to a list of original selected family instance Ids.
            Dictionary<ElementId, List<ElementId>> hostToFamilyMap = new Dictionary<ElementId, List<ElementId>>();
            foreach (ElementId id in selIds)
            {
                Element elem = doc.GetElement(id);
                FamilyInstance fi = elem as FamilyInstance;
                if (fi != null && fi.Host != null)
                {
                    ElementId hostId = fi.Host.Id;
                    if (!hostToFamilyMap.ContainsKey(hostId))
                    {
                        hostToFamilyMap[hostId] = new List<ElementId>();
                    }
                    hostToFamilyMap[hostId].Add(id);
                }
            }

            if (hostToFamilyMap.Count == 0)
            {
                TaskDialog.Show("Select Hosts", "No hosted elements found in the current selection.");
                return Result.Failed;
            }

            // Group host elements by their type and category.
            // We'll create a unique key from the type name and category.
            Dictionary<string, Dictionary<string, object>> hostTypeData = new Dictionary<string, Dictionary<string, object>>();
            // Also keep a mapping from host Id to its type key.
            Dictionary<ElementId, string> hostIdToTypeKey = new Dictionary<ElementId, string>();
            foreach (var kvp in hostToFamilyMap)
            {
                ElementId hostId = kvp.Key;
                Element host = doc.GetElement(hostId);
                if (host == null)
                    continue;
                ElementId typeId = host.GetTypeId();
                if (typeId == ElementId.InvalidElementId)
                    continue;
                Element typeElem = doc.GetElement(typeId);
                if (typeElem != null)
                {
                    string typeName = typeElem.Name;
                    string categoryName = host.Category != null ? host.Category.Name : "No Category";
                    string key = typeName + "|" + categoryName;
                    if (!hostTypeData.ContainsKey(key))
                    {
                        hostTypeData[key] = new Dictionary<string, object>
                        {
                            { "Type", typeName },
                            { "Category", categoryName }
                        };
                    }
                    hostIdToTypeKey[hostId] = key;
                }
            }

            // Prepare data for the DataGrid.
            List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();
            foreach (var entry in hostTypeData.Values)
            {
                data.Add(entry);
            }

            // Define the columns to be displayed.
            List<string> columns = new List<string> { "Type", "Category" };

            // Display the DataGrid to the user. The third parameter 'false' indicates single selection.
            var selectedRows = CustomGUIs.DataGrid(data, columns, false);
            if (selectedRows == null || selectedRows.Count == 0)
            {
                TaskDialog.Show("Select Hosts", "No host type was selected.");
                return Result.Cancelled;
            }

            // Assume single selection: retrieve the chosen type and category.
            Dictionary<string, object> selectedRow = selectedRows[0];
            string selType = selectedRow["Type"].ToString();
            string selCategory = selectedRow["Category"].ToString();
            string selectedKey = selType + "|" + selCategory;

            // Build the final selection.
            // Include:
            // 1. The host elements of the chosen type.
            // 2. The original family instances whose host is of that type.
            HashSet<ElementId> finalSelection = new HashSet<ElementId>();
            foreach (var kvp in hostIdToTypeKey)
            {
                ElementId hostId = kvp.Key;
                string key = kvp.Value;
                if (key == selectedKey)
                {
                    // Add the host element.
                    finalSelection.Add(hostId);
                    // Add the associated original family instances.
                    if (hostToFamilyMap.ContainsKey(hostId))
                    {
                        foreach (ElementId fiId in hostToFamilyMap[hostId])
                        {
                            finalSelection.Add(fiId);
                        }
                    }
                }
            }

            // Update the current selection.
            uidoc.Selection.SetElementIds(finalSelection.ToList());
            TaskDialog.Show("Select Hosts", $"Selection updated with {finalSelection.Count} element(s).");

            return Result.Succeeded;
        }
    }
}
