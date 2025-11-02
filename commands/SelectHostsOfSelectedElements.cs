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
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document  doc   = uidoc.Document;

            // ---------------------------------------------
            // 1) Validate there is a current selection
            // ---------------------------------------------
            ICollection<ElementId> selIds = uidoc.GetSelectionIds();
            if (selIds == null || selIds.Count == 0)
            {
                TaskDialog.Show("Select Hosts", "Please select some elements first.");
                return Result.Failed;
            }

            // ---------------------------------------------
            // 2) Map each selected FamilyInstance -> Host
            // ---------------------------------------------
            var hostToFamilyMap = new Dictionary<ElementId, List<ElementId>>();
            foreach (ElementId id in selIds)
            {
                if (doc.GetElement(id) is FamilyInstance fi && fi.Host != null)
                {
                    ElementId hostId = fi.Host.Id;
                    if (!hostToFamilyMap.ContainsKey(hostId))
                        hostToFamilyMap[hostId] = new List<ElementId>();

                    hostToFamilyMap[hostId].Add(id);
                }
            }

            if (hostToFamilyMap.Count == 0)
            {
                TaskDialog.Show("Select Hosts", "No hosted elements found in the current selection.");
                return Result.Failed;
            }

            // ---------------------------------------------
            // 3) Build host-type metadata
            //    * hostTypeData: unique rows for grid
            //    * hostIdToKey : hostId  -> row-key
            // ---------------------------------------------
            var hostTypeData  = new Dictionary<string, Dictionary<string, object>>();
            var hostIdToKey   = new Dictionary<ElementId, string>();

            foreach (ElementId hostId in hostToFamilyMap.Keys)
            {
                Element host = doc.GetElement(hostId);
                if (host == null) continue;

                ElementId typeId = host.GetTypeId();
                if (typeId == ElementId.InvalidElementId) continue;

                Element typeElem = doc.GetElement(typeId);
                if (typeElem == null) continue;

                string typeName     = typeElem.Name;
                string categoryName = host.Category?.Name ?? "No Category";
                string key          = $"{typeName}|{categoryName}";

                if (!hostTypeData.ContainsKey(key))
                {
                    hostTypeData[key] = new Dictionary<string, object>
                    {
                        { "Type",     typeName     },
                        { "Category", categoryName }
                    };
                }

                hostIdToKey[hostId] = key;
            }

            // ---------------------------------------------
            // 4) Show grid (multi-select enabled)
            // ---------------------------------------------
            var rows    = hostTypeData.Values.ToList();
            var columns = new List<string> { "Type", "Category" };

            // *** 3rd arg = true → allow multi-row selection ***
            var selectedRows = CustomGUIs.DataGrid(rows, columns, false);

            if (selectedRows == null || selectedRows.Count == 0)
            {
                return Result.Cancelled;
            }

            // ---------------------------------------------
            // 5) Collect ALL keys chosen in the grid
            // ---------------------------------------------
            var selectedKeys = new HashSet<string>();
            foreach (var row in selectedRows)
            {
                string t = row["Type"].ToString();
                string c = row["Category"].ToString();
                selectedKeys.Add($"{t}|{c}");
            }

            // ---------------------------------------------
            // 6) Build final selection set
            //    • each matching host
            //    • each originally selected family instance
            // ---------------------------------------------
            var finalSelection = new HashSet<ElementId>();

            foreach (var kvp in hostIdToKey)
            {
                ElementId hostId = kvp.Key;
                string    key    = kvp.Value;

                if (!selectedKeys.Contains(key)) continue; // skip non-chosen types

                finalSelection.Add(hostId);                // add host

                if (hostToFamilyMap.TryGetValue(hostId, out var fiIds))
                    foreach (var fiId in fiIds)
                        finalSelection.Add(fiId);           // add its original FIs
            }

            // ---------------------------------------------
            // 7) Update Revit’s active selection
            // ---------------------------------------------
            uidoc.SetSelectionIds(finalSelection.ToList());

            return Result.Succeeded;
        }
    }
}
