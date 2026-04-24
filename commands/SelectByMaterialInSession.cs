using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("")]
public class SelectByMaterialInSession : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        var app = uiapp.Application;

        // material name -> MaterialInfo (tracks across documents)
        var materialInfoMap = new Dictionary<string, MaterialInfo>();

        var documentTitles = new List<string>();
        foreach (Document doc in app.Documents)
        {
            if (doc.IsLinked) continue;
            documentTitles.Add(doc.Title);

            var collector = new FilteredElementCollector(doc).WhereElementIsNotElementType();
            foreach (Element elem in collector)
            {
                ICollection<ElementId> matIds;
                try { matIds = elem.GetMaterialIds(false); } catch { continue; }

                ICollection<ElementId> paintIds;
                try { paintIds = elem.GetMaterialIds(true); } catch { paintIds = new List<ElementId>(); }

                var allMatIds = matIds.Union(paintIds).Distinct();
                foreach (ElementId matId in allMatIds)
                {
                    Material mat = doc.GetElement(matId) as Material;
                    if (mat == null) continue;

                    string key = mat.Name;
                    if (!materialInfoMap.ContainsKey(key))
                    {
                        materialInfoMap[key] = new MaterialInfo
                        {
                            MaterialName = mat.Name,
                            MaterialClass = mat.MaterialClass ?? ""
                        };
                    }
                    materialInfoMap[key].AddElement(doc, elem.Id, elem.UniqueId);
                }
            }
        }

        if (materialInfoMap.Count == 0)
        {
            TaskDialog.Show("No Materials", "No materials found in open documents.");
            return Result.Cancelled;
        }

        // Build DataGrid rows
        var materialList = new List<Dictionary<string, object>>();
        foreach (var info in materialInfoMap.Values.OrderBy(m => m.MaterialName))
        {
            var entry = new Dictionary<string, object>
            {
                { "Name", info.MaterialName },
                { "Material Class", info.MaterialClass },
                { "_MaterialInfo", info }
            };
            foreach (string docTitle in documentTitles)
                entry[docTitle] = info.ElementsByDocument.ContainsKey(docTitle) ? info.ElementsByDocument[docTitle].Count : 0;

            materialList.Add(entry);
        }

        var propertyNames = new List<string> { "Name", "Material Class" };
        propertyNames.AddRange(documentTitles);

        List<Dictionary<string, object>> selected = CustomGUIs.DataGrid(materialList, propertyNames, false);
        if (selected.Count == 0)
            return Result.Cancelled;

        // Build SelectionItems and merge into SelectionStorage
        var selectionItems = new List<SelectionItem>();
        foreach (var row in selected)
        {
            var info = (MaterialInfo)row["_MaterialInfo"];
            foreach (var docEntry in info.ElementsByDocument)
            {
                string docTitle = docEntry.Key;
                string docPath = GetDocumentPath(app, docTitle);
                foreach (var elemEntry in docEntry.Value)
                {
                    selectionItems.Add(new SelectionItem
                    {
                        DocumentTitle = docTitle,
                        DocumentPath = docPath,
                        UniqueId = elemEntry.Value,
#if REVIT2024 || REVIT2025 || REVIT2026
                        ElementIdValue = (int)elemEntry.Key.Value,
#else
                        ElementIdValue = elemEntry.Key.IntegerValue,
#endif
                        SessionId = null
                    });
                }
            }
        }

        var existingSelection = SelectionStorage.LoadSelection();
        var existingKeys = new HashSet<string>(existingSelection.Select(s => $"{s.DocumentTitle}|{s.UniqueId}"));
        foreach (var item in selectionItems)
        {
            if (!existingKeys.Contains($"{item.DocumentTitle}|{item.UniqueId}"))
                existingSelection.Add(item);
        }
        SelectionStorage.SaveSelection(existingSelection);

        return Result.Succeeded;
    }

    private string GetDocumentPath(Autodesk.Revit.ApplicationServices.Application app, string documentTitle)
    {
        foreach (Document doc in app.Documents)
        {
            if (doc.Title == documentTitle)
                return doc.PathName ?? documentTitle;
        }
        return documentTitle;
    }

    private class MaterialInfo
    {
        public string MaterialName { get; set; }
        public string MaterialClass { get; set; }

        // Document Title -> (ElementId -> UniqueId)
        public Dictionary<string, Dictionary<ElementId, string>> ElementsByDocument { get; set; }
            = new Dictionary<string, Dictionary<ElementId, string>>();

        public void AddElement(Document doc, ElementId elementId, string uniqueId)
        {
            string docTitle = doc.Title;
            if (!ElementsByDocument.ContainsKey(docTitle))
                ElementsByDocument[docTitle] = new Dictionary<ElementId, string>();
            ElementsByDocument[docTitle][elementId] = uniqueId;
        }
    }
}
