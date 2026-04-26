using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitBallet.Commands;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
[CommandMeta("")]
public class SelectByCategoriesInSession : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
        Document activeDoc = uiapp.ActiveUIDocument.Document;

        // Build category counts across ALL open documents
        Dictionary<ElementId, CategoryInfo> categoryInfoMap = new Dictionary<ElementId, CategoryInfo>();

#if REVIT2024 || REVIT2025 || REVIT2026
        ElementId filterSentinelId      = new ElementId(-3000001L);
        ElementId sharedParamSentinelId = new ElementId(-3000002L);
        ElementId globalParamSentinelId = new ElementId(-3000003L);
#else
        ElementId filterSentinelId      = new ElementId(-3000001);
        ElementId sharedParamSentinelId = new ElementId(-3000002);
        ElementId globalParamSentinelId = new ElementId(-3000003);
#endif

        foreach (Document doc in app.Documents)
        {
            if (doc.IsLinked) continue; // Skip linked documents

            // Collect all sheets
            FilteredElementCollector sheetCollector = new FilteredElementCollector(doc);
            foreach (ViewSheet sheet in sheetCollector.OfClass(typeof(ViewSheet)))
            {
                ElementId categoryId = ((long)BuiltInCategory.OST_Sheets).ToElementId();
                if (!categoryInfoMap.ContainsKey(categoryId))
                {
                    categoryInfoMap[categoryId] = new CategoryInfo
                    {
                        CategoryId = categoryId,
                        CategoryName = "Sheets",
                        IsSheet = true
                    };
                }
                categoryInfoMap[categoryId].AddElement(doc, sheet.Id, sheet.UniqueId, false);
            }

            // Collect all views and view templates
            FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
            foreach (View view in viewCollector.OfClass(typeof(View)))
            {
                if (view is ViewSheet) continue; // Already counted

                ElementId categoryId = ((long)BuiltInCategory.OST_Viewers).ToElementId();

                if (view.IsTemplate)
                {
                    if (!categoryInfoMap.ContainsKey(categoryId))
                    {
                        categoryInfoMap[categoryId] = new CategoryInfo
                        {
                            CategoryId = categoryId,
                            CategoryName = "View Templates",
                            IsViewTemplate = true
                        };
                    }
                    categoryInfoMap[categoryId].AddElement(doc, view.Id, view.UniqueId, false);
                }
                else
                {
                    if (!categoryInfoMap.ContainsKey(categoryId))
                    {
                        categoryInfoMap[categoryId] = new CategoryInfo
                        {
                            CategoryId = categoryId,
                            CategoryName = "Views",
                            IsView = true
                        };
                    }
                    categoryInfoMap[categoryId].AddElement(doc, view.Id, view.UniqueId, false);
                }
            }

            // Collect view filters (ParameterFilterElement — project-level, no category)
            foreach (ParameterFilterElement pfe in new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement)).Cast<ParameterFilterElement>())
            {
                if (!categoryInfoMap.ContainsKey(filterSentinelId))
                {
                    categoryInfoMap[filterSentinelId] = new CategoryInfo
                    {
                        CategoryId = filterSentinelId,
                        CategoryName = "View Filters",
                        IsViewFilter = true
                    };
                }
                categoryInfoMap[filterSentinelId].AddElement(doc, pfe.Id, pfe.UniqueId, false);
            }

            foreach (SharedParameterElement sp in new FilteredElementCollector(doc)
                .OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>())
            {
                if (!categoryInfoMap.ContainsKey(sharedParamSentinelId))
                    categoryInfoMap[sharedParamSentinelId] = new CategoryInfo
                        { CategoryId = sharedParamSentinelId, CategoryName = "Shared Parameters" };
                categoryInfoMap[sharedParamSentinelId].AddElement(doc, sp.Id, sp.UniqueId, false);
            }

            foreach (GlobalParameter gp in new FilteredElementCollector(doc)
                .OfClass(typeof(GlobalParameter)).Cast<GlobalParameter>())
            {
                if (!categoryInfoMap.ContainsKey(globalParamSentinelId))
                    categoryInfoMap[globalParamSentinelId] = new CategoryInfo
                        { CategoryId = globalParamSentinelId, CategoryName = "Global Parameters" };
                categoryInfoMap[globalParamSentinelId].AddElement(doc, gp.Id, gp.UniqueId, false);
            }

            // Collect all other elements
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.WhereElementIsNotElementType();

            foreach (Element elem in collector)
            {
                if (elem is View) continue; // Already processed

                Category category = elem.Category;
                if (category == null) continue;

                ElementId categoryId = category.Id;

                // Skip OST_Viewers for non-view elements
                if (categoryId.AsLong() == (int)BuiltInCategory.OST_Viewers)
                    continue;

                bool isDirectShape = elem is DirectShape;

                if (!categoryInfoMap.ContainsKey(categoryId))
                {
                    categoryInfoMap[categoryId] = new CategoryInfo
                    {
                        CategoryId = categoryId,
                        CategoryName = category.Name,
                        IsDirectShape = isDirectShape
                    };
                }

                categoryInfoMap[categoryId].AddElement(doc, elem.Id, elem.UniqueId, isDirectShape);
            }
        }

        // Get list of all open document titles (non-linked only)
        var documentTitles = new List<string>();
        foreach (Document doc in app.Documents)
        {
            if (!doc.IsLinked)
            {
                documentTitles.Add(doc.Title);
            }
        }

        // Convert to list for DataGrid with per-document counts
        List<Dictionary<string, object>> categoryList = new List<Dictionary<string, object>>();

        foreach (var categoryInfo in categoryInfoMap.Values.OrderBy(c => c.CategoryName))
        {
            // Add entry for regular elements (if any)
            if (categoryInfo.ElementsByDocument.Count > 0 && !categoryInfo.IsDirectShape)
            {
                var entry = new Dictionary<string, object>
                {
                    { "Name", categoryInfo.CategoryName },
                    { "CategoryId", categoryInfo.CategoryId },
                    { "CategoryInfo", categoryInfo },
                    { "IsDirectShape", false },
                    { "IsView", categoryInfo.IsView },
                    { "IsViewTemplate", categoryInfo.IsViewTemplate },
                    { "IsSheet", categoryInfo.IsSheet },
                    { "IsViewFilter", categoryInfo.IsViewFilter }
                };

                // Add count column for each document
                foreach (string docTitle in documentTitles)
                {
                    int count = 0;
                    if (categoryInfo.ElementsByDocument.ContainsKey(docTitle))
                    {
                        count = categoryInfo.ElementsByDocument[docTitle].Count;
                    }
                    entry[docTitle] = count;
                }

                categoryList.Add(entry);
            }

            // Add entry for Direct Shapes (if any)
            if (categoryInfo.DirectShapesByDocument.Count > 0)
            {
                var entry = new Dictionary<string, object>
                {
                    { "Name", $"Direct Shapes: {categoryInfo.CategoryName}" },
                    { "CategoryId", categoryInfo.CategoryId },
                    { "CategoryInfo", categoryInfo },
                    { "IsDirectShape", true }
                };

                // Add count column for each document
                foreach (string docTitle in documentTitles)
                {
                    int count = 0;
                    if (categoryInfo.DirectShapesByDocument.ContainsKey(docTitle))
                    {
                        count = categoryInfo.DirectShapesByDocument[docTitle].Count;
                    }
                    entry[docTitle] = count;
                }

                categoryList.Add(entry);
            }
        }

        if (categoryList.Count == 0)
        {
            TaskDialog.Show("No Categories", "No categories found in open documents.");
            return Result.Cancelled;
        }

        // Define properties to display: Name + each document title
        var propertyNames = new List<string> { "Name" };
        propertyNames.AddRange(documentTitles);

        // Show DataGrid
        List<Dictionary<string, object>> selectedCategories = CustomGUIs.DataGrid(categoryList, propertyNames, false);
        if (selectedCategories.Count == 0)
            return Result.Cancelled;

        // Gather selection items for selected categories
        List<SelectionItem> selectionItems = new List<SelectionItem>();

        foreach (var selectedCategory in selectedCategories)
        {
            var categoryInfo = (CategoryInfo)selectedCategory["CategoryInfo"];
            bool isDirectShape = (bool)selectedCategory["IsDirectShape"];

            if (isDirectShape)
            {
                // Add Direct Shapes
                foreach (var docEntry in categoryInfo.DirectShapesByDocument)
                {
                    string docTitle = docEntry.Key;
                    foreach (var elemEntry in docEntry.Value)
                    {
                        selectionItems.Add(new SelectionItem
                        {
                            DocumentTitle = docTitle,
                            DocumentPath = GetDocumentPath(app, docTitle),
                            UniqueId = elemEntry.Value,
#if REVIT2024 || REVIT2025 || REVIT2026
                            ElementIdValue = (int)elemEntry.Key.Value,
#else
                            ElementIdValue = elemEntry.Key.IntegerValue,
#endif
                            SessionId = null // Will be auto-generated
                        });
                    }
                }
            }
            else
            {
                // Add regular elements
                foreach (var docEntry in categoryInfo.ElementsByDocument)
                {
                    string docTitle = docEntry.Key;
                    foreach (var elemEntry in docEntry.Value)
                    {
                        selectionItems.Add(new SelectionItem
                        {
                            DocumentTitle = docTitle,
                            DocumentPath = GetDocumentPath(app, docTitle),
                            UniqueId = elemEntry.Value,
#if REVIT2024 || REVIT2025 || REVIT2026
                            ElementIdValue = (int)elemEntry.Key.Value,
#else
                            ElementIdValue = elemEntry.Key.IntegerValue,
#endif
                            SessionId = null // Will be auto-generated
                        });
                    }
                }
            }
        }

        // Load existing selection and merge
        var existingSelection = SelectionStorage.LoadSelection();
        var existingUniqueIds = new HashSet<string>(existingSelection.Select(s => $"{s.DocumentTitle}|{s.UniqueId}"));

        // Add new items that don't already exist
        foreach (var item in selectionItems)
        {
            string key = $"{item.DocumentTitle}|{item.UniqueId}";
            if (!existingUniqueIds.Contains(key))
            {
                existingSelection.Add(item);
            }
        }

        // Save selection
        SelectionStorage.SaveSelection(existingSelection);

        return Result.Succeeded;
    }

    private string GetDocumentPath(Autodesk.Revit.ApplicationServices.Application app, string documentTitle)
    {
        foreach (Document doc in app.Documents)
        {
            if (doc.Title == documentTitle)
            {
                return doc.PathName ?? documentTitle;
            }
        }
        return documentTitle;
    }

    /// <summary>
    /// Tracks category information across documents.
    /// </summary>
    private class CategoryInfo
    {
        public ElementId CategoryId { get; set; }
        public string CategoryName { get; set; }
        public bool IsDirectShape { get; set; }
        public bool IsView { get; set; }
        public bool IsViewTemplate { get; set; }
        public bool IsSheet { get; set; }
        public bool IsViewFilter { get; set; }

        // Document Title -> (ElementId -> UniqueId)
        public Dictionary<string, Dictionary<ElementId, string>> ElementsByDocument { get; set; }
            = new Dictionary<string, Dictionary<ElementId, string>>();

        public Dictionary<string, Dictionary<ElementId, string>> DirectShapesByDocument { get; set; }
            = new Dictionary<string, Dictionary<ElementId, string>>();

        public int TotalCount => ElementsByDocument.Sum(d => d.Value.Count);
        public int DirectShapeCount => DirectShapesByDocument.Sum(d => d.Value.Count);

        public void AddElement(Document doc, ElementId elementId, string uniqueId, bool isDirectShape)
        {
            string docTitle = doc.Title;

            if (isDirectShape)
            {
                if (!DirectShapesByDocument.ContainsKey(docTitle))
                    DirectShapesByDocument[docTitle] = new Dictionary<ElementId, string>();
                DirectShapesByDocument[docTitle][elementId] = uniqueId;
            }
            else
            {
                if (!ElementsByDocument.ContainsKey(docTitle))
                    ElementsByDocument[docTitle] = new Dictionary<ElementId, string>();
                ElementsByDocument[docTitle][elementId] = uniqueId;
            }
        }
    }
}
