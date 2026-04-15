using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitAddin
{
    /// <summary>
    /// Provides standardized input resolution for commands following the canonical
    /// pattern described in architecture-vision.org.
    ///
    /// Resolution rules:
    /// - View/Viewport input + nothing selected → active view
    /// - View/Viewport input + selection → filter to matching type
    /// - Sheet + nothing selected → active view if it's a sheet
    /// - Other element type → DataGrid picker if nothing relevant selected
    /// - Always silently filter mixed selections to the relevant type
    /// </summary>
    public static class InputResolver
    {
        /// <summary>
        /// Resolves views from the current selection or active view.
        /// </summary>
        /// <typeparam name="T">The specific View type to resolve (e.g., ViewPlan, View3D)</typeparam>
        /// <param name="uiDoc">The active UIDocument</param>
        /// <returns>List of resolved views</returns>
        public static List<T> ResolveViews<T>(UIDocument uiDoc) where T : View
        {
            Document doc = uiDoc.Document;
            ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
            List<T> views = new List<T>();

            // If something is selected, filter to matching view type
            if (selectedIds.Count > 0)
            {
                foreach (ElementId id in selectedIds)
                {
                    Element element = doc.GetElement(id);

                    // Direct view selection
                    if (element is T view)
                    {
                        views.Add(view);
                    }
                    // Viewport selection - extract the view
                    else if (element is Viewport viewport)
                    {
                        View viewFromViewport = doc.GetElement(viewport.ViewId) as View;
                        if (viewFromViewport is T typedView)
                        {
                            views.Add(typedView);
                        }
                    }
                }
            }

            // If no views found in selection, fall back to active view
            if (views.Count == 0 && uiDoc.ActiveView is T activeView)
            {
                views.Add(activeView);
            }

            return views;
        }

        /// <summary>
        /// Resolves views from the current selection or active view (non-generic version).
        /// </summary>
        /// <param name="uiDoc">The active UIDocument</param>
        /// <returns>List of resolved views</returns>
        public static List<View> ResolveViews(UIDocument uiDoc)
        {
            Document doc = uiDoc.Document;
            ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
            List<View> views = new List<View>();

            // If something is selected, filter to views
            if (selectedIds.Count > 0)
            {
                foreach (ElementId id in selectedIds)
                {
                    Element element = doc.GetElement(id);

                    // Direct view selection
                    if (element is View view)
                    {
                        views.Add(view);
                    }
                    // Viewport selection - extract the view
                    else if (element is Viewport viewport)
                    {
                        View viewFromViewport = doc.GetElement(viewport.ViewId) as View;
                        if (viewFromViewport != null)
                        {
                            views.Add(viewFromViewport);
                        }
                    }
                }
            }

            // If no views found in selection, fall back to active view
            if (views.Count == 0 && uiDoc.ActiveView != null)
            {
                views.Add(uiDoc.ActiveView);
            }

            return views;
        }

        /// <summary>
        /// Resolves sheets from the current selection or active view (if it's a sheet).
        /// </summary>
        /// <param name="uiDoc">The active UIDocument</param>
        /// <returns>List of resolved sheets</returns>
        public static List<ViewSheet> ResolveSheets(UIDocument uiDoc)
        {
            Document doc = uiDoc.Document;
            ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
            List<ViewSheet> sheets = new List<ViewSheet>();

            // If something is selected, filter to sheets
            if (selectedIds.Count > 0)
            {
                foreach (ElementId id in selectedIds)
                {
                    Element element = doc.GetElement(id);

                    // Direct sheet selection
                    if (element is ViewSheet sheet)
                    {
                        if (!sheets.Contains(sheet))
                        {
                            sheets.Add(sheet);
                        }
                    }
                    // Viewport selection - extract the sheet it belongs to
                    else if (element is Viewport viewport && viewport.SheetId != ElementId.InvalidElementId)
                    {
                        ViewSheet sheetFromViewport = doc.GetElement(viewport.SheetId) as ViewSheet;
                        if (sheetFromViewport != null && !sheets.Contains(sheetFromViewport))
                        {
                            sheets.Add(sheetFromViewport);
                        }
                    }
                }
            }

            // If no sheets found in selection, fall back to active sheet
            if (sheets.Count == 0 && uiDoc.ActiveView is ViewSheet activeSheet)
            {
                sheets.Add(activeSheet);
            }

            return sheets;
        }

        /// <summary>
        /// Resolves elements from the current selection, filtered by category.
        /// If nothing relevant is selected, returns null to indicate that the
        /// caller should show a DataGrid picker.
        /// </summary>
        /// <param name="uiDoc">The active UIDocument</param>
        /// <param name="categories">Categories to filter for</param>
        /// <returns>List of resolved elements, or null if DataGrid picker should be shown</returns>
        public static List<Element> ResolveElements(UIDocument uiDoc, params BuiltInCategory[] categories)
        {
            Document doc = uiDoc.Document;
            ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
            List<Element> elements = new List<Element>();

            // If nothing selected, return null to indicate DataGrid picker should be shown
            if (selectedIds.Count == 0)
            {
                return null;
            }

            // Filter selection to matching categories
            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element == null) continue;

                // If no categories specified, accept all elements
                if (categories.Length == 0)
                {
                    elements.Add(element);
                }
                else
                {
                    // Check if element matches any of the specified categories
                    foreach (BuiltInCategory cat in categories)
                    {
                        if (element.Category != null &&
                            element.Category.Id.AsLong() == (int)cat)
                        {
                            elements.Add(element);
                            break;
                        }
                    }
                }
            }

            // If no elements matched the filter, return null to indicate DataGrid picker
            if (elements.Count == 0)
            {
                return null;
            }

            return elements;
        }

        /// <summary>
        /// Resolves elements from selection or shows DataGrid picker as fallback.
        /// This is a convenience method that combines ResolveElements with DataGrid picker.
        /// </summary>
        /// <param name="uiDoc">The active UIDocument</param>
        /// <param name="pickerData">DataGrid data to show if no selection</param>
        /// <param name="pickerColumns">DataGrid columns to show</param>
        /// <param name="categories">Categories to filter for</param>
        /// <returns>List of resolved elements (from selection or DataGrid)</returns>
        public static List<Element> ResolveElementsOrPick(
            UIDocument uiDoc,
            List<Dictionary<string, object>> pickerData,
            List<string> pickerColumns,
            params BuiltInCategory[] categories)
        {
            Document doc = uiDoc.Document;

            // Try to resolve from selection first
            List<Element> elements = ResolveElements(uiDoc, categories);

            // If selection didn't yield anything, show DataGrid picker
            if (elements == null || elements.Count == 0)
            {
                var selected = CustomGUIs.DataGrid(pickerData, pickerColumns, false);
                if (selected == null || selected.Count == 0)
                {
                    return new List<Element>();
                }

                // Extract elements from DataGrid selection
                elements = new List<Element>();
                foreach (var row in selected)
                {
                    if (row.TryGetValue("ElementIdObject", out object idObj) && idObj is ElementId elementId)
                    {
                        Element elem = doc.GetElement(elementId);
                        if (elem != null)
                        {
                            elements.Add(elem);
                        }
                    }
                }
            }

            return elements;
        }
    }
}
