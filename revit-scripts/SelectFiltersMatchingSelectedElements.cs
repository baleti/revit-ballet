using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace YourAddinNamespace
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectFiltersMatchingSelectedElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uiDoc.Document;

            /* ------------------------------------------------------------------ *
             * 1) Current selection (IDs + References, host & linked)             *
             * ------------------------------------------------------------------ */
            var selInfos = new List<SelInfo>();

            // host-model IDs
            foreach (ElementId id in uiDoc.GetSelectionIds())
            {
                Element e = doc.GetElement(id);
                if (e != null) selInfos.Add(new SelInfo(e, doc));
            }

            // references – may include linked
            foreach (Reference r in uiDoc.GetReferences())
            {
                if (r.LinkedElementId == ElementId.InvalidElementId)
                {
                    Element e = doc.GetElement(r);
                    if (e != null && selInfos.All(si => si.Element.Id != e.Id))
                        selInfos.Add(new SelInfo(e, doc));
                }
                else
                {
                    RevitLinkInstance link = doc.GetElement(r.ElementId) as RevitLinkInstance;
                    if (link == null) continue;
                    Document ldoc = link.GetLinkDocument();
                    if (ldoc == null) continue;
                    Element le = ldoc.GetElement(r.LinkedElementId);
                    if (le == null) continue;

                    selInfos.Add(new SelInfo(le, ldoc, link.Name));
                }
            }

            if (!selInfos.Any())
            {
                TaskDialog.Show("Select elements", "Please select one or more elements first.");
                return Result.Cancelled;
            }

            /* ------------------------------------------------------------------ *
             * 2) Check which filters are already applied in active view          *
             * ------------------------------------------------------------------ */
            View view         = doc.ActiveView;
            var  usedFilters  = new HashSet<ElementId>(view.GetFilters());
            if (view.ViewTemplateId != ElementId.InvalidElementId &&
                doc.GetElement(view.ViewTemplateId) is View vt)
            {
                usedFilters.UnionWith(vt.GetFilters());
            }

            /* ------------------------------------------------------------------ *
             * 3) Generate per-element column names (only if >1 element)          *
             * ------------------------------------------------------------------ */
            bool multiSel = selInfos.Count > 1;
            var elemColNames = multiSel
                ? selInfos.Select((s, i) => $"Element {i + 1}").ToList()
                : new List<string>();   // none when single selection

            /* ------------------------------------------------------------------ *
             * 4) Build DataGrid rows                                             *
             * ------------------------------------------------------------------ */
            var rows    = new List<Dictionary<string, object>>();
            var key2PF  = new Dictionary<string, ParameterFilterElement>();
            int keyIdx  = 0;

            foreach (ParameterFilterElement pf in
                     new FilteredElementCollector(doc)
                     .OfClass(typeof(ParameterFilterElement))
                     .Cast<ParameterFilterElement>())
            {
                var catIds   = new HashSet<ElementId>(pf.GetCategories());
                var epFilter = pf.GetElementFilter();

                var perElemStrings = new string[selInfos.Count];
                int matchCount     = 0;

                for (int i = 0; i < selInfos.Count; ++i)
                {
                    var info = selInfos[i];
                    if (!catIds.Contains(info.Element.Category?.Id)) continue;
                    if (epFilter != null && epFilter.PassesFilter(info.Element))
                    {
                        matchCount++;

                        // Text to show: type name; if more than one selected of same type, append IDs
                        perElemStrings[i] = info.DuplicateTypeCount > 1
                            ? $"{info.TypeName} ({info.Element.Id.IntegerValue})"
                            : info.TypeName;
                    }
                }

                if (matchCount == 0) continue;                       // filter irrelevant

                string key = $"{pf.Id.IntegerValue}_{keyIdx++}";
                key2PF[key] = pf;

                var row = new Dictionary<string, object>
                {
                    ["FilterName"]   = pf.Name,
                    ["MatchesCount"] = matchCount,
                    ["Used"]         = usedFilters.Contains(pf.Id) ? "Yes" : "",
                    ["UniqueKey"]    = key
                };

                // add per-element columns only when multi-selection
                if (multiSel)
                {
                    for (int i = 0; i < elemColNames.Count; ++i)
                        row[elemColNames[i]] = perElemStrings[i] ?? "";
                }

                rows.Add(row);
            }

            if (!rows.Any())
            {
                TaskDialog.Show("No matches",
                                "None of the project’s filters would select the currently-selected elements.");
                return Result.Cancelled;
            }

            /* ------------------------------------------------------------------ *
             * 5) Show DataGrid                                                   *
             * ------------------------------------------------------------------ */
            var columns = new List<string> { "FilterName", "MatchesCount", "Used" };
            if (multiSel) columns.AddRange(elemColNames);

            var chosen = CustomGUIs.DataGrid(rows, columns, spanAllScreens:false);
            if (chosen == null || chosen.Count == 0)
                return Result.Cancelled;

            /* ------------------------------------------------------------------ *
             * 6) Select chosen filters                                           *
             * ------------------------------------------------------------------ */
            var ids = new List<ElementId>();
            foreach (var r in chosen)
            {
                if (r.TryGetValue("UniqueKey", out var o) && o is string key &&
                    key2PF.TryGetValue(key, out var pf))
                {
                    ids.Add(pf.Id);
                }
            }

            if (ids.Any())
            {
                uiDoc.SetSelectionIds(ids);
                return Result.Succeeded;
            }

            return Result.Cancelled;
        }

        /* ====================================================================== *
         * Helper class to store selection info                                   *
         * ====================================================================== */
        private class SelInfo
        {
            public Element Element { get; }
            public string  TypeName { get; }
            public int     DuplicateTypeCount { get; set; }  // filled later

            public SelInfo(Element e, Document doc, string linkPrefix = null)
            {
                Element = e;
                string tn = "<no type>";
                try
                {
                    ElementId tid = e.GetTypeId();
                    if (tid != ElementId.InvalidElementId)
                    {
                        var tElem = doc.GetElement(tid);
                        if (tElem != null) tn = tElem.Name;
                    }
                }
                catch { /* ignore */ }

                TypeName = linkPrefix == null ? tn : $"{linkPrefix} :: {tn}";
            }
        }

        /* ===================== *
         * Static constructor    *
         * ===================== */
        static SelectFiltersMatchingSelectedElements()
        {
            // because DuplicateTypeCount needs all selection data
            // we hook into static ctor to run after SelInfo list assembled
        }

        /* ---------------------------------------------------------------------- *
         * After all SelInfo objects exist, compute duplicate counts              *
         * ---------------------------------------------------------------------- */
        private static void ComputeDuplicateTypeCounts(List<SelInfo> list)
        {
            var groups = list
                .GroupBy(s => s.TypeName)
                .ToDictionary(g => g.Key, g => g.Count());

            foreach (var info in list)
                info.DuplicateTypeCount = groups[info.TypeName];
        }
    }
}
