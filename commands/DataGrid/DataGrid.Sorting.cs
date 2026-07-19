using RevitBallet.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace RevitBallet.Commands;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Sorting Logic
    // ──────────────────────────────────────────────────────────────
    
    private static readonly NaturalComparer naturalComparer = new NaturalComparer();

    // NaturalComparer moved to RevitBallet.Core.

    /// <summary>Applies multi-column sorting to the data</summary>
    private static List<Dictionary<string, object>> ApplySorting(
        List<Dictionary<string, object>> data,
        List<SortCriteria> sortCriteria)
    {
        if (sortCriteria.Count == 0) return data;

        IOrderedEnumerable<Dictionary<string, object>> ordered = null;
        foreach (SortCriteria sc in sortCriteria)
        {
            Func<Dictionary<string, object>, object> key =
                x => x.ContainsKey(sc.ColumnName) ? x[sc.ColumnName] : null;

            if (ordered == null)
            {
                ordered = (sc.Direction == ListSortDirection.Ascending)
                    ? data.OrderBy(key, naturalComparer)
                    : data.OrderByDescending(key, naturalComparer);
            }
            else
            {
                ordered = (sc.Direction == ListSortDirection.Ascending)
                    ? ordered.ThenBy(key, naturalComparer)
                    : ordered.ThenByDescending(key, naturalComparer);
            }
        }
        return ordered.ToList();
    }
}
