using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using RevitBallet.Core;

namespace RevitBallet.Commands;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Filtering Logic (Optimized)
    // ──────────────────────────────────────────────────────────────

    /// <summary>Applies all filters to the data and returns filtered result</summary>
    private static List<Dictionary<string, object>> ApplyFilters(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        string searchText,
        DataGridView grid)
    {
        // Quick return for empty filter
        if (string.IsNullOrWhiteSpace(searchText))
        {
            UpdateColumnVisibilityOptimized(grid, new HashSet<List<string>>(new ListStringComparer()), new List<ColumnOrderInfo>(), new List<bool>());
            return entries;
        }

        // Parse filter groups once (parser lives in RevitBallet.Core)
        List<FilterGroup> filterGroups = FilterQueryParser.Parse(searchText);

        // Update column visibility and ordering (optimized)
        HashSet<List<string>> allColVisibilityFilters = new HashSet<List<string>>(
            filterGroups.SelectMany(g => g.ColVisibilityFilters),
            new ListStringComparer());

        // Collect all column ordering from all groups
        List<ColumnOrderInfo> allColumnOrdering = filterGroups.SelectMany(g => g.ColumnOrdering).ToList();
        
        // Collect exact match flags for visibility
        List<bool> allColVisibilityExactMatch = filterGroups.SelectMany(g => g.ColVisibilityExactMatch).ToList();

        UpdateColumnVisibilityOptimized(grid, allColVisibilityFilters, allColumnOrdering, allColVisibilityExactMatch);

        // Row matching lives in RevitBallet.Core; Revit-specific lookups are injected
        List<Dictionary<string, object>> filtered = RowFilter.Filter(entries, filterGroups, BuildRowMatchContext(propertyNames));

        return filtered;
    }

    /// <summary>Optimized column visibility update - only updates when changed</summary>
    private static void UpdateColumnVisibilityOptimized(DataGridView grid, HashSet<List<string>> filters, List<ColumnOrderInfo> columnOrdering, List<bool> exactMatchFlags)
    {
        // Create a key for the current filter state (including ordering)
        string filterKey = string.Join("|", filters.Select(f => string.Join(",", f)));
        string orderingKey = string.Join("|", columnOrdering.Select(o => o.Position + ":" + string.Join(",", o.ColumnParts) + ":" + o.IsExactMatch));
        string exactMatchKey = string.Join("|", exactMatchFlags.Select(e => e ? "1" : "0"));
        string combinedKey = filterKey + "||" + orderingKey + "||" + exactMatchKey;

        // Skip if nothing changed
        if (combinedKey == _lastColumnVisibilityFilter + "||" + _lastColumnOrderingFilter)
            return;

        _lastColumnVisibilityFilter = filterKey + "||" + exactMatchKey;
        _lastColumnOrderingFilter = orderingKey;
        
        var newVisible = new HashSet<string>();

        // Update visibility
        if (filters.Count > 0)
        {
            var filterList = filters.ToList();
            for (int i = 0; i < filterList.Count; i++)
            {
                var parts = filterList[i];
                bool isExactMatch = i < exactMatchFlags.Count && exactMatchFlags[i];
                
                foreach (DataGridViewColumn col in grid.Columns)
                {
                    string colName = col.HeaderText.ToLowerInvariant();
                    bool show = false;
                    
                    if (isExactMatch)
                    {
                        // For exact match, join parts with space and compare
                        string exactPattern = string.Join(" ", parts);
                        show = colName == exactPattern;
                    }
                    else
                    {
                        // Original behavior: all parts must be contained
                        show = parts.All(p => colName.Contains(p));
                    }
                    
                    if (show) newVisible.Add(col.Name);
                }
            }
        }
        else
        {
            // No filters, show all columns
            foreach (DataGridViewColumn col in grid.Columns)
            {
                newVisible.Add(col.Name);
            }
        }

        // Apply column ordering if any
        if (columnOrdering.Count > 0)
        {
            grid.SuspendLayout(); // Prevent flicker

            // First, update visibility
            foreach (DataGridViewColumn col in grid.Columns)
            {
                col.Visible = newVisible.Contains(col.Name);
            }

            // Create a list to track new display indices
            var columnPositions = new List<Tuple<DataGridViewColumn, int>>();

            // Process columns with explicit ordering
            foreach (var orderInfo in columnOrdering.OrderBy(o => o.Position))
            {
                foreach (DataGridViewColumn col in grid.Columns)
                {
                    if (!col.Visible) continue;

                    string colName = col.HeaderText.ToLowerInvariant();
                    bool matches = false;
                    
                    if (orderInfo.IsExactMatch)
                    {
                        // For exact match, join parts with space and compare
                        string exactPattern = string.Join(" ", orderInfo.ColumnParts);
                        matches = colName == exactPattern;
                    }
                    else
                    {
                        // Original behavior: all parts must be contained
                        matches = orderInfo.ColumnParts.All(part => colName.Contains(part));
                    }
                    
                    if (matches)
                    {
                        // Check if this column is already in the list
                        if (!columnPositions.Any(cp => cp.Item1 == col))
                        {
                            columnPositions.Add(Tuple.Create(col, orderInfo.Position));
                        }
                    }
                }
            }

            // Add remaining visible columns that don't have explicit ordering
            int nextPosition = columnPositions.Count > 0 ? columnPositions.Max(cp => cp.Item2) + 1 : 1;
            foreach (DataGridViewColumn col in grid.Columns)
            {
                if (col.Visible && !columnPositions.Any(cp => cp.Item1 == col))
                {
                    columnPositions.Add(Tuple.Create(col, nextPosition++));
                }
            }

            // Apply the new display order
            int displayIndex = 0;
            foreach (var colPos in columnPositions.OrderBy(cp => cp.Item2))
            {
                colPos.Item1.DisplayIndex = displayIndex++;
            }

            grid.ResumeLayout();
            _lastVisibleColumns = newVisible;
        }
        else if (!_lastVisibleColumns.SetEquals(newVisible))
        {
            // Only update visibility if no ordering is specified and visibility changed
            grid.SuspendLayout(); // Prevent flicker
            foreach (DataGridViewColumn col in grid.Columns)
            {
                col.Visible = newVisible.Contains(col.Name);
            }
            grid.ResumeLayout();
            _lastVisibleColumns = newVisible;
        }
    }

    /// <summary>Builds the Core row-match context from the grid's cached state.</summary>
    private static RowMatchContext BuildRowMatchContext(List<string> propertyNames)
    {
        return new RowMatchContext
        {
            PropertyNames = propertyNames,
            SearchIndexByColumn = _searchIndexByColumn,
            SearchIndexAllColumns = _searchIndexAllColumns,
            SelectionSetLookup = GetSelectionSetElementIds,
            ForeignIdConverter = obj => obj is Autodesk.Revit.DB.ElementId eid ? eid.AsLong() : 0
        };
    }

}
