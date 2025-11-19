using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Virtual Mode Configuration
    // ──────────────────────────────────────────────────────────────

    /// <summary>Configure DataGridView for virtual mode operation</summary>
    private static void ConfigureVirtualMode(
        DataGridView grid, 
        List<Dictionary<string, object>> originalData,
        List<Dictionary<string, object>> filteredData,
        List<string> propertyNames)
    {
        // Store references
        _cachedOriginalData = originalData;
        _cachedFilteredData = filteredData;
        _currentGrid = grid;

        // Enable virtual mode
        grid.VirtualMode = true;

        // Handle cell value requests
        grid.CellValueNeeded -= Grid_CellValueNeeded; // Remove any existing handler
        grid.CellValueNeeded += Grid_CellValueNeeded;

        // Set row count
        grid.RowCount = filteredData.Count;

        // Build search index for the data
        BuildSearchIndex(originalData, propertyNames);
    }

    /// <summary>Virtual mode cell value provider</summary>
    private static void Grid_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
    {
        var grid = sender as DataGridView;
        if (grid == null || _cachedFilteredData == null) return;

        if (e.RowIndex >= 0 && e.RowIndex < _cachedFilteredData.Count)
        {
            var row = _cachedFilteredData[e.RowIndex];
            string columnName = grid.Columns[e.ColumnIndex].Name;

            object value;
            if (row.TryGetValue(columnName, out value))
            {
                e.Value = value;
            }
            else
            {
                e.Value = "";
            }
        }
    }

    /// <summary>Update virtual grid after filtering</summary>
    private static void UpdateVirtualGrid(DataGridView grid, List<Dictionary<string, object>> filteredData)
    {
        _cachedFilteredData = filteredData;
        
        // Suspend layout to prevent flicker
        grid.SuspendLayout();
        
        // Update row count
        grid.RowCount = filteredData.Count;
        
        // Force refresh
        grid.Invalidate();
        grid.ResumeLayout();
    }
}
