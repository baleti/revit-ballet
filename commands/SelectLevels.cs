using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace MyRevitAddin
{
    [Transaction(TransactionMode.ReadOnly)]
    public class SelectLevels : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the active document.
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            // Collect all Level elements in the document.
            List<Level> levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();

            // Prepare a list to hold level data (each level is represented as a dictionary).
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            // Start with some common property names.
            List<string> propertyNames = new List<string> { "Id", "Name", "Elevation" };

            // Loop through each level.
            foreach (Level lvl in levels)
            {
                // Create a dictionary for this level.
                Dictionary<string, object> entry = new Dictionary<string, object>();

                // Add basic properties.
                entry["Id"] = lvl.Id.ToString();
                entry["Name"] = lvl.Name;
                // Convert elevation from feet to millimeters, and round to the nearest integer.
                entry["Elevation"] = (int)Math.Round(lvl.Elevation * 304.8, MidpointRounding.AwayFromZero);
                
                // Iterate over each parameter of the level and add it to the dictionary.
                // Also add any new parameter names to the propertyNames list.
                foreach (Parameter p in lvl.Parameters)
                {
                    string paramName = p.Definition.Name;
                    // Avoid duplicates.
                    if (!entry.ContainsKey(paramName))
                    {
                        string value = p.AsValueString();
                        if (value == null)
                            value = p.AsString();
                        // If no value is available, try retrieving a double value.
                        if (value == null)
                        {
                            try 
                            {
                                value = p.AsDouble().ToString();
                            }
                            catch 
                            {
                                value = string.Empty;
                            }
                        }
                        entry[paramName] = value;

                        // If this parameter has not been added as a column, add it now.
                        if (!propertyNames.Contains(paramName))
                        {
                            propertyNames.Add(paramName);
                        }
                    }
                }
                
                entries.Add(entry);
            }

            // Display the data grid and capture the selected entries.
            List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false);

            // If the user didn't select anything, abort the command immediately.
            if (selectedEntries == null || selectedEntries.Count == 0)
            {
                return Result.Succeeded;
            }

            // Extract the element Ids from the selected entries.
            List<ElementId> selectedIds = new List<ElementId>();
            foreach (var entry in selectedEntries)
            {
                if (entry.ContainsKey("Id") && entry["Id"] != null)
                {
                    if (int.TryParse(entry["Id"].ToString(), out int idInt))
                    {
                        selectedIds.Add(new ElementId(idInt));
                    }
                }
            }

            // Update the active selection with the chosen level elements.
            uidoc.SetSelectionIds(selectedIds);

            return Result.Succeeded;
        }
    }
}
