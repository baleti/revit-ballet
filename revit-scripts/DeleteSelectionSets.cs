using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class DeleteSelectionSets : IExternalCommand
    {
        public class SavedSelection
        {
            public string Name { get; set; }
            public string ElementCount { get; set; }
            public FilterElement Filter { get; set; }

            public override string ToString() => Name;
        }

        public Result Execute(ExternalCommandData commandData, ref string messageText, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Get all saved selection filters
                List<SavedSelection> savedSelections = new List<SavedSelection>();
                
                // Find all selection filters
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(SelectionFilterElement));
                
                var filters = collector.ToElements().Cast<SelectionFilterElement>();

                foreach (SelectionFilterElement filter in filters)
                {
                    // Since getting detailed info about selection filters is version-dependent,
                    // we'll just store the basic information
                    savedSelections.Add(new SavedSelection
                    {
                        Name = filter.Name,
                        ElementCount = "-",
                        Filter = filter
                    });
                }

                if (savedSelections.Count == 0)
                {
                    TaskDialog.Show("No Saved Selections", 
                        "There are no saved selections in the current document.");
                    return Result.Succeeded;
                }

                // Configure and show selection dialog
                List<string> propertyNames = new List<string> 
                { 
                    "Name"
                };

                var selectedFilters = CustomGUIs.DataGrid(
                    savedSelections,
                    propertyNames,
                    null,
                    "Select Saved Selections to Delete"
                );

                if (selectedFilters.Count > 0)
                {
                    using (Transaction trans = new Transaction(doc, "Delete Saved Selections"))
                    {
                        trans.Start();
                        foreach (var selection in selectedFilters)
                        {
                            doc.Delete(selection.Filter.Id);
                        }
                        trans.Commit();
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                messageText = ex.Message;
                return Result.Failed;
            }
        }
    }
}
