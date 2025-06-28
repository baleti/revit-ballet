using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class ListViewsByDetailItemSelected : IExternalCommand
{
    private HashSet<string> uniqueViewNames;
    private ListBox resultListBox;

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        // Create a new HashSet to store unique view names
        uniqueViewNames = new HashSet<string>();

        // Create a new ListBox
        resultListBox = new ListBox();

        // Get the active Revit document
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get the selected element
        ElementId selectedElementId = uidoc.GetSelectionIds().FirstOrDefault();

        if (selectedElementId != null)
        {
            Element selectedElement = doc.GetElement(selectedElementId);
            ElementId familyId = selectedElement.GetTypeId();

            // Get all instances of the family type
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<ElementId> instancesIds = collector
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Where(x => x.GetTypeId() == familyId)
                .Select(x => x.Id)
                .ToList();

            // Get the views where instances are placed
            List<Autodesk.Revit.DB.View> viewsWithInstances = new List<Autodesk.Revit.DB.View>();
            foreach (ElementId instanceId in instancesIds)
            {
                ICollection<ElementId> viewIds = ElementId.InvalidElementId != instanceId
                    ? ElementId.InvalidElementId != doc.GetElement(instanceId).OwnerViewId
                        ? new List<ElementId>() { doc.GetElement(instanceId).OwnerViewId }
                        : new List<ElementId>()
                    : new List<ElementId>();

                foreach (ElementId viewId in viewIds)
                {
                    var view = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                    if (view != null && !uniqueViewNames.Contains(view.Name))
                    {
                        viewsWithInstances.Add(view);
                        uniqueViewNames.Add(view.Name);
                    }
                }
            }

            // Populate the ListBox with unique view names
            foreach (Autodesk.Revit.DB.View view in viewsWithInstances)
            {
                resultListBox.Items.Add(view.Name);
            }

            // Show the form with the ListBox
            // Create a new form
            var form = new System.Windows.Forms.Form();
            form.StartPosition = FormStartPosition.CenterScreen;

            // Add a text box for searching
            var searchBox = new System.Windows.Forms.TextBox();
            searchBox.Dock = DockStyle.Top;
            searchBox.TextChanged += (sender, e) =>
            {
                FilterListBoxItems(resultListBox, searchBox.Text);
            };

            // Set the width of the list box dynamically based on the length of the longest view name
            int maxWidth = 0;
            Graphics g = resultListBox.CreateGraphics();
            foreach (string item in resultListBox.Items)
            {
                SizeF size = g.MeasureString(item, resultListBox.Font);
                maxWidth = Math.Max(maxWidth, (int)size.Width);
            }
            g.Dispose();
            // Set the width of the list box with some padding
            form.Width = maxWidth + 20;
            resultListBox.Dock = DockStyle.Fill;
            form.Controls.Add(resultListBox);
            form.Controls.Add(searchBox);
            searchBox.Select();

            // Set the height of the list box dynamically based on the number of entries
            int totalHeight = resultListBox.ItemHeight * resultListBox.Items.Count;
            form.Height = totalHeight + searchBox.Height + 50; // Adjust as needed

            // Event handler for key press (Enter key)
            resultListBox.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (resultListBox.SelectedItem != null)
                {
                    OpenSelectedView(uidoc, doc, resultListBox.SelectedItem.ToString(), form);
                }

            }
            else if (e.KeyCode == Keys.Escape)
            {
                form.Close();
            }
        };
        resultListBox.DoubleClick += (sender, e) =>
        {
            if (resultListBox.SelectedItem != null)
            {
                OpenSelectedView(uidoc, doc, resultListBox.SelectedItem.ToString(), form);
            }
        };
        searchBox.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                form.Close();
            }
            else if (e.KeyCode == Keys.Down)
            {
                resultListBox.Select();
            }
        };

        // Show the form
        Application.Run(form);
        }
        else
        {
            TaskDialog.Show("Error", "No element selected.");
        }

        return Result.Succeeded;
    }
    private void FilterListBoxItems(ListBox listBox, string searchText)
    {
        List<string> filteredItems = new List<string>();
        foreach (string item in listBox.Items)
        {
            if (item.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                filteredItems.Add(item);
            }
        }
        listBox.BeginUpdate();
        listBox.Items.Clear();
        listBox.Items.AddRange(filteredItems.ToArray());
        listBox.EndUpdate();
    }
    private void OpenSelectedView(UIDocument uidoc, Document doc, string selectedViewName, System.Windows.Forms.Form form)
    {
        // Get all views in the document
        FilteredElementCollector collector = new FilteredElementCollector(doc);
        ICollection<Element> views = collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements();

        // Loop through views to find the one with the specified name
        foreach (Element elem in views)
        {
            var view = elem as Autodesk.Revit.DB.View;
            if (view != null && view.Name == selectedViewName)
            {
                // Open the view
                try
                {
                    uidoc.ActiveView = view;
                    form.Close(); // Close the form after view activation
                    return;
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    TaskDialog.Show("Error", "Failed to open view. View may be invalid.");
                    return;
                }
            }
        }
        // If the loop completes without finding the view, display a message
        TaskDialog.Show("Error", "View '" + selectedViewName + "' not found.");
    }
}
