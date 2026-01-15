using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SynchronizeDocumentsInSession : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument activeUidoc = uiApp.ActiveUIDocument;
        Document activeDoc = activeUidoc?.Document;

        if (activeDoc == null)
        {
            TaskDialog.Show("Error", "No active document.");
            return Result.Failed;
        }

        // Collect all open workshared documents
        var documentsData = new List<Dictionary<string, object>>();

        foreach (Document doc in uiApp.Application.Documents)
        {
            // Skip linked documents (read-only references)
            if (doc.IsLinked)
                continue;

            // Skip family documents
            if (doc.IsFamilyDocument)
                continue;

            // Only include workshared documents
            if (!doc.IsWorkshared)
                continue;

            string projectName = doc.Title;
            bool isActiveDoc = (doc == activeDoc);

            var dict = new Dictionary<string, object>
            {
                ["Document"] = projectName,
                ["Active"] = isActiveDoc ? "Yes" : "No",
                ["__Document"] = doc
            };

            documentsData.Add(dict);
        }

        if (documentsData.Count == 0)
        {
            TaskDialog.Show("Info", "No workshared documents open.");
            return Result.Failed;
        }

        // Sort by document name
        documentsData = documentsData.OrderBy(d => d["Document"].ToString()).ToList();

        // Build property names
        var propertyNames = new List<string> { "Document", "Active" };

        // Show the grid with multi-selection enabled
        CustomGUIs.SetCurrentUIDocument(activeUidoc);
        var selectedDicts = CustomGUIs.DataGrid(documentsData, propertyNames, true); // true = multi-select

        if (selectedDicts == null || selectedDicts.Count == 0)
            return Result.Cancelled;

        // Synchronize selected documents sequentially
        foreach (var selectedDict in selectedDicts)
        {
            Document targetDoc = selectedDict["__Document"] as Document;
            if (targetDoc == null)
                continue;

            try
            {
                // Perform synchronization
                SynchronizeDocument(targetDoc);
            }
            catch (Exception ex)
            {
                // Only show dialog on error
                TaskDialog.Show("Synchronization Error",
                    $"Failed to synchronize '{targetDoc.Title}':\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        return Result.Succeeded;
    }

    private void SynchronizeDocument(Document doc)
    {
        if (!doc.IsWorkshared)
        {
            throw new InvalidOperationException("Document is not workshared.");
        }

        // Create synchronization options
        var transactOptions = new TransactWithCentralOptions();
        var syncOptions = new SynchronizeWithCentralOptions();

        // Configure sync options to match "Synchronize Now" defaults
        syncOptions.Comment = "Synchronized via Revit Ballet";
        syncOptions.Compact = false; // Do NOT compact - that's a separate, slower operation
        syncOptions.SaveLocalBefore = false; // No need to save before sync
        syncOptions.SaveLocalAfter = true; // Save local file after sync (standard behavior)

        // Perform synchronization
        doc.SynchronizeWithCentral(transactOptions, syncOptions);
    }
}
