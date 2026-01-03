using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using RevitBallet.Commands;

[Transaction(TransactionMode.Manual)]
public class SynchronizeDocuments : IExternalCommand
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
        var results = new List<string>();
        int successCount = 0;
        int failureCount = 0;

        foreach (var selectedDict in selectedDicts)
        {
            Document targetDoc = selectedDict["__Document"] as Document;
            if (targetDoc == null)
                continue;

            string docName = targetDoc.Title;

            try
            {
                // Perform synchronization
                SynchronizeDocument(targetDoc);
                results.Add($"✓ {docName}");
                successCount++;
            }
            catch (Exception ex)
            {
                results.Add($"✗ {docName}: {ex.Message}");
                failureCount++;
            }
        }

        // Show results
        string resultMessage = $"Synchronization Complete\n\n" +
                             $"Success: {successCount}\n" +
                             $"Failed: {failureCount}\n\n" +
                             string.Join("\n", results);

        TaskDialog.Show("Synchronize Documents", resultMessage);

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

        // Configure sync options (similar to "Synchronize Now" defaults)
        // TODO: Verify these settings match "Synchronize Now" behavior
        syncOptions.Comment = "Synchronized via Revit Ballet";
        syncOptions.Compact = true; // Compact central model
        syncOptions.SaveLocalBefore = true; // Save local file before sync
        syncOptions.SaveLocalAfter = true; // Save local file after sync

        // Perform synchronization
        doc.SynchronizeWithCentral(transactOptions, syncOptions);
    }
}
