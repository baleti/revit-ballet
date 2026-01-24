using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitBallet.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DeleteProjectParameters : IExternalCommand
    {
        private class ProjectParameterData
        {
            public Definition Definition = null;
            public ElementBinding Binding = null;
            public string Name = null;
            public bool IsShared = false;
            public int UsageCount = 0;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                if (doc.IsFamilyDocument)
                {
                    TaskDialog.Show("Error", "This command cannot be used in family documents.");
                    return Result.Failed;
                }

                return ShowParametersAndDelete(doc, uidoc);
            }
            catch (OperationCanceledException)
            {
                message = "Operation cancelled by user.";
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to process project parameters: {ex.Message}");
                return Result.Failed;
            }
        }

        private Result ShowParametersAndDelete(Document doc, UIDocument uidoc)
        {
            // Get all project parameters
            List<ProjectParameterData> projectParametersData = GetProjectParameterData(doc);

            // Count parameter usage with progress dialog
            Dictionary<string, int> usageCounts = CountParameterUsageWithProgress(doc, projectParametersData);

            // Update usage counts in parameter data
            foreach (var paramData in projectParametersData)
            {
                if (usageCounts.ContainsKey(paramData.Name))
                {
                    paramData.UsageCount = usageCounts[paramData.Name];
                }
            }

            // Build entries for DataGrid - ALL parameters with usage counts
            List<Dictionary<string, object>> paramEntries = new List<Dictionary<string, object>>();

            foreach (var paramData in projectParametersData)
            {
                var entry = new Dictionary<string, object>
                {
                    { "Parameter Name", paramData.Name },
                    { "Count", paramData.UsageCount },
                    { "Is Shared", paramData.IsShared ? "Yes" : "No" },
                    { "Binding", GetBindingDescription(paramData.Binding) },
                    { "ParameterName", paramData.Name } // Hidden key for deletion
                };

                paramEntries.Add(entry);
            }

            // Sort by Count descending (most used first)
            paramEntries = paramEntries
                .OrderByDescending(e => (int)e["Count"])
                .ThenBy(e => e["Parameter Name"].ToString())
                .ToList();

            // Display in DataGrid
            CustomGUIs.SetCurrentUIDocument(uidoc);
            var propertyNames = new List<string> { "Parameter Name", "Count", "Is Shared", "Binding" };
            var selectedEntries = CustomGUIs.DataGrid(paramEntries, propertyNames, false);

            if (selectedEntries.Count == 0)
            {
                return Result.Cancelled;
            }

            // Get parameter names to delete from selected entries
            List<string> paramNamesToDelete = selectedEntries
                .Where(entry => entry.ContainsKey("ParameterName"))
                .Select(entry => entry["ParameterName"].ToString())
                .ToList();

            // Delete selected parameters
            DeleteParameters(doc, paramNamesToDelete);

            return Result.Succeeded;
        }

        private Dictionary<string, int> CountParameterUsageWithProgress(Document doc, List<ProjectParameterData> projectParametersData)
        {
            Dictionary<string, int> usageCounts = new Dictionary<string, int>();

            // Initialize all parameters with 0 count
            foreach (var paramData in projectParametersData)
            {
                usageCounts[paramData.Name] = 0;
            }

            HashSet<string> allProjectParams = new HashSet<string>(projectParametersData.Select(p => p.Name));

            // Collect and count instances
            var instances = new FilteredElementCollector(doc).WhereElementIsNotElementType().ToList();
            int instanceCount = instances.Count;
            int elementsProcessed = 0;
            int updateInterval = 100;

            // Scan instances with async progress dialog
            using (var progress = new AsyncProgressDialog("Analyzing parameter usage", 1500))
            {
                progress.Start();
                progress.SetTotal(instanceCount);

                foreach (var elem in instances)
                {
                    try
                    {
                        foreach (Parameter p in elem.Parameters)
                        {
                            // Skip built-in parameters (negative IDs)
                            if (p.Id.IntegerValue < 0)
                                continue;

                            string paramName = p.Definition.Name;
                            if (allProjectParams.Contains(paramName))
                            {
                                if (HasValueFast(p))
                                {
                                    usageCounts[paramName]++;
                                }
                            }
                        }
                    }
                    catch (Exception) { }

                    elementsProcessed++;

                    // Update progress every N elements
                    if (elementsProcessed % updateInterval == 0)
                    {
                        progress.SetProgress(elementsProcessed);
                        progress.CheckIfShouldShow();

                        if (progress.IsCancelled)
                            throw new OperationCanceledException("Operation cancelled by user.");

                        // Yield to UI thread to let timer tick
                        System.Windows.Forms.Application.DoEvents();
                    }
                }

                // Final progress update
                progress.SetProgress(elementsProcessed);
            }

            // Scan types (no progress needed, much faster)
            var types = new FilteredElementCollector(doc).WhereElementIsElementType().ToList();

            foreach (var elem in types)
            {
                try
                {
                    foreach (Parameter p in elem.Parameters)
                    {
                        if (p.Id.IntegerValue < 0)
                            continue;

                        string paramName = p.Definition.Name;
                        if (allProjectParams.Contains(paramName))
                        {
                            if (HasValueFast(p))
                            {
                                usageCounts[paramName]++;
                            }
                        }
                    }
                }
                catch (Exception) { }
            }

            return usageCounts;
        }

        // Optimized version with fewer checks
        private static bool HasValueFast(Parameter p)
        {
            try
            {
                if (!p.HasValue)
                    return false;

                // For Integer and Double, if HasValue is true, we have a value
                // Only need to check String and ElementId for empty/invalid values
                StorageType storageType = p.StorageType;
                if (storageType == StorageType.Integer || storageType == StorageType.Double)
                    return true;

                if (storageType == StorageType.String)
                    return !string.IsNullOrEmpty(p.AsString());

                if (storageType == StorageType.ElementId)
                {
                    ElementId elemId = p.AsElementId();
                    return elemId != null && elemId.IntegerValue != -1;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private string GetBindingDescription(ElementBinding binding)
        {
            if (binding == null)
                return "Unknown";

            if (binding is InstanceBinding)
                return "Instance";
            else if (binding is TypeBinding)
                return "Type";
            else
                return binding.GetType().Name;
        }

        private List<ProjectParameterData> GetProjectParameterData(Document doc)
        {
            if (doc == null)
                throw new ArgumentNullException("doc");

            if (doc.IsFamilyDocument)
                throw new Exception("doc can not be a family document.");

            List<ProjectParameterData> result = new List<ProjectParameterData>();

            BindingMap map = doc.ParameterBindings;
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();

            while (it.MoveNext())
            {
                ProjectParameterData newProjectParameterData = new ProjectParameterData
                {
                    Definition = it.Key,
                    Name = it.Key.Name,
                    Binding = it.Current as ElementBinding
                };

                // Check if it's a shared parameter
                if (it.Key is ExternalDefinition externalDef)
                {
                    newProjectParameterData.IsShared = true;
                }
                else
                {
                    newProjectParameterData.IsShared = false;
                }

                result.Add(newProjectParameterData);
            }

            return result;
        }

        private void DeleteParameters(Document doc, List<string> paramNames)
        {
            IEnumerable<ParameterElement> allParams = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(ParameterElement))
                .Cast<ParameterElement>();

            IList<ElementId> paramsToDelete = new List<ElementId>();

            foreach (var param in allParams)
            {
                if (paramNames.Contains(param.GetDefinition().Name))
                {
                    paramsToDelete.Add(param.Id);
                }
            }

            using (Transaction t = new Transaction(doc, "Delete Selected Project Parameters"))
            {
                t.Start();
                foreach (var paramId in paramsToDelete)
                {
                    doc.Delete(paramId);
                }
                t.Commit();
            }
        }
    }
}
