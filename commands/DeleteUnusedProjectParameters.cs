using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitBallet.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DeleteUnusedProjectParameters : IExternalCommand
    {
        private class ProjectParameterData
        {
            public Definition Definition = null;
            public ElementBinding Binding = null;
            public string Name = null;
            public bool IsSharedStatusKnown = false;
            public bool IsShared = false;
            public string GUID = null;
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

                CheckAndRemoveUnusedParameters(doc);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to clean up project parameters: {ex.Message}");
                return Result.Failed;
            }
        }

        private void CheckAndRemoveUnusedParameters(Document doc)
        {
            List<ProjectParameterData> projectParametersData = GetProjectParameterData(doc);

            // Build a set of ALL project parameter names
            HashSet<string> allProjectParams = new HashSet<string>(
                projectParametersData.Select(p => p.Name)
            );

            // Build a set of USED parameter names by scanning elements once
            HashSet<string> usedParams = new HashSet<string>();

            // Check instances
            var instances = new FilteredElementCollector(doc).WhereElementIsNotElementType();
            foreach (var elem in instances)
            {
                try
                {
                    foreach (Parameter p in elem.Parameters)
                    {
                        // Only track project parameters (not built-in)
                        if (p.IsShared || !p.IsReadOnly)
                        {
                            if (HasValue(p) && allProjectParams.Contains(p.Definition.Name))
                            {
                                usedParams.Add(p.Definition.Name);
                            }
                        }
                    }
                }
                catch (Exception) { }

                // Early exit if we've found all parameters
                if (usedParams.Count == allProjectParams.Count)
                    break;
            }

            // Check types
            var types = new FilteredElementCollector(doc).WhereElementIsElementType();
            foreach (var elem in types)
            {
                try
                {
                    foreach (Parameter p in elem.Parameters)
                    {
                        // Only track project parameters (not built-in)
                        if (p.IsShared || !p.IsReadOnly)
                        {
                            if (HasValue(p) && allProjectParams.Contains(p.Definition.Name))
                            {
                                usedParams.Add(p.Definition.Name);
                            }
                        }
                    }
                }
                catch (Exception) { }

                // Early exit if we've found all parameters
                if (usedParams.Count == allProjectParams.Count)
                    break;
            }

            // Find unused parameters
            List<string> unusedParams = allProjectParams.Except(usedParams).ToList();

            DeleteUnusedParameters(doc, unusedParams);
        }

        private bool HasValue(Parameter p)
        {
            if (!p.HasValue)
                return false;

            // Check based on storage type
            switch (p.StorageType)
            {
                case StorageType.String:
                    return !string.IsNullOrEmpty(p.AsString());
                case StorageType.Integer:
                    return true; // If HasValue is true, integer has a value
                case StorageType.Double:
                    return true; // If HasValue is true, double has a value
                case StorageType.ElementId:
                    ElementId elemId = p.AsElementId();
                    return elemId != null && elemId != ElementId.InvalidElementId;
                default:
                    return false;
            }
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

                result.Add(newProjectParameterData);
            }

            return result;
        }

        private void DeleteUnusedParameters(Document doc, List<string> paramNames)
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

            using (Transaction t = new Transaction(doc, "Remove Unused Parameters"))
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
