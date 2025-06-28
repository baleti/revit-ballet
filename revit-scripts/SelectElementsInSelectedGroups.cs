using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class SelectElementsInSelectedGroups : IExternalCommand
{
    // Categories and types to exclude
    private static readonly HashSet<string> ExcludedCategories = new HashSet<string>
    {
        "No Category",
        "Automatic Sketch Dimensions",
        "<Sketch>"
    };

    private static readonly HashSet<string> ExcludedTypeNames = new HashSet<string>
    {
        "<Sketch>",
        "SketchPlane",
        "ModelLine"
    };

    private bool ShouldIncludeElement(Element element)
    {
        if (element == null) return false;

        // Check category
        if (element.Category == null || ExcludedCategories.Contains(element.Category.Name))
            return false;

        // Check element type
        string typeName = element.GetType().Name;
        if (ExcludedTypeNames.Contains(typeName))
            return false;

        // Skip internal elements (but allow filled regions which might start with "<")
        if (element.Name.StartsWith("<") && !(element is FilledRegion))
            return false;

        return true;
    }

    private string GetElementTypeName(Element element)
    {
        if (element is FilledRegion) return "Filled Region";
        if (element is DetailLine) return "Detail Line";
        if (element is FamilyInstance fi && fi.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents)
            return "Detail Component";
        if (element is FamilyInstance fi2)
        {
            return fi2.Symbol?.FamilyName ?? element.GetType().Name;
        }
        return element.GetType().Name;
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            // Get selected groups
            IList<Element> selectedElements = uidoc.GetSelectionIds()
                .Select(id => doc.GetElement(id))
                .Where(e => e is Group)
                .ToList();

            if (!selectedElements.Any())
            {
                TaskDialog.Show("Error", "Please select at least one group (model or detail) first.");
                return Result.Failed;
            }

            // Collect all elements from selected groups
            List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
            
            foreach (Element groupElement in selectedElements)
            {
                Group group = groupElement as Group;
                if (group != null)
                {
                    var memberIds = group.GetMemberIds();
                    foreach (ElementId memberId in memberIds)
                    {
                        Element member = doc.GetElement(memberId);
                        
                        // Skip unwanted elements
                        if (!ShouldIncludeElement(member))
                            continue;

                        // Get element type name
                        string typeName = GetElementTypeName(member);

                        // Get family name if applicable
                        string familyName = "N/A";
                        if (member is FamilyInstance familyInstance)
                        {
                            var symbol = familyInstance.Symbol;
                            if (symbol != null && symbol.Family != null)
                            {
                                familyName = symbol.Family.Name;
                            }
                        }
                        else if (member is FilledRegion)
                        {
                            FilledRegion fr = member as FilledRegion;
                            var fillType = doc.GetElement(fr.GetTypeId()) as FilledRegionType;
                            familyName = fillType?.Name ?? "Unknown Fill Type";
                        }

                        entries.Add(new Dictionary<string, object>
                        {
                            { "Category", member.Category.Name },
                            { "Family/Type", familyName != "N/A" ? familyName : typeName },
                            { "Group", group.Name },
                            { "Name", string.IsNullOrEmpty(member.Name) ? "(unnamed)" : member.Name },
                            { "_elementId", member.Id.IntegerValue }  // Hidden field for selection
                        });
                    }
                }
            }

            if (!entries.Any())
            {
                TaskDialog.Show("Information", "No applicable elements found in the selected groups.");
                return Result.Succeeded;
            }

            // Define visible columns for the DataGrid
            List<string> propertyNames = new List<string>
            {
                "Category",
                "Family/Type",
                "Group",
                "Name",
                "_elementId"  // Hidden column for selection
            };

            // Show DataGrid and get selected entries
            List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(entries, propertyNames, false);

            if (selectedEntries != null && selectedEntries.Any())
            {
                // Create a list of ElementIds from selected entries
                List<ElementId> selectedIds = selectedEntries
                    .Select(entry => new ElementId((int)entry["_elementId"]))
                    .ToList();

                // Select the elements in Revit
                uidoc.SetSelectionIds(selectedIds);
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
