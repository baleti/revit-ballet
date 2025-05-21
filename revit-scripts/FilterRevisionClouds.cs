#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace MyCompany.RevitCommands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class FilterRevisionClouds : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uidoc.Document;

            // --- get the user's current selection -------------------------------------------
            IList<ElementId> selIds = uidoc.Selection.GetElementIds().ToList();
            if (selIds.Count == 0)
            {
                TaskDialog.Show("FilterRevisionClouds",
                    "Nothing is selected.\nPlease pick one or more revision clouds and run the command again.");
                return Result.Cancelled;
            }

            // --- keep only revision clouds --------------------------------------------------
            List<RevisionCloud> revisionClouds = selIds
                .Select(id => doc.GetElement(id))
                .OfType<RevisionCloud>()
                .ToList();

            if (revisionClouds.Count == 0)
            {
                TaskDialog.Show("FilterRevisionClouds",
                    "None of the selected elements are revision clouds.");
                return Result.Cancelled;
            }

            // --- visible column headers ------------------------------------------------------
            var columns = new List<string>
            {
                "Element ID",
                "Sheet Number",
                "Sheet Name",
                "Issued By",
                "Issued To",
                "Name",
                "Revision Number",
                "Revision Description",
                "Cloud Count"
            };

            // --- build grid data + quick lookup map ------------------------------------------
            var gridData = new List<Dictionary<string, object>>();
            var cloudIdMap = new Dictionary<string, ElementId>(); // <ElementId string, cloud Id>

            foreach (RevisionCloud cloud in revisionClouds)
            {
                // Get the sheet information if the cloud is placed on a sheet
                string sheetNumber = "Not on Sheet";
                string sheetName = "Not on Sheet";
                
                View ownerView = doc.GetElement(cloud.OwnerViewId) as View;
                if (ownerView is ViewSheet sheet)
                {
                    sheetNumber = sheet.SheetNumber;
                    sheetName = sheet.Name;
                }

                // Get revision information
                string issuedBy = string.Empty;
                string issuedTo = string.Empty;
                string revisionNumber = string.Empty;
                string revisionDescription = string.Empty;

                if (cloud.RevisionId != ElementId.InvalidElementId)
                {
                    Revision revision = doc.GetElement(cloud.RevisionId) as Revision;
                    if (revision != null)
                    {
                        issuedBy = revision.IssuedBy ?? string.Empty;
                        issuedTo = revision.IssuedTo ?? string.Empty;
                        revisionNumber = revision.SequenceNumber.ToString();
                        revisionDescription = revision.Description ?? string.Empty;
                    }
                }

                // Get cloud name (if it has one)
                string cloudName = string.Empty;
                Parameter nameParam = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                if (nameParam != null)
                {
                    cloudName = nameParam.AsString() ?? string.Empty;
                }

                // Count the number of separate cloud sketches in this revision cloud
                int cloudCount = 1; // Default to 1 cloud
                try
                {
                    // Try to get the geometry and count separate pieces
                    Options options = new Options();
                    GeometryElement geomElem = cloud.get_Geometry(options);
                    if (geomElem != null)
                    {
                        int curveCount = 0;
                        foreach (GeometryObject geomObj in geomElem)
                        {
                            if (geomObj is GeometryInstance geomInst)
                            {
                                GeometryElement instGeom = geomInst.GetInstanceGeometry();
                                if (instGeom != null)
                                {
                                    foreach (GeometryObject instObj in instGeom)
                                    {
                                        if (instObj is Curve)
                                            curveCount++;
                                    }
                                }
                            }
                            else if (geomObj is Curve)
                            {
                                curveCount++;
                            }
                        }
                        // Estimate cloud count based on curve count (rough approximation)
                        cloudCount = Math.Max(1, curveCount / 10); // Adjust divisor as needed
                    }
                }
                catch
                {
                    // If geometry access fails, keep default of 1
                    cloudCount = 1;
                }

                // Create a unique key for this cloud using its ElementId
                string elementIdStr = cloud.Id.IntegerValue.ToString();

                gridData.Add(new Dictionary<string, object>
                {
                    ["Element ID"] = elementIdStr,
                    ["Sheet Number"] = sheetNumber,
                    ["Sheet Name"] = sheetName,
                    ["Issued By"] = issuedBy,
                    ["Issued To"] = issuedTo,
                    ["Name"] = cloudName,
                    ["Revision Number"] = revisionNumber,
                    ["Revision Description"] = revisionDescription,
                    ["Cloud Count"] = cloudCount.ToString()
                });

                // Remember the id for later
                cloudIdMap[elementIdStr] = cloud.Id;
            }

            // --- show chooser grid -----------------------------------------------------------
            List<Dictionary<string, object>> chosen =
                CustomGUIs.DataGrid(gridData, columns, spanAllScreens: false);

            if (chosen == null || chosen.Count == 0)
                return Result.Cancelled;   // user cancelled or unchecked everything

            // --- translate chosen rows back into ElementIds ----------------------------------
            var newIds = new List<ElementId>();

            foreach (Dictionary<string, object> row in chosen)
            {
                // Use the Element ID column to directly map back to the correct element
                if (row.TryGetValue("Element ID", out object elementIdObj) && elementIdObj != null)
                {
                    string elementIdStr = elementIdObj.ToString();
                    if (cloudIdMap.TryGetValue(elementIdStr, out ElementId id))
                    {
                        newIds.Add(id);
                    }
                }
            }

            if (newIds.Count == 0)
                return Result.Cancelled;   // should not happen, but play safe

            // --- REPLACE the selection with only the revision clouds the user checked --------
            uidoc.Selection.SetElementIds(newIds);

            return Result.Succeeded;
        }
    }
}
