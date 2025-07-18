using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace RevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class CombineSelectedFilledRegions : IExternalCommand
    {
        private struct TagInfo
        {
            public ElementId TypeId;
            public XYZ HeadPosition;
            public bool HasLeader;
            public XYZ LeaderEnd;
            public XYZ LeaderElbow;
            public bool HasElbow;
            public TagOrientation Orientation;
            public LeaderEndCondition EndCondition;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            // Get selected filled regions
            ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
            List<FilledRegion> selectedRegions = new List<FilledRegion>();

            foreach (ElementId id in selectedIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is FilledRegion region)
                {
                    selectedRegions.Add(region);
                }
            }

            if (selectedRegions.Count < 2)
            {
                TaskDialog.Show("Error", "Please select at least two filled regions to combine.");
                return Result.Failed;
            }

            // Verify all regions are in the same view
            ElementId viewId = selectedRegions[0].OwnerViewId;
            if (!selectedRegions.All(r => r.OwnerViewId.IntegerValue == viewId.IntegerValue))
            {
                TaskDialog.Show("Error", "All selected filled regions must be in the same view.");
                return Result.Failed;
            }

            // Collect associated tags
            FilteredElementCollector tagCollector = new FilteredElementCollector(doc, viewId).OfClass(typeof(IndependentTag));
            List<IndependentTag> associatedTags = new List<IndependentTag>();
            foreach (IndependentTag tag in tagCollector.ToElements().Cast<IndependentTag>())
            {
                Element taggedElement = tag.GetTaggedLocalElements().FirstOrDefault();
                if (taggedElement != null && selectedIds.Contains(taggedElement.Id))
                {
                    associatedTags.Add(tag);
                }
            }

            List<TagInfo> tagInfos = new List<TagInfo>();
            foreach (var tag in associatedTags)
            {
                Reference taggedRef = tag.GetTaggedReferences().FirstOrDefault();
                TagInfo info = new TagInfo
                {
                    TypeId = tag.GetTypeId(),
                    HeadPosition = tag.TagHeadPosition,
                    HasLeader = tag.HasLeader,
                    Orientation = tag.TagOrientation
                };

                if (info.HasLeader && taggedRef != null)
                {
                    info.EndCondition = tag.LeaderEndCondition;
                    info.HasElbow = tag.HasLeaderElbow(taggedRef);
                    if (info.EndCondition == LeaderEndCondition.Free)
                    {
                        info.LeaderEnd = tag.GetLeaderEnd(taggedRef);
                    }
                    if (info.HasElbow)
                    {
                        info.LeaderElbow = tag.GetLeaderElbow(taggedRef);
                    }
                }

                tagInfos.Add(info);
            }

            // Determine the type for the combined region
            var uniqueTypeIds = selectedRegions.Select(r => r.GetTypeId()).Distinct().ToList();
            ElementId combinedTypeId;

            if (uniqueTypeIds.Count == 1)
            {
                combinedTypeId = uniqueTypeIds[0];
            }
            else
            {
                // Prepare data for DataGrid
                List<Dictionary<string, object>> typeData = new List<Dictionary<string, object>>();
                Dictionary<string, ElementId> nameToTypeIdMap = new Dictionary<string, ElementId>();

                foreach (ElementId typeId in uniqueTypeIds)
                {
                    string name = doc.GetElement(typeId).Name;
                    nameToTypeIdMap[name] = typeId;
                    typeData.Add(new Dictionary<string, object>
                    {
                        { "Type Name", name }
                    });
                }

                // Sort by name
                typeData = typeData.OrderBy(d => d["Type Name"].ToString()).ToList();

                // Define columns
                List<string> columns = new List<string> { "Type Name" };

                // Show selection dialog
                List<Dictionary<string, object>> selectedTypes = CustomGUIs.DataGrid(
                    typeData,
                    columns,
                    false,  // Don't span all screens
                    null    // No initial selection
                );

                // Check if user selected a type
                string selectedTypeName = null;
                if (selectedTypes != null && selectedTypes.Count > 0)
                {
                    selectedTypeName = selectedTypes[0]["Type Name"].ToString();
                    combinedTypeId = nameToTypeIdMap[selectedTypeName];
                }
                else
                {
                    message = "No type selected.";
                    return Result.Cancelled;
                }
            }

            using (Transaction trans = new Transaction(doc, "Combine Filled Regions"))
            {
                trans.Start();

                // Get the boundary style from the first region
                FilledRegion firstRegion = selectedRegions[0];
                GraphicsStyle boundaryStyle = GetBoundaryLineStyle(doc, firstRegion);

                // Collect all boundary loops from all selected regions
                List<CurveLoop> allLoops = new List<CurveLoop>();
                foreach (FilledRegion region in selectedRegions)
                {
                    IList<CurveLoop> loops = region.GetBoundaries();
                    allLoops.AddRange(loops);
                }

                // Create a new filled region with all loops combined
                FilledRegion combinedRegion = FilledRegion.Create(
                    doc,
                    combinedTypeId,
                    viewId,
                    allLoops
                );

                // Apply the boundary style from the first region to the new combined region
                if (boundaryStyle != null && combinedRegion != null)
                {
                    SetBoundaryLineStyle(doc, combinedRegion, boundaryStyle);
                }

                // Recreate tags on the new region
                foreach (var info in tagInfos)
                {
                    // Activate the symbol if necessary
                    FamilySymbol symbol = doc.GetElement(info.TypeId) as FamilySymbol;
                    if (symbol != null && !symbol.IsActive)
                    {
                        symbol.Activate();
                    }

                    Reference reference = new Reference(combinedRegion);
                    XYZ addLocation = info.HasLeader ? (info.LeaderEnd ?? info.HeadPosition) : info.HeadPosition;
                    IndependentTag newTag = IndependentTag.Create(
                        doc,
                        info.TypeId,
                        viewId,
                        reference,
                        info.HasLeader,
                        info.Orientation,
                        addLocation
                    );

                    if (info.HasLeader)
                    {
                        newTag.LeaderEndCondition = info.EndCondition;
                        if (info.EndCondition == LeaderEndCondition.Free && info.LeaderEnd != null)
                        {
                            newTag.SetLeaderEnd(reference, info.LeaderEnd);
                        }
                        if (info.HasElbow && info.LeaderElbow != null)
                        {
                            newTag.SetLeaderElbow(reference, info.LeaderElbow);
                        }
                        newTag.TagHeadPosition = info.HeadPosition;
                    }
                }

                // Delete original filled regions and associated tags
                List<ElementId> toDelete = selectedRegions.Select(r => r.Id).ToList();
                toDelete.AddRange(associatedTags.Select(t => t.Id));
                doc.Delete(toDelete);

                // Select the newly created combined region
                sel.SetElementIds(new List<ElementId> { combinedRegion.Id });

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Retrieves the line style (GraphicsStyle) of the first boundary line
        /// used by the given FilledRegion, by examining its dependent elements.
        /// </summary>
        private GraphicsStyle GetBoundaryLineStyle(Document doc, FilledRegion region)
        {
            var dependentIds = region.GetDependentElements(null); // Gets any dependent elements, including boundary lines
            foreach (ElementId depId in dependentIds)
            {
                Element e = doc.GetElement(depId);
                if (e is CurveElement curveElem)
                {
                    // The 'LineStyle' property is actually a GraphicsStyle on CurveElement
                    return curveElem.LineStyle as GraphicsStyle;
                }
            }
            return null;
        }

        /// <summary>
        /// Sets the boundary line style (GraphicsStyle) for all boundary lines
        /// of the newly created FilledRegion (again using its dependent elements).
        /// </summary>
        private void SetBoundaryLineStyle(Document doc, FilledRegion newRegion, GraphicsStyle style)
        {
            var dependentIds = newRegion.GetDependentElements(null);
            foreach (ElementId depId in dependentIds)
            {
                Element e = doc.GetElement(depId);
                if (e is CurveElement curveElem)
                {
                    curveElem.LineStyle = style;
                }
            }
        }
    }
}
