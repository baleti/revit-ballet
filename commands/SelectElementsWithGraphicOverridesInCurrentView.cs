using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.ReadOnly)]
[Regeneration(RegenerationOption.Manual)]
public class SelectElementsWithGraphicOverridesInCurrentView : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        try
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            if (activeView == null)
            {
                // No dialogs requested—just fail quietly with a message for Revit's status.
                message = "No active view found.";
                return Result.Failed;
            }

            // Collect elements in the active view (instances only)
            var collector = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType();

            // Gather the ids of elements that have overrides
            var overriddenIds = new HashSet<ElementId>();

            foreach (Element elem in collector)
            {
                // Skip view-specific/invalid categories or negative ids
                if (elem.Category == null || elem.Id.AsLong() < 0)
                    continue;

                var ogs = activeView.GetElementOverrides(elem.Id);
                if (HasOverrides(ogs))
                    overriddenIds.Add(elem.Id);
            }

            // If nothing found, just succeed without changing selection.
            if (overriddenIds.Count == 0)
                return Result.Succeeded;

            // Merge with existing selection (additive)
            var currentSel = uidoc.Selection.GetElementIds();
            foreach (var id in currentSel)
                overriddenIds.Add(id);

            uidoc.Selection.SetElementIds(overriddenIds);

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }

    private bool HasOverrides(OverrideGraphicSettings settings)
    {
        return settings.ProjectionLineColor.IsValid ||
               settings.ProjectionLineWeight != -1 ||
               settings.ProjectionLinePatternId.AsLong() != -1 ||
               settings.CutLineColor.IsValid ||
               settings.CutLineWeight != -1 ||
               settings.CutLinePatternId.AsLong() != -1 ||
               settings.SurfaceForegroundPatternId.AsLong() != -1 ||
               settings.SurfaceForegroundPatternColor.IsValid ||
               settings.SurfaceBackgroundPatternId.AsLong() != -1 ||
               settings.SurfaceBackgroundPatternColor.IsValid ||
               settings.CutForegroundPatternId.AsLong() != -1 ||
               settings.CutForegroundPatternColor.IsValid ||
               settings.CutBackgroundPatternId.AsLong() != -1 ||
               settings.CutBackgroundPatternColor.IsValid ||
               settings.Transparency != 0 ||
               settings.Halftone ||
               settings.DetailLevel != ViewDetailLevel.Undefined;
    }
}
