#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
#endregion

namespace RevitAddin
{
  [Transaction(TransactionMode.Manual)]
  public class ChangeLevelOfSelectedElements : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      // Get the currently selected elements.
      ICollection<ElementId> selIds = uidoc.GetSelectionIds();
      if (selIds == null || selIds.Count == 0)
      {
        message = "Please select one or more elements.";
        return Result.Failed;
      }

      // Get all Levels in the document.
      List<Level> levels = new FilteredElementCollector(doc)
                             .OfClass(typeof(Level))
                             .Cast<Level>()
                             .OrderBy(l => l.Elevation)
                             .ToList();

      // Prepare level entries for the GUI grid.
      List<Dictionary<string, object>> levelEntries = levels.Select(l => new Dictionary<string, object>
      {
        { "Name", l.Name },
        { "Elevation", l.Elevation }
      }).ToList();

      // Display the levels in a grid (using your provided CustomGUIs.DataGrid method)
      List<Dictionary<string, object>> selLevelEntry =
         CustomGUIs.DataGrid(levelEntries, new List<string> { "Name", "Elevation" }, false);

      if (selLevelEntry == null || selLevelEntry.Count == 0)
      {
        message = "No level selected.";
        return Result.Cancelled;
      }

      string targetLevelName = selLevelEntry[0]["Name"].ToString();
      Level targetLevel = levels.FirstOrDefault(l => l.Name.Equals(targetLevelName, StringComparison.OrdinalIgnoreCase));
      if (targetLevel == null)
      {
        message = "The selected level was not found.";
        return Result.Failed;
      }

      using (Transaction trans = new Transaction(doc, "Move Selected Elements To Level"))
      {
        trans.Start();

        // Process each selected element.
        foreach (ElementId id in selIds.ToList())
        {
          Element sourceElem = doc.GetElement(id);
          if (sourceElem == null) continue;

          // --- Case 1. Handle Walls.
          if (sourceElem is Wall wall)
          {
            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve == null) continue;

            Level oldLevel = doc.GetElement(wall.LevelId) as Level;
            double baseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0.0;
            double globalElev = (oldLevel != null ? oldLevel.Elevation : 0) + baseOffset;
            double newOffset = globalElev - targetLevel.Elevation;

            // Get wall height from its parameter.
            double wallHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 0.0;

            // Determine structural status using its parameter.
            Parameter structuralParam = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT);
            bool isStructural = (structuralParam != null && structuralParam.AsInteger() == 1);

            // Create a new wall.
            Wall newWall = Wall.Create(
              doc,
              locCurve.Curve,
              targetLevel.Id,
              wall.WallType.Id,
              wallHeight,
              newOffset,
              wall.Flipped,
              isStructural);

            // Copy instance parameters from the old wall to the new wall.
            CopyInstanceParameters(sourceElem, newWall);

            // Delete the original wall.
            doc.Delete(wall.Id);
          }
          // --- Case 2. Handle FamilyInstances (e.g. doors, windows).
          else if (sourceElem is FamilyInstance fi)
          {
            LocationPoint locPt = fi.Location as LocationPoint;
            if (locPt == null) continue;

            Level oldRefLevel = doc.GetElement(fi.LevelId) as Level;
            double instOffset = fi.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)?.AsDouble() ?? 0.0;
            double globalElev = (oldRefLevel != null ? oldRefLevel.Elevation : 0) + instOffset;
            double newInstOffset = globalElev - targetLevel.Elevation;

            FamilyInstance newFi = null;
            try
            {
              if (fi.Host != null)
              {
                newFi = doc.Create.NewFamilyInstance(
                  locPt.Point,
                  fi.Symbol,
                  fi.Host,
                  targetLevel,
                  fi.StructuralType);
              }
              else
              {
                newFi = doc.Create.NewFamilyInstance(
                  locPt.Point,
                  fi.Symbol,
                  targetLevel,
                  fi.StructuralType);
              }
            }
            catch (Exception)
            {
              continue;
            }

            // Update the instance's elevation offset.
            Parameter newOffsetParam = newFi.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
            if (newOffsetParam != null && !newOffsetParam.IsReadOnly)
            {
              newOffsetParam.Set(newInstOffset);
            }

            // Copy instance parameters from the original family instance.
            CopyInstanceParameters(sourceElem, newFi);

            // Delete the original family instance.
            doc.Delete(fi.Id);
          }
          // --- Case 3. Generic fallback for other element types.
          else
          {
            ElementId[] newIds = ElementTransformUtils.CopyElement(doc, sourceElem.Id, XYZ.Zero).ToArray();
            if (newIds.Length == 0)
              continue;

            Element newElem = doc.GetElement(newIds[0]);
            Parameter levelParam = newElem.get_Parameter(BuiltInParameter.LEVEL_PARAM);
            if (levelParam != null && !levelParam.IsReadOnly)
            {
              levelParam.Set(targetLevel.Id);
            }

            Parameter offParam = newElem.LookupParameter("Elevation from Level");
            if (offParam != null && !offParam.IsReadOnly)
            {
              double oldGlobalElev = 0;
              Parameter origOffParam = sourceElem.LookupParameter("Elevation from Level");
              if (origOffParam != null)
              {
                ElementId oldLevelId = sourceElem.get_Parameter(BuiltInParameter.LEVEL_PARAM)?.AsElementId();
                Level oldLevel = oldLevelId != null ? doc.GetElement(oldLevelId) as Level : null;
                oldGlobalElev = (oldLevel != null ? oldLevel.Elevation : 0) + origOffParam.AsDouble();
              }
              double newOff = oldGlobalElev - targetLevel.Elevation;
              offParam.Set(newOff);
            }

            // Copy instance parameters from the source element.
            CopyInstanceParameters(sourceElem, newElem);

            doc.Delete(sourceElem.Id);
          }
        } // end foreach

        trans.Commit();
      }

      return Result.Succeeded;
    }

    /// <summary>
    /// Copies writable instance parameters from the source element to the target element.
    /// </summary>
    private void CopyInstanceParameters(Element source, Element target)
    {
      // Iterate over all parameters of the source element.
      foreach (Parameter srcParam in source.Parameters)
      {
        // Skip if the parameter is read-only.
        if (srcParam.IsReadOnly)
          continue;

        // Try to find a corresponding parameter in the target element.
        Parameter tgtParam = target.LookupParameter(srcParam.Definition.Name);
        if (tgtParam == null || tgtParam.IsReadOnly)
          continue;

        // Copy the value based on storage type.
        switch (srcParam.StorageType)
        {
          case StorageType.Double:
            try
            {
              double val = srcParam.AsDouble();
              tgtParam.Set(val);
            }
            catch { }
            break;
          case StorageType.Integer:
            try
            {
              int val = srcParam.AsInteger();
              tgtParam.Set(val);
            }
            catch { }
            break;
          case StorageType.String:
            try
            {
              string val = srcParam.AsString();
              tgtParam.Set(val);
            }
            catch { }
            break;
          case StorageType.ElementId:
            try
            {
              ElementId val = srcParam.AsElementId();
              tgtParam.Set(val);
            }
            catch { }
            break;
          default:
            break;
        }
      }
    }
  }
}
