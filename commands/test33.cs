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
  public class SwapFamilyTypeOfSelectedElements : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      // Get the currently selected elements.
      ICollection<ElementId> selIds = uidoc.GetSelectionIds();
      if (selIds == null || selIds.Count == 0)
      {
        message = "Please select one or more family instance elements.";
        return Result.Failed;
      }

      // Retrieve all FamilySymbols (family types) in the document.
      List<FamilySymbol> familySymbols = new FilteredElementCollector(doc)
                                          .OfClass(typeof(FamilySymbol))
                                          .Cast<FamilySymbol>()
                                          .OrderBy(fs => fs.Family.Name)
                                          .ThenBy(fs => fs.Name)
                                          .ToList();

      // Prepare entries for the DataGrid.
      List<Dictionary<string, object>> familyTypeEntries = familySymbols.Select(fs => new Dictionary<string, object>
      {
        { "Id", fs.Id.Value }, // include the Id so we can look it up later.
        { "Family", fs.Family.Name },
        { "Type", fs.Name }
      }).ToList();

      // Include "Id" in the displayed property names so the returned dictionary contains the key.
      List<Dictionary<string, object>> selFamilyEntry =
         CustomGUIs.DataGrid(familyTypeEntries, new List<string> { "Id", "Family", "Type" }, false);

      if (selFamilyEntry == null || selFamilyEntry.Count == 0)
      {
        message = "No family type selected.";
        return Result.Cancelled;
      }

      // Retrieve the selected FamilySymbol using the "Id" key.
      int selectedIdValue = Convert.ToInt32(selFamilyEntry[0]["Id"]);
      ElementId selectedFamilySymbolId = new ElementId((long)selectedIdValue);
      FamilySymbol newFamilySymbol = doc.GetElement(selectedFamilySymbolId) as FamilySymbol;
      if (newFamilySymbol == null)
      {
        message = "Selected family type not found.";
        return Result.Failed;
      }

      // Ensure the family symbol is active.
      if (!newFamilySymbol.IsActive)
      {
        using (Transaction t = new Transaction(doc, "Activate Family Symbol"))
        {
          t.Start();
          newFamilySymbol.Activate();
          t.Commit();
        }
      }

      using (Transaction trans = new Transaction(doc, "Swap Family Type Of Selected Elements"))
      {
        trans.Start();

        foreach (ElementId id in selIds.ToList())
        {
          Element sourceElem = doc.GetElement(id);
          if (sourceElem == null)
            continue;

          // Process only FamilyInstance elements.
          if (sourceElem is FamilyInstance fi)
          {
            LocationPoint locPt = fi.Location as LocationPoint;
            if (locPt == null)
              continue;

            // Verify that the selected family type belongs to the same category.
            if (fi.Symbol.Family.FamilyCategory.Id.Value != newFamilySymbol.Family.FamilyCategory.Id.Value)
            {
              // Skip elements whose category doesn't match the selected family type.
              continue;
            }

            FamilyInstance newFi = null;
            try
            {
              Level instLevel = doc.GetElement(fi.LevelId) as Level;
              if (fi.Host != null)
              {
                newFi = doc.Create.NewFamilyInstance(
                  locPt.Point,
                  newFamilySymbol,
                  fi.Host,
                  instLevel,
                  fi.StructuralType);
              }
              else
              {
                newFi = doc.Create.NewFamilyInstance(
                  locPt.Point,
                  newFamilySymbol,
                  instLevel,
                  fi.StructuralType);
              }
            }
            catch (Exception)
            {
              // If creation fails, skip this element.
              continue;
            }

            // Copy instance parameters from the original element.
            CopyInstanceParameters(sourceElem, newFi);

            // Delete the original element.
            doc.Delete(sourceElem.Id);
          }
          else
          {
            // Skip non-FamilyInstance elements.
            continue;
          }
        }

        trans.Commit();
      }

      return Result.Succeeded;
    }

    /// <summary>
    /// Copies writable instance parameters from the source element to the target element.
    /// </summary>
    private void CopyInstanceParameters(Element source, Element target)
    {
      foreach (Parameter srcParam in source.Parameters)
      {
        // Skip read-only parameters.
        if (srcParam.IsReadOnly)
          continue;

        Parameter tgtParam = target.LookupParameter(srcParam.Definition.Name);
        if (tgtParam == null || tgtParam.IsReadOnly)
          continue;

        try
        {
          switch (srcParam.StorageType)
          {
            case StorageType.Double:
              tgtParam.Set(srcParam.AsDouble());
              break;
            case StorageType.Integer:
              tgtParam.Set(srcParam.AsInteger());
              break;
            case StorageType.String:
              tgtParam.Set(srcParam.AsString());
              break;
            case StorageType.ElementId:
              tgtParam.Set(srcParam.AsElementId());
              break;
          }
        }
        catch
        {
          // Ignore errors for parameters that cannot be copied.
        }
      }
    }
  }
}
