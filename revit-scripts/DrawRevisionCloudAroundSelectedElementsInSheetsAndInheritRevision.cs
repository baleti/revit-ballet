#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace MyCompany.RevitCommands
{
  [Transaction(TransactionMode.Manual)]
  public class DrawRevisionCloudAroundSelectedElementsInSheetsAndInheritRevision : IExternalCommand
  {
    private const double OFFSET_MODEL = 0.20;      // ft – halo in MODEL space

    // ------------------------------------------------------------------
    // Small helpers
    // ------------------------------------------------------------------
    private static bool IsTransformableView(View v) =>
      v.ViewType != ViewType.DraftingView &&
      v.ViewType != ViewType.Legend;

    private static bool IsVisibleInView(Document doc, View v, ElementId id) =>
      new FilteredElementCollector(doc, v.Id).WhereElementIsNotElementType()
                                             .Any(e => e.Id == id);

    // ------------------------------------------------------------------
    // Helper to get element description
    // ------------------------------------------------------------------
    private static string GetElementDescription(Element el)
    {
      string family = "Unknown";
      string type = "Unknown";
      string id = el.Id.ToString();

      if (el is FamilyInstance fi)
      {
        family = fi.Symbol?.Family?.Name ?? "Unknown Family";
        type = fi.Symbol?.Name ?? "Unknown Type";
      }
      else if (el.Category != null)
      {
        family = el.Category.Name;
        type = el.Name ?? el.GetType().Name;
      }
      else
      {
        family = el.GetType().Name;
        type = el.Name ?? "";
      }

      return $"{family} - {type} - {id}";
    }

    // ------------------------------------------------------------------
    // 2-D rectangle helper used for Boolean union
    // ------------------------------------------------------------------
    private class Rect2D
    {
      public double Xmin, Ymin, Xmax, Ymax;
      public HashSet<ElementId> ContributingElements { get; set; }

      public Rect2D(double xmin, double ymin, double xmax, double ymax)
      {
        Xmin = xmin; Ymin = ymin; Xmax = xmax; Ymax = ymax;
        ContributingElements = new HashSet<ElementId>();
      }

      public bool IntersectsOrTouches(Rect2D other) =>
        !(other.Xmin > Xmax || other.Xmax < Xmin
       || other.Ymin > Ymax || other.Ymax < Ymin);

      public void UnionWith(Rect2D other)
      {
        Xmin = Math.Min(Xmin, other.Xmin);
        Ymin = Math.Min(Ymin, other.Ymin);
        Xmax = Math.Max(Xmax, other.Xmax);
        Ymax = Math.Max(Ymax, other.Ymax);
        
        // Merge contributing elements
        foreach (var id in other.ContributingElements)
          ContributingElements.Add(id);
      }
    }

    // merge-or-add, keeps the list free of overlapping rectangles
    private static void MergeOrAddRect(List<Rect2D> list, Rect2D r)
    {
      for (int i = 0; i < list.Count; )
      {
        if (list[i].IntersectsOrTouches(r))
        {
          r.UnionWith(list[i]);
          list.RemoveAt(i);       // absorbed; restart scan
          i = 0;
        }
        else
          ++i;
      }
      list.Add(r);
    }

    // ==================================================================
    // MAIN EXECUTE
    // ==================================================================
    public Result Execute(
      ExternalCommandData cd,
      ref string message,
      ElementSet elements)
    {
      UIDocument uiDoc = cd.Application.ActiveUIDocument;
      Document   doc   = uiDoc.Document;

      //------------------------------------------------------------------
      // 0) Ensure a selection
      //------------------------------------------------------------------
      ICollection<ElementId> selIds = uiDoc.GetSelectionIds();
      if (selIds.Count == 0)
      {
        TaskDialog.Show("Revision Clouds",
          "Nothing is selected. Select one or more elements and run again.");
        return Result.Cancelled;
      }

      //------------------------------------------------------------------
      // 1) Build a map: viewId → (sheet, viewport) pairs
      //------------------------------------------------------------------
      var viewToVports = new Dictionary<ElementId, List<(ViewSheet, Viewport)>>();

      foreach (ViewSheet sh in new FilteredElementCollector(doc)
                                 .OfClass(typeof(ViewSheet))
                                 .Cast<ViewSheet>())
      {
        foreach (ElementId vpid in sh.GetAllViewports())
          if (doc.GetElement(vpid) is Viewport vp)
          {
            if (!viewToVports.TryGetValue(vp.ViewId, out var list))
            {
              list = new List<(ViewSheet, Viewport)>();
              viewToVports[vp.ViewId] = list;
            }
            list.Add((sh, vp));
          }
      }

      //------------------------------------------------------------------
      // 2) Decide sheets: auto vs. DataGrid
      //------------------------------------------------------------------
      bool allViewDependent = selIds.All(id =>
      {
        Element e = doc.GetElement(id);
        return e != null && e.OwnerViewId != ElementId.InvalidElementId;
      });

      HashSet<ElementId> chosenSheetIds = new HashSet<ElementId>();
      var failedElems = new HashSet<ElementId>();   // elements that never get a cloud

      if (allViewDependent)
      {
        // ---------- auto-detect sheets ----------
        foreach (ElementId elId in selIds)
        {
          Element   el  = doc.GetElement(elId);
          ElementId vid = el.OwnerViewId;

          if (!viewToVports.TryGetValue(vid, out var vps))
          {
            failedElems.Add(elId);          // owner view is not placed on a sheet
            continue;
          }

          foreach (var pair in vps)
            chosenSheetIds.Add(pair.Item1.Id);
        }
      }
      else
      {
        // ---------- DataGrid prompt ----------
        IList<ViewSheet> allSheets = viewToVports
                                     .SelectMany(kv => kv.Value)
                                     .Select(p => p.Item1)
                                     .Distinct()
                                     .OrderBy(sh => sh.SheetNumber)
                                     .ToList();

        if (allSheets.Count == 0)
        {
          TaskDialog.Show("Revision Clouds",
            "This project has no sheets with viewports.");
          return Result.Cancelled;
        }

        var cols = new List<string>
        {
          "Sheet Number", "Sheet Title",
          "Revision Sequence", "Revision Date",
          "Revision Description", "Issued To", "Issued By"
        };

        var gridRows = new List<Dictionary<string, object>>();
        var numToId  = new Dictionary<string, ElementId>();

        foreach (ViewSheet vs in allSheets)
        {
          IList<ElementId> revIds = vs.GetAllRevisionIds();

          var seq = new List<string>();
          var dat = new List<string>();
          var des = new List<string>();
          var to  = new List<string>();
          var by  = new List<string>();

          foreach (ElementId rid in revIds)
            if (doc.GetElement(rid) is Revision rv)
            {
              seq.Add(rv.SequenceNumber.ToString());
              dat.Add(rv.RevisionDate ?? string.Empty);
              des.Add(rv.Description  ?? string.Empty);
              to .Add(rv.IssuedTo     ?? string.Empty);
              by .Add(rv.IssuedBy     ?? string.Empty);
            }

          gridRows.Add(new Dictionary<string, object>
          {
            ["Sheet Number"]         = vs.SheetNumber,
            ["Sheet Title"]          = vs.Name,
            ["Revision Sequence"]    = string.Join(", ", seq),
            ["Revision Date"]        = string.Join(", ", dat),
            ["Revision Description"] = string.Join(", ", des),
            ["Issued To"]            = string.Join(", ", to),
            ["Issued By"]            = string.Join(", ", by)
          });

          numToId[vs.SheetNumber] = vs.Id;
        }

        var picked = CustomGUIs.DataGrid(gridRows, cols, spanAllScreens:false);
        if (picked == null || picked.Count == 0)
          return Result.Cancelled;

        foreach (var row in picked)
        {
          string sn = row["Sheet Number"].ToString();
          if (numToId.TryGetValue(sn, out ElementId id))
            chosenSheetIds.Add(id);
        }
      }

      // If nothing left to do, notify & exit
      if (chosenSheetIds.Count == 0)
      {
        TaskDialog.Show("Revision Clouds",
          "The selected element(s) are in Drafting or Legend views " +
          "that cannot be placed on sheets, therefore no revision clouds were created.");
        return Result.Cancelled;
      }

      //------------------------------------------------------------------
      // 3) Transform cache
      //------------------------------------------------------------------
      var xformCache = new Dictionary<(ElementId, ElementId), Transform>();

      Transform GetXform(View v, Viewport vp, ElementId shId)
      {
        if (!IsTransformableView(v)) return null;

        var key = (shId, v.Id);
        if (xformCache.TryGetValue(key, out Transform xf))
          return xf;

        try
        {
          xf = vp.GetProjectionToSheetTransform()
                 .Multiply(v.GetModelToProjectionTransforms()[0]
                             .GetModelToProjectionTransform());
          xformCache[key] = xf;
          return xf;
        }
        catch
        {
          return null;  // drafting / legend returns null
        }
      }

      //------------------------------------------------------------------
      // 4) Build rectangles per sheet – ELEMENT-FIRST
      //------------------------------------------------------------------
      var rectsBySheet = new Dictionary<ElementId, List<Rect2D>>();

      XYZ[] delta =
      {
        new XYZ(0,0,0), new XYZ(0,0,1), new XYZ(0,1,0), new XYZ(0,1,1),
        new XYZ(1,0,0), new XYZ(1,0,1), new XYZ(1,1,0), new XYZ(1,1,1)
      };

      foreach (ElementId elId in selIds)
      {
        Element el = doc.GetElement(elId);
        if (el == null) { failedElems.Add(elId); continue; }

        bool viewDep   = el.OwnerViewId != ElementId.InvalidElementId;
        bool addedAny  = false;

        IEnumerable<ElementId> candidateSheets;

        if (viewDep)
        {
          // Only sheets that host the element's owning view
          if (!viewToVports.TryGetValue(el.OwnerViewId, out var vps))
          {
            failedElems.Add(elId);            // view not on any sheet
            continue;
          }
          candidateSheets = vps
                           .Select(p => p.Item1.Id)
                           .Where(id => chosenSheetIds.Contains(id));
        }
        else
        {
          candidateSheets = chosenSheetIds;
        }

        foreach (ElementId shId in candidateSheets)
        {
          ViewSheet sheet = doc.GetElement(shId) as ViewSheet;
          if (sheet == null) continue;

          var vports = sheet.GetAllViewports()
                            .Select(id => doc.GetElement(id) as Viewport)
                            .Where(vp => vp != null)
                            .ToList();

          foreach (Viewport vp in vports)
          {
            if (viewDep && vp.ViewId != el.OwnerViewId) continue;

            View v = doc.GetElement(vp.ViewId) as View;
            if (v == null) continue;

            if (!IsVisibleInView(doc, v, el.Id)) continue;

            Transform xform = GetXform(v, vp, shId);
            if (xform == null) continue;         // drafting / legend

            BoundingBoxXYZ bb = el.get_BoundingBox(v);
            if (bb == null) continue;            // hidden by crop, filters, etc.

            // expand + project
            XYZ min = bb.Min - new XYZ(OFFSET_MODEL, OFFSET_MODEL, OFFSET_MODEL);
            XYZ max = bb.Max + new XYZ(OFFSET_MODEL, OFFSET_MODEL, OFFSET_MODEL);

            double sxMin = double.PositiveInfinity, syMin = double.PositiveInfinity;
            double sxMax = double.NegativeInfinity, syMax = double.NegativeInfinity;

            foreach (XYZ d in delta)
            {
              XYZ mc = new XYZ(
                d.X == 0 ? min.X : max.X,
                d.Y == 0 ? min.Y : max.Y,
                d.Z == 0 ? min.Z : max.Z);

              XYZ sp = xform.OfPoint(bb.Transform.OfPoint(mc));

              sxMin = Math.Min(sxMin, sp.X);
              syMin = Math.Min(syMin, sp.Y);
              sxMax = Math.Max(sxMax, sp.X);
              syMax = Math.Max(syMax, sp.Y);
            }

            var r = new Rect2D(sxMin, syMin, sxMax, syMax);
            r.ContributingElements.Add(elId);  // Track which element contributed to this rectangle

            if (!rectsBySheet.TryGetValue(shId, out var list))
            {
              list = new List<Rect2D>();
              rectsBySheet[shId] = list;
            }

            MergeOrAddRect(list, r);  // **BOOLEAN UNION HERE**
            addedAny = true;
            break;                    // once is enough for this sheet
          }
        }

        if (!addedAny)
          failedElems.Add(elId);      // not visible on any chosen sheet
      }

      //------------------------------------------------------------------
      // 5) Create clouds – one per sheet with merged rectangles
      //    using the latest revision from each sheet
      //------------------------------------------------------------------
      if (rectsBySheet.Count == 0)
      {
        TaskDialog.Show("Revision Clouds",
          "No revision clouds were created – none of the selected elements " +
          "are visible on the chosen sheets.");
        return Result.Cancelled;
      }

      int skippedSheets = 0;
      List<string> skippedSheetNumbers = new List<string>();

      using (Transaction t = new Transaction(doc, "Revision Clouds with Latest Revision"))
      {
        t.Start();
        foreach (var kv in rectsBySheet)        // kv.Key = sheetId
        {
          var rects = kv.Value;
          if (rects.Count == 0) continue;

          ViewSheet sheet = doc.GetElement(kv.Key) as ViewSheet;
          if (sheet == null) continue;

          // Get the latest revision for this sheet
          ElementId latestRevisionId = GetLatestRevisionForSheet(doc, sheet);
          
          if (latestRevisionId == null || latestRevisionId == ElementId.InvalidElementId)
          {
            skippedSheets++;
            skippedSheetNumbers.Add(sheet.SheetNumber);
            continue;  // Skip this sheet if no revision found
          }

          foreach (Rect2D r in rects)
          {
            XYZ bl = new XYZ(r.Xmin, r.Ymin, 0);
            XYZ tl = new XYZ(r.Xmin, r.Ymax, 0);
            XYZ tr = new XYZ(r.Xmax, r.Ymax, 0);
            XYZ br = new XYZ(r.Xmax, r.Ymin, 0);

            List<Curve> curves = new List<Curve>
            {
              Line.CreateBound(bl, tl),
              Line.CreateBound(tl, tr),
              Line.CreateBound(tr, br),
              Line.CreateBound(br, bl)
            };

            RevisionCloud cloud = RevisionCloud.Create(doc,
                                 sheet,
                                 latestRevisionId,
                                 curves);

            // Build comment with element descriptions
            if (cloud != null && r.ContributingElements.Count > 0)
            {
              StringBuilder comment = new StringBuilder("Elements: ");
              List<string> descriptions = new List<string>();
              
              foreach (ElementId id in r.ContributingElements)
              {
                Element elem = doc.GetElement(id);
                if (elem != null)
                {
                  descriptions.Add(GetElementDescription(elem));
                }
              }
              
              comment.Append(string.Join("; ", descriptions));
              
              // Set the Comments parameter
              Parameter commentsParam = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
              if (commentsParam != null && !commentsParam.IsReadOnly)
              {
                commentsParam.Set(comment.ToString());
              }
            }
          }
        }
        t.Commit();
      }

      return Result.Succeeded;
    }

    //--------------------------------------------------------------------
    // Helper: Get the latest revision for a sheet
    //--------------------------------------------------------------------
    private ElementId GetLatestRevisionForSheet(Document doc, ViewSheet sheet)
    {
      // First, check revisions assigned directly to the sheet
      ICollection<ElementId> sheetRevisionIds = sheet.GetAdditionalRevisionIds();
      
      // Also get all revision clouds on the sheet and its views to find their revisions
      HashSet<ElementId> allRevisionIds = new HashSet<ElementId>(sheetRevisionIds);
      
      // Add the sheet itself to check for revision clouds
      var viewsToCheck = new List<ElementId> { sheet.Id };
      
      // Add all views placed on this sheet
      var placedViews = sheet.GetAllPlacedViews();
      viewsToCheck.AddRange(placedViews);
      
      // Collect revision clouds in all these views
      var revisionClouds = new FilteredElementCollector(doc)
          .WhereElementIsNotElementType()
          .OfClass(typeof(RevisionCloud))
          .Cast<RevisionCloud>()
          .Where(rc => viewsToCheck.Contains(rc.OwnerViewId))
          .ToList();
      
      // Add revision IDs from revision clouds
      foreach (var cloud in revisionClouds)
      {
        if (cloud.RevisionId != ElementId.InvalidElementId)
        {
          allRevisionIds.Add(cloud.RevisionId);
        }
      }
      
      if (!allRevisionIds.Any())
      {
        return ElementId.InvalidElementId;
      }
      
      // Get all revision elements and find the one with the highest sequence number
      ElementId latestRevisionId = ElementId.InvalidElementId;
      int highestSequence = -1;
      
      foreach (ElementId revId in allRevisionIds)
      {
        Revision revision = doc.GetElement(revId) as Revision;
        if (revision != null && revision.Visibility != RevisionVisibility.Hidden)
        {
          int sequence = revision.SequenceNumber;
          if (sequence > highestSequence)
          {
            highestSequence = sequence;
            latestRevisionId = revId;
          }
        }
      }
      
      return latestRevisionId;
    }
  }
}
