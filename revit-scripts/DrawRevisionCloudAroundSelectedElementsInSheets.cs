#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
#endregion

namespace MyCompany.RevitCommands
{
  [Transaction(TransactionMode.Manual)]
  public class DrawRevisionCloudAroundSelectedElementsInSheets : IExternalCommand
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
    // 2-D rectangle helper used for Boolean union
    // ------------------------------------------------------------------
    private class Rect2D
    {
      public double Xmin, Ymin, Xmax, Ymax;

      public Rect2D(double xmin, double ymin, double xmax, double ymax)
      {
        Xmin = xmin; Ymin = ymin; Xmax = xmax; Ymax = ymax;
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
      // 3) Ask which revision to assign clouds to
      //------------------------------------------------------------------
      Revision rev = PickRevision(doc);
      if (rev == null) return Result.Cancelled;   // user cancelled picker

      //------------------------------------------------------------------
      // 4) Transform cache
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
      // 5) Build rectangles per sheet – ELEMENT-FIRST
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
          // Only sheets that host the element’s owning view
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
      // 6) Create clouds – one per sheet with merged rectangles
      //------------------------------------------------------------------
      if (rectsBySheet.Count == 0)
      {
        TaskDialog.Show("Revision Clouds",
          "No revision clouds were created – none of the selected elements " +
          "are visible on the chosen sheets.");
        return Result.Cancelled;
      }

      using (Transaction t = new Transaction(doc, "Revision Clouds"))
      {
        t.Start();
        foreach (var kv in rectsBySheet)        // kv.Key = sheetId
        {
          var rects = kv.Value;
          if (rects.Count == 0) continue;

          List<Curve> curves = new List<Curve>();

          foreach (Rect2D r in rects)
          {
            XYZ bl = new XYZ(r.Xmin, r.Ymin, 0);
            XYZ tl = new XYZ(r.Xmin, r.Ymax, 0);
            XYZ tr = new XYZ(r.Xmax, r.Ymax, 0);
            XYZ br = new XYZ(r.Xmax, r.Ymin, 0);

            curves.Add(Line.CreateBound(bl, tl));
            curves.Add(Line.CreateBound(tl, tr));
            curves.Add(Line.CreateBound(tr, br));
            curves.Add(Line.CreateBound(br, bl));
          }

          RevisionCloud.Create(doc,
                               doc.GetElement(kv.Key) as ViewSheet,
                               rev.Id,
                               curves);
        }
        t.Commit();
      }

      //------------------------------------------------------------------
      // 7) Report skipped items, if any
      //------------------------------------------------------------------
      if (failedElems.Count > 0)
      {
        TaskDialog.Show("Revision Clouds",
          $"{failedElems.Count} of {selIds.Count} selected element(s) " +
          "are in Drafting/Legend views or not visible on any chosen sheet, " +
          "so no revision clouds were created for them.");
      }

      return Result.Succeeded;
    }

    //--------------------------------------------------------------------
    // Helper: pick or create a visible revision
    //--------------------------------------------------------------------
    private static Revision PickRevision(Document doc)
    {
      IList<Revision> revs = new FilteredElementCollector(doc)
                             .OfClass(typeof(Revision))
                             .Cast<Revision>()
                             .Where(r => r.Visibility != RevisionVisibility.Hidden)
                             .OrderBy(r => r.SequenceNumber)
                             .ToList();

      if (revs.Count == 0)
      {
        using (Transaction t = new Transaction(doc, "Create Revision"))
        {
          t.Start();
          Revision r = Revision.Create(doc);
          t.Commit();
          return r;
        }
      }

      var cols = new List<string>
      {
        "Revision Sequence", "Revision Date",
        "Revision Description", "Issued To", "Issued By"
      };

      var rows = revs.Select(rv => new Dictionary<string, object>
      {
        ["Revision Sequence"]    = rv.SequenceNumber,
        ["Revision Date"]        = rv.RevisionDate,
        ["Revision Description"] = rv.Description,
        ["Issued To"]            = rv.IssuedTo,
        ["Issued By"]            = rv.IssuedBy
      }).ToList();

      var pick = CustomGUIs.DataGrid(rows, cols, false,
                                     new List<int>{ rows.Count-1 });

      if (pick == null || pick.Count == 0) return null;

      int seq = Convert.ToInt32(pick[0]["Revision Sequence"]);
      return revs.First(rv => rv.SequenceNumber == seq);
    }
  }
}
