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
  public class DrawRevisionCloudAroundSelectedElements : IExternalCommand
  {
    private const double OFFSET_MODEL = 0.20; // ft – halo in MODEL space

    public Result Execute(
      ExternalCommandData cd,
      ref string message,
      ElementSet elements)
    {
      UIDocument uiDoc = cd.Application.ActiveUIDocument;
      Document   doc   = uiDoc.Document;

      //------------------------------------------------------------------
      // 0) Collect the selection
      //------------------------------------------------------------------
      ICollection<ElementId> selIds = uiDoc.Selection.GetElementIds();
      if (selIds.Count == 0)
      {
        TaskDialog.Show("Revision Clouds",
          "Nothing is selected. Select one or more elements and run again.");
        return Result.Cancelled;
      }

      //------------------------------------------------------------------
      // 1) Sheet discovery: auto vs. prompt
      //------------------------------------------------------------------
      bool allViewDependent = selIds.All(id =>
      {
        Element e = doc.GetElement(id);
        return e != null && e.OwnerViewId != ElementId.InvalidElementId;
      });

      // Cache all view placements once
      var viewToViewports = new Dictionary<ElementId, List<(ViewSheet, Viewport)>>();

      foreach (ViewSheet sh in new FilteredElementCollector(doc)
                                 .OfClass(typeof(ViewSheet))
                                 .Cast<ViewSheet>())
      {
        foreach (ElementId vpid in sh.GetAllViewports())
        {
          if (doc.GetElement(vpid) is Viewport vp)
          {
            if (!viewToViewports.TryGetValue(vp.ViewId, out var list))
            {
              list = new List<(ViewSheet, Viewport)>();
              viewToViewports[vp.ViewId] = list;
            }
            list.Add((sh, vp));
          }
        }
      }

      HashSet<ElementId> chosenSheetIds = new HashSet<ElementId>();

      if (allViewDependent)
      {
        // --- AUTO mode -------------------------------------------------
        foreach (ElementId elId in selIds)
        {
          Element   el   = doc.GetElement(elId);
          ElementId vId  = el.OwnerViewId;
          if (viewToViewports.TryGetValue(vId, out var vps))
            foreach (var pair in vps)
              chosenSheetIds.Add(pair.Item1.Id);
        }

        if (chosenSheetIds.Count == 0)
        {
          message = "Selected annotation elements are not placed on any sheet.";
          return Result.Failed;
        }
      }
      else
      {
        // --- PROMPT mode (DataGrid) ------------------------------------
        IList<ViewSheet> sheetsForPrompt = viewToViewports
                                           .SelectMany(kv => kv.Value)
                                           .Select(p => p.Item1)
                                           .Distinct()
                                           .OrderBy(sh => sh.SheetNumber)
                                           .ToList();

        if (sheetsForPrompt.Count == 0)
        {
          message = "No sheets in the model contain viewports.";
          return Result.Failed;
        }

        var columns = new List<string>
        {
          "Sheet Number", "Sheet Title",
          "Revision Sequence", "Revision Date",
          "Revision Description", "Issued To", "Issued By"
        };

        var gridData        = new List<Dictionary<string, object>>();
        var sheetNumToIdMap = new Dictionary<string, ElementId>();

        foreach (ViewSheet vs in sheetsForPrompt)
        {
          IList<ElementId> revIds = vs.GetAllRevisionIds();

          var seq   = new List<string>();
          var date  = new List<string>();
          var desc  = new List<string>();
          var to    = new List<string>();
          var by    = new List<string>();

          foreach (ElementId rid in revIds)
          {
            if (doc.GetElement(rid) is Revision rv)
            {
              seq .Add(rv.SequenceNumber.ToString());
              date.Add(rv.RevisionDate ?? string.Empty);
              desc.Add(rv.Description  ?? string.Empty);
              to  .Add(rv.IssuedTo     ?? string.Empty);
              by  .Add(rv.IssuedBy     ?? string.Empty);
            }
          }

          gridData.Add(new Dictionary<string, object>
          {
            ["Sheet Number"]         = vs.SheetNumber,
            ["Sheet Title"]          = vs.Name,
            ["Revision Sequence"]    = string.Join(", ", seq),
            ["Revision Date"]        = string.Join(", ", date),
            ["Revision Description"] = string.Join(", ", desc),
            ["Issued To"]            = string.Join(", ", to),
            ["Issued By"]            = string.Join(", ", by)
          });

          sheetNumToIdMap[vs.SheetNumber] = vs.Id;
        }

        List<Dictionary<string, object>> chosen =
          CustomGUIs.DataGrid(gridData, columns, spanAllScreens: false);

        if (chosen == null || chosen.Count == 0)
          return Result.Cancelled;

        foreach (var row in chosen)
        {
          string sn = row["Sheet Number"].ToString();
          if (sheetNumToIdMap.TryGetValue(sn, out ElementId id))
            chosenSheetIds.Add(id);
        }

        if (chosenSheetIds.Count == 0)
          return Result.Cancelled;
      }

      //------------------------------------------------------------------
      // 2) Ask the user which revision to assign the clouds to
      //------------------------------------------------------------------
      IList<Revision> allRevs = new FilteredElementCollector(doc)
                                .OfClass(typeof(Revision))
                                .Cast<Revision>()
                                .Where(r => r.Visibility != RevisionVisibility.Hidden)
                                .OrderBy(r => r.SequenceNumber)
                                .ToList();

      Revision chosenRevision = null;

      if (allRevs.Count == 0)
      {
        // Create one silently
        using (Transaction t = new Transaction(doc, "Create Revision"))
        {
          t.Start();
          chosenRevision = Revision.Create(doc);
          t.Commit();
        }
      }
      else
      {
        var revCols = new List<string>
        {
          "Revision Sequence", "Revision Date",
          "Revision Description", "Issued To", "Issued By"
        };

        var revRows = allRevs.Select(rv => new Dictionary<string, object>
        {
          ["Revision Sequence"]    = rv.SequenceNumber,
          ["Revision Date"]        = rv.RevisionDate,
          ["Revision Description"] = rv.Description,
          ["Issued To"]            = rv.IssuedTo,
          ["Issued By"]            = rv.IssuedBy
        }).ToList();

        List<Dictionary<string, object>> revChosen =
          CustomGUIs.DataGrid(revRows, revCols, false,
                              new List<int> { revRows.Count - 1 });

        if (revChosen == null || revChosen.Count == 0)
          return Result.Cancelled;

        int seq = Convert.ToInt32(revChosen[0]["Revision Sequence"]);
        chosenRevision = allRevs.First(rv => rv.SequenceNumber == seq);
      }

      //------------------------------------------------------------------
      // 3) Transform cache with **fallback for 2-D views**
      //------------------------------------------------------------------
      var xformCache = new Dictionary<(ElementId, ElementId), Transform>();

      Transform GetModelToSheetXform(View v, Viewport vp, ElementId sheetId)
      {
        var key = (sheetId, v.Id);
        if (xformCache.TryGetValue(key, out Transform xf))
          return xf;

        Transform projToSheet = vp.GetProjectionToSheetTransform();

        try
        {
          IList<TransformWithBoundary> list = v.GetModelToProjectionTransforms();
          if (list.Count > 0)
            xf = projToSheet.Multiply(list[0].GetModelToProjectionTransform());
          else
            xf = projToSheet; // 2-D view returns no transforms
        }
        catch (Autodesk.Revit.Exceptions.InvalidOperationException)
        {
          xf = projToSheet;   // drafting / legend view
        }

        xformCache[key] = xf;
        return xf;
      }

      //------------------------------------------------------------------
      // 4) For every sheet build the rectangles
      //------------------------------------------------------------------
      var curvesBySheet = new Dictionary<ElementId, List<Curve>>();

      XYZ[] delta =
      {
        new XYZ(0,0,0), new XYZ(0,0,1), new XYZ(0,1,0), new XYZ(0,1,1),
        new XYZ(1,0,0), new XYZ(1,0,1), new XYZ(1,1,0), new XYZ(1,1,1)
      };

      foreach (ElementId sheetId in chosenSheetIds)
      {
        ViewSheet sheet = doc.GetElement(sheetId) as ViewSheet;
        if (sheet == null) continue;

        var vports = sheet.GetAllViewports()
                          .Select(id => doc.GetElement(id) as Viewport)
                          .Where(vp => vp != null)
                          .ToList();

        foreach (ElementId elId in selIds)
        {
          Element el = doc.GetElement(elId);
          if (el == null) continue;

          bool viewDep = el.OwnerViewId != ElementId.InvalidElementId;

          foreach (Viewport vp in vports)
          {
            if (viewDep && vp.ViewId != el.OwnerViewId) continue;

            View v = doc.GetElement(vp.ViewId) as View;
            if (v == null) continue;

            BoundingBoxXYZ bb = el.get_BoundingBox(v);
            if (bb == null) continue;

            XYZ min = bb.Min - new XYZ(OFFSET_MODEL, OFFSET_MODEL, OFFSET_MODEL);
            XYZ max = bb.Max + new XYZ(OFFSET_MODEL, OFFSET_MODEL, OFFSET_MODEL);

            double sxMin = double.PositiveInfinity, syMin = double.PositiveInfinity;
            double sxMax = double.NegativeInfinity, syMax = double.NegativeInfinity;

            Transform modelToSheet = GetModelToSheetXform(v, vp, sheetId);

            foreach (XYZ d in delta)
            {
              XYZ mc = new XYZ(
                d.X == 0 ? min.X : max.X,
                d.Y == 0 ? min.Y : max.Y,
                d.Z == 0 ? min.Z : max.Z);

              XYZ sp = modelToSheet.OfPoint(bb.Transform.OfPoint(mc));

              sxMin = Math.Min(sxMin, sp.X);
              syMin = Math.Min(syMin, sp.Y);
              sxMax = Math.Max(sxMax, sp.X);
              syMax = Math.Max(syMax, sp.Y);
            }

            // rectangle in sheet space
            XYZ bl = new XYZ(sxMin, syMin, 0);
            XYZ tl = new XYZ(sxMin, syMax, 0);
            XYZ tr = new XYZ(sxMax, syMax, 0);
            XYZ br = new XYZ(sxMax, syMin, 0);

            if (!curvesBySheet.TryGetValue(sheetId, out var list))
            {
              list = new List<Curve>();
              curvesBySheet[sheetId] = list;
            }

            list.Add(Line.CreateBound(bl, tl));
            list.Add(Line.CreateBound(tl, tr));
            list.Add(Line.CreateBound(tr, br));
            list.Add(Line.CreateBound(br, bl));

            if (viewDep) break; // only one viewport needed
          }
        }
      }

      if (curvesBySheet.Count == 0)
      {
        message = "Selected elements are not visible in the chosen / detected sheets.";
        return Result.Failed;
      }

      //------------------------------------------------------------------
      // 5) Create clouds – one per sheet – with chosen revision
      //------------------------------------------------------------------
      using (Transaction t = new Transaction(doc, "Revision Clouds"))
      {
        t.Start();

        foreach (var kv in curvesBySheet)
        {
          if (kv.Value.Count > 0)
          {
            ViewSheet sh = doc.GetElement(kv.Key) as ViewSheet;
            RevisionCloud.Create(doc, sh, chosenRevision.Id, kv.Value);
          }
        }

        t.Commit();
      }

      return Result.Succeeded;
    }
  }
}
