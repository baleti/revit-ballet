using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RevitCommands
{
    /// <summary>
    /// Synchronises the zoom rectangle of the active sheet with user‑selected sheets.
    ///
    /// * If the command is launched while a **viewport is activated**, it automatically
    ///   de‑activates (by switching back to the parent <see cref="ViewSheet"/>) and treats
    ///   that sheet as the reference view.
    /// * All target selections that point at an activated viewport are likewise promoted to
    ///   their parent sheets so the synchronisation always happens sheet‑to‑sheet.
    /// * The reference rectangle is stored as percentages of sheet width/height so it works on
    ///   sheets with different sizes/title‑blocks.
    ///
    /// Built for Revit 2023+. In older versions the helper falls back to assigning
    /// <see cref="UIDocument.ActiveView"/>.
    ///
    /// 2025‑05‑12: Reporting summary TaskDialog removed at user request – the command now runs
    /// silently unless a blocking error occurs.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SynchronizeSelectedViews : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc   = uiApp.ActiveUIDocument;
            Document   doc     = uiDoc.Document;

            try
            {
                //------------------------------------------------------------------
                // 1. Ensure we are on a sheet (de‑activate viewport if necessary)
                //------------------------------------------------------------------
                View activeView = uiDoc.ActiveView;
                activeView = PromoteToSheetIfViewportIsActivated(uiDoc, activeView);

                if (!(activeView is ViewSheet activeSheet))
                {
                    TaskDialog.Show("Synchronise Views",
                        "The active view is not a sheet nor placed on a sheet. " +
                        "Please start the command from a sheet.");
                    return Result.Failed;
                }

                UIView activeUIView = uiDoc.GetOpenUIViews()
                                            .FirstOrDefault(v => v.ViewId == activeView.Id);
                if (activeUIView == null)
                {
                    TaskDialog.Show("Synchronise Views", "Cannot access UI data for the active sheet.");
                    return Result.Failed;
                }

                IList<XYZ> sheetCorners = activeUIView.GetZoomCorners();
                if (sheetCorners == null || sheetCorners.Count < 2)
                {
                    TaskDialog.Show("Synchronise Views", "Cannot read zoom rectangle from the active sheet.");
                    return Result.Failed;
                }

                var normalised = GetNormalisedRect(activeSheet, sheetCorners[0], sheetCorners[1]);

                //------------------------------------------------------------------
                // 2. Let the user pick target sheets (promote activated viewports)
                //------------------------------------------------------------------
                List<ViewSheet> targets = CollectTargetSheets(uiDoc, activeSheet, doc);
                if (targets.Count == 0)
                {
                    TaskDialog.Show("Synchronise Views", "No target sheets selected.");
                    return Result.Cancelled;
                }

                //------------------------------------------------------------------
                // 3. Copy the rectangle to every sheet (silent on success)
                //------------------------------------------------------------------
                foreach (ViewSheet sheet in targets)
                {
                    try
                    {
                        if (!OpenAndActivateView(uiDoc, sheet))
                            continue;   // silently skip on activation failure

                        UIView ui = uiDoc.GetOpenUIViews().First(u => u.ViewId == sheet.Id);
                        XYZ[] rect = BuildRectInSheet(sheet, normalised);
                        ui.ZoomAndCenterRectangle(rect[0], rect[1]);
                    }
                    catch
                    {
                        // silently ignore individual failures – user wanted no final report
                    }
                }

                // Return to the starting sheet
                OpenAndActivateView(uiDoc, activeSheet);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        //------------------------------------------------------------------
        // Promote – if user is inside an activated viewport, switch to its sheet
        //------------------------------------------------------------------
        private static View PromoteToSheetIfViewportIsActivated(UIDocument uiDoc, View v)
        {
            if (v is ViewSheet) return v;

            Document doc = uiDoc.Document;
            ViewSheet sheet = GetParentSheet(doc, v);
            if (sheet != null)
            {
                // Open the sheet – this also de‑activates the viewport in the UI
                OpenAndActivateView(uiDoc, sheet);
                return sheet;
            }
            return v;
        }

        //------------------------------------------------------------------
        // Collect target sheets – promote any selected activated views
        //------------------------------------------------------------------
        private static List<ViewSheet> CollectTargetSheets(UIDocument uiDoc, ViewSheet referenceSheet, Document doc)
        {
            HashSet<ElementId> sheetIds = new HashSet<ElementId>();
            List<ViewSheet>    list     = new List<ViewSheet>();

            void AddIfSheet(ElementId id)
            {
                View v = doc.GetElement(id) as View;
                if (v == null) return;
                ViewSheet sheet = v as ViewSheet ?? GetParentSheet(doc, v);
                if (sheet != null && sheet.Id != referenceSheet.Id && sheetIds.Add(sheet.Id))
                    list.Add(sheet);
            }

            // Pre‑selection
            foreach (ElementId id in uiDoc.Selection.GetElementIds()) AddIfSheet(id);

            // PickObjects if none pre‑selected
            if (list.Count == 0)
            {
                try
                {
                    IList<Reference> picked = uiDoc.Selection.PickObjects(
                        ObjectType.Element,
                        new ViewSelectionFilter(),
                        "Select sheets to synchronise");
                    foreach (Reference r in picked) AddIfSheet(r.ElementId);
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            }
            return list;
        }

        //------------------------------------------------------------------
        // Open view helper (Revit 2023+ or fallback)
        //------------------------------------------------------------------
        private static bool OpenAndActivateView(UIDocument uiDoc, View view)
        {
            try
            {
#if REVIT_2023_OR_LATER
                return uiDoc.OpenAndActivateView(view);
#else
                uiDoc.ActiveView = view;
                return true;
#endif
            }
            catch { return false; }
        }

        //------------------------------------------------------------------
        // Find the sheet that hosts a given placed view (viewport) – null if none
        //------------------------------------------------------------------
        private static ViewSheet GetParentSheet(Document doc, View view)
        {
            Viewport vp = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .FirstOrDefault(p => p.ViewId == view.Id);
            return vp == null ? null : doc.GetElement(vp.SheetId) as ViewSheet;
        }

        //------------------------------------------------------------------
        // Normalise rectangle to sheet UV space (0‑1)
        //------------------------------------------------------------------
        private static (double u0, double v0, double u1, double v1) GetNormalisedRect(
            ViewSheet sheet, XYZ p1, XYZ p2)
        {
            BoundingBoxUV bb = sheet.Outline;
            double w = bb.Max.U - bb.Min.U;
            double h = bb.Max.V - bb.Min.V;

            double u1 = (p1.X - bb.Min.U) / w;
            double v1 = (p1.Y - bb.Min.V) / h;
            double u2 = (p2.X - bb.Min.U) / w;
            double v2 = (p2.Y - bb.Min.V) / h;

            return (
                Math.Min(u1, u2),
                Math.Min(v1, v2),
                Math.Max(u1, u2),
                Math.Max(v1, v2));
        }

        //------------------------------------------------------------------
        // Rebuild rectangle on another sheet from 0‑1 UV coords
        //------------------------------------------------------------------
        private static XYZ[] BuildRectInSheet(ViewSheet sheet, (double u0, double v0, double u1, double v1) n)
        {
            BoundingBoxUV bb = sheet.Outline;
            double w = bb.Max.U - bb.Min.U;
            double h = bb.Max.V - bb.Min.V;

            double x0 = bb.Min.U + n.u0 * w;
            double y0 = bb.Min.V + n.v0 * h;
            double x1 = bb.Min.U + n.u1 * w;
            double y1 = bb.Min.V + n.v1 * h;

            return new[]
            {
                new XYZ(x0, y0, 0),
                new XYZ(x1, y1, 0)
            };
        }
    }

    //------------------------------------------------------------------
    // Selection filter – allow any view or viewport; schedules excluded
    //------------------------------------------------------------------
    public class ViewSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) =>
            elem is View && !(elem is ViewSchedule);
        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}
