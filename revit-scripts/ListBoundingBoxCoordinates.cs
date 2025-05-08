using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using SWF = System.Windows.Forms;          // Windows-Forms alias

namespace RevitCommands
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ListBoundingBoxCoordinates : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document   doc   = uidoc.Document;

            ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
            if (selIds.Count == 0)
            {
                TaskDialog.Show("Error", "Please select at least one element.");
                return Result.Failed;
            }

            var elementData = new List<Dictionary<string, object>>();

            //------------------------------------------------------------------
            // 1. Gather geometry + parameters
            //------------------------------------------------------------------
            foreach (ElementId id in selIds)
            {
                Element elem = doc.GetElement(id);

                XYZ minPt, maxPt;
                double width, depth, height;

                // ─── 1) Viewport on a sheet ─────────────────────────────────
                if (elem is Viewport vp)
                {
                    Outline ol = vp.GetBoxOutline();              // sheet coords
                    minPt = ol.MinimumPoint;
                    maxPt = ol.MaximumPoint;

                    width  = Math.Abs(maxPt.X - minPt.X);
                    height = Math.Abs(maxPt.Y - minPt.Y);
                    depth  = 0;
                }
                // ─── 2) Title block (no view activation) ───────────────────
                else if (elem is FamilyInstance fiTB &&
                         fiTB.Category != null &&
                         fiTB.Category.Id.IntegerValue == (int)BuiltInCategory.OST_TitleBlocks)
                {
                    BoundingBoxXYZ symBB = fiTB.Symbol.get_BoundingBox(null);
                    if (symBB == null) continue;

                    Transform tf = fiTB.GetTransform();           // sheet space
                    minPt = tf.OfPoint(symBB.Min);
                    maxPt = tf.OfPoint(symBB.Max);

                    width  = Math.Abs(maxPt.X - minPt.X);
                    height = Math.Abs(maxPt.Y - minPt.Y);
                    depth  = 0;
                }
                // ─── 3) Any other element (model or detail) ────────────────
                else
                {
                    BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                    if (bb == null) continue;

                    Transform tf = Transform.Identity;
                    FamilyInstance fi = elem as FamilyInstance;
                    if (fi != null) tf = fi.GetTransform();

                    minPt = tf.OfPoint(bb.Min);
                    maxPt = tf.OfPoint(bb.Max);

                    width  = Math.Abs(maxPt.X - minPt.X);
                    depth  = Math.Abs(maxPt.Y - minPt.Y);
                    height = Math.Abs(maxPt.Z - minPt.Z);
                }

                // Centroid in same space
                XYZ ctrPt = new XYZ((minPt.X + maxPt.X) * 0.5,
                                    (minPt.Y + maxPt.Y) * 0.5,
                                    (minPt.Z + maxPt.Z) * 0.5);

                var data = new Dictionary<string, object>
                {
                    { "ElementId", elem.Id.IntegerValue },
                    { "Name",      elem.Name },
                    { "Category",  elem.Category != null ? elem.Category.Name : "N/A" },
                    { "Family",    (elem as FamilyInstance) != null ?
                                       ((FamilyInstance)elem).Symbol.Family.Name : "N/A" },
                    { "Type",      elem.GetType().Name },

                    { "Width",     Math.Round(width  * 304.8, 2) },
                    { "Depth",     Math.Round(depth  * 304.8, 2) },
                    { "Height",    Math.Round(height * 304.8, 2) },

                    { "Min X",     Math.Round(minPt.X * 304.8, 2) },
                    { "Min Y",     Math.Round(minPt.Y * 304.8, 2) },
                    { "Min Z",     Math.Round(minPt.Z * 304.8, 2) },
                    { "Max X",     Math.Round(maxPt.X * 304.8, 2) },
                    { "Max Y",     Math.Round(maxPt.Y * 304.8, 2) },
                    { "Max Z",     Math.Round(maxPt.Z * 304.8, 2) },

                    { "Centroid X",Math.Round(ctrPt.X * 304.8, 2) },
                    { "Centroid Y",Math.Round(ctrPt.Y * 304.8, 2) },
                    { "Centroid Z",Math.Round(ctrPt.Z * 304.8, 2) },
                };

                AddCommonParameters(elem, data);
                elementData.Add(data);
            }

            if (elementData.Count == 0)
            {
                TaskDialog.Show("Info", "No data to display.");
                return Result.Succeeded;
            }

            //------------------------------------------------------------------
            // 2. Transpose -> rows = properties, columns = elements
            //------------------------------------------------------------------
            var headers       = elementData.Select(d => "Id " + d["ElementId"]).ToList();
            var propertyNames = elementData[0].Keys.ToList();

            DataTable table = new DataTable();
            table.Columns.Add("Property");
            foreach (string h in headers) table.Columns.Add(h);

            foreach (string prop in propertyNames)
            {
                DataRow row = table.NewRow();
                row["Property"] = prop;
                for (int i = 0; i < elementData.Count; ++i)
                {
                    object v;
                    if (elementData[i].TryGetValue(prop, out v)) row[i + 1] = v;
                }
                table.Rows.Add(row);
            }

            //------------------------------------------------------------------
            // 3. Show grid (empty caption) – Esc closes, Ctrl+C copies cells
            //------------------------------------------------------------------
            try
            {
                ShowGrid(table, "");           // ← no caption text
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // Copy every parameter that has a value
        private static void AddCommonParameters(Element elem,
                                                IDictionary<string, object> data)
        {
            foreach (Parameter p in elem.Parameters)
            {
                if (!p.HasValue) continue;
                string name = p.Definition.Name;
                if (data.ContainsKey(name)) continue;

                object val = null;
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        val = Math.Round(p.AsDouble() * 304.8, 2); break;
                    case StorageType.Integer:
                        val = p.AsInteger(); break;
                    case StorageType.String:
                        val = p.AsString(); break;
                    case StorageType.ElementId:
                        val = p.AsElementId().IntegerValue; break;
                }
                if (val != null) data[name] = val;
            }
        }

        // Modal WinForms grid helper
        private static void ShowGrid(DataTable table, string caption)
        {
            using (SWF.Form form = new SWF.Form())
            using (SWF.DataGridView dgv = new SWF.DataGridView())
            {
                form.Text = caption;
                form.StartPosition = SWF.FormStartPosition.CenterScreen;
                form.Width  = 1000;
                form.Height = 600;

                form.KeyPreview = true;
                form.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == SWF.Keys.Escape) form.Close();
                };

                dgv.Dock = SWF.DockStyle.Fill;
                dgv.DataSource = table;
                dgv.ReadOnly = true;
                dgv.SelectionMode = SWF.DataGridViewSelectionMode.CellSelect;
                dgv.MultiSelect = true;
                dgv.AllowUserToAddRows = false;
                dgv.ClipboardCopyMode =
                    SWF.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
                dgv.AutoSizeColumnsMode = SWF.DataGridViewAutoSizeColumnsMode.DisplayedCells;
                dgv.AutoResizeRowHeadersWidth(
                    SWF.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

                form.Controls.Add(dgv);
                form.ShowDialog();
            }
        }
    }
}
