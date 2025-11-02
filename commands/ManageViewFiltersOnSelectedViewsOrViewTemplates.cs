// File: FilterCommands.cs
// Place this single file in your project.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
// Aliases to disambiguate Windows Forms and Drawing types.
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace MyRevitCommands
{
   /// <summary>
   /// Shared helper methods for filter commands.
   /// </summary>
   public static class FilterCommandHelper
   {
      // Returns valid views (or views referenced by viewports) from the current selection.
      public static List<View> GetSelectedViews(UIDocument uiDoc, Document doc)
      {
         List<View> selectedViews = new List<View>();
         ICollection<ElementId> selectedIds = uiDoc.GetSelectionIds();
         foreach (ElementId id in selectedIds)
         {
            Element element = doc.GetElement(id);
            if (element is Viewport viewport)
            {
               ElementId viewId = viewport.ViewId;
               View view = doc.GetElement(viewId) as View;
               if (view != null && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
                  selectedViews.Add(view);
            }
            else if (element is View view && !(view is ViewSchedule) && view.ViewType != ViewType.DrawingSheet)
            {
               selectedViews.Add(view);
            }
         }
         return selectedViews;
      }

      // Returns target view IDs. If a view is a template then use its own ID;
      // otherwise if it has a view template assigned, use that; else use the view.
      public static HashSet<ElementId> GetTargetViewIds(List<View> views)
      {
         HashSet<ElementId> targetIds = new HashSet<ElementId>();
         foreach (View view in views)
         {
            if (view.IsTemplate)
               targetIds.Add(view.Id);
            else if (view.ViewTemplateId != null && view.ViewTemplateId != ElementId.InvalidElementId)
               targetIds.Add(view.ViewTemplateId);
            else
               targetIds.Add(view.Id);
         }
         return targetIds;
      }

      // Builds grid entries for a collection of ParameterFilterElement.
      // Extra keys ("Rule Category", "Rule Subject", "Rule Verb", "Rule Value")
      // are populated using reflection to call the legacy GetRules() method and to obtain rule details.
      public static List<Dictionary<string, object>> BuildFilterGridEntries(Document doc, IEnumerable<ParameterFilterElement> filters)
      {
         List<Dictionary<string, object>> entries = new List<Dictionary<string, object>>();
         foreach (ParameterFilterElement filter in filters)
         {
            var entry = new Dictionary<string, object>();
            entry["Name"] = filter.Name;
            entry["Id"] = filter.Id.IntegerValue;
            
            // Use reflection to attempt to call the legacy GetRules() method.
            IList<FilterRule> rules = null;
            MethodInfo mi = filter.GetType().GetMethod("GetRules", BindingFlags.Instance | BindingFlags.Public);
            if (mi != null)
            {
               object result = mi.Invoke(filter, null);
               rules = result as IList<FilterRule>;
            }
            
            if (rules != null && rules.Count > 0)
            {
               FilterRule rule = rules[0];

               // Retrieve the parameter ID via reflection.
               ElementId paramId = ElementId.InvalidElementId;
               PropertyInfo paramIdProp = rule.GetType().GetProperty("ParameterId");
               if (paramIdProp != null)
               {
                  object val = paramIdProp.GetValue(rule, null);
                  if (val is ElementId idVal)
                     paramId = idVal;
               }
               ParameterElement paramElem = doc.GetElement(paramId) as ParameterElement;
               
               // Get a group (category) string from the parameter.
               string ruleCategory = "n/a";
               if (paramElem != null)
               {
                  PropertyInfo groupProp = paramElem.GetType().GetProperty("Group");
                  if (groupProp != null)
                  {
                     object grpVal = groupProp.GetValue(paramElem, null);
                     ruleCategory = grpVal != null ? grpVal.ToString() : "";
                  }
                  else
                     ruleCategory = "";
               }
               
               // Use the parameter's name as the subject.
               string ruleSubject = paramElem != null ? paramElem.Name : "";
               
               // Get the operator from the rule.
               string ruleVerb = "n/a";
               PropertyInfo opProp = rule.GetType().GetProperty("RuleOperator");
               if (opProp != null)
               {
                  object opVal = opProp.GetValue(rule, null);
                  ruleVerb = opVal != null ? opVal.ToString() : "";
               }
               
               // Get the rule's value.
               string ruleValue = "n/a";
               PropertyInfo valueProp = rule.GetType().GetProperty("Value");
               if (valueProp != null)
               {
                  object val = valueProp.GetValue(rule, null);
                  ruleValue = val != null ? val.ToString() : "";
               }
               
               entry["Rule Category"] = ruleCategory;
               entry["Rule Subject"] = ruleSubject;
               entry["Rule Verb"] = ruleVerb;
               entry["Rule Value"] = ruleValue;
            }
            else
            {
               entry["Rule Category"] = "";
               entry["Rule Subject"] = "";
               entry["Rule Verb"] = "";
               entry["Rule Value"] = "";
            }
            entries.Add(entry);
         }
         return entries;
      }

      // Returns all ParameterFilterElement objects applied on the given target view IDs.
      public static List<ParameterFilterElement> GetAppliedFilterElements(Document doc, HashSet<ElementId> targetIds)
      {
         HashSet<ElementId> appliedIds = new HashSet<ElementId>();
         foreach (ElementId id in targetIds)
         {
            View view = doc.GetElement(id) as View;
            if (view != null)
            {
               foreach (ElementId fid in view.GetFilters())
                  appliedIds.Add(fid);
            }
         }
         List<ParameterFilterElement> filters = new List<ParameterFilterElement>();
         foreach (ElementId fid in appliedIds)
         {
            ParameterFilterElement filterElem = doc.GetElement(fid) as ParameterFilterElement;
            if (filterElem != null)
               filters.Add(filterElem);
         }
         return filters;
      }
   }

   /// <summary>
   /// A common override options dialog displayed as a single‐row DataGridView.
   /// The grid has 8 columns (with headers) in the following order:
   ///   0: Enable Filter (checkbox)
   ///   1: Visibility (checkbox)
   ///   2: Projection Override (checkbox)
   ///   3: Projection Line Color (button cell; default white)
   ///   4: Projection Transparency (numeric text cell; default "0")
   ///   5: Cut Override (checkbox)
   ///   6: Cut Line Color (button cell; default white)
   ///   7: Halftone (checkbox)
   /// </summary>
   public class FilterOverrideDataGridForm : WinForms.Form
   {
      private WinForms.DataGridView dataGridView;
      private WinForms.Button btnOK;
      private WinForms.Button btnCancel;

      // Column indices.
      private const int COL_ENABLE_FILTER = 0;
      private const int COL_VISIBILITY = 1;
      private const int COL_PROJ_OVERRIDE = 2;
      private const int COL_PROJ_COLOR = 3;
      private const int COL_PROJ_TRANSP = 4;
      private const int COL_CUT_OVERRIDE = 5;
      private const int COL_CUT_COLOR = 6;
      private const int COL_HALFTONE = 7;

      public FilterOverrideDataGridForm()
      {
         this.Text = "Filter Override Options";
         this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
         this.StartPosition = WinForms.FormStartPosition.CenterScreen;
         this.ClientSize = new Drawing.Size(750, 150);
         this.MaximizeBox = false;
         this.MinimizeBox = false;

         dataGridView = new WinForms.DataGridView();
         dataGridView.Dock = WinForms.DockStyle.Top;
         dataGridView.Height = 70;
         dataGridView.AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.Fill;
         dataGridView.RowHeadersVisible = false;
         dataGridView.AllowUserToAddRows = false;
         dataGridView.AllowUserToDeleteRows = false;
         dataGridView.SelectionMode = WinForms.DataGridViewSelectionMode.CellSelect;
         dataGridView.MultiSelect = false;
         dataGridView.EditMode = WinForms.DataGridViewEditMode.EditOnEnter;

         // Add columns.
         dataGridView.Columns.Add(new WinForms.DataGridViewCheckBoxColumn() { HeaderText = "Enable Filter", Name = "colEnable" });
         dataGridView.Columns.Add(new WinForms.DataGridViewCheckBoxColumn() { HeaderText = "Visibility", Name = "colVisibility" });
         dataGridView.Columns.Add(new WinForms.DataGridViewCheckBoxColumn() { HeaderText = "Projection Override", Name = "colProjOverride" });
         dataGridView.Columns.Add(new WinForms.DataGridViewButtonColumn() { HeaderText = "Projection Line Color", Name = "colProjColor" });
         dataGridView.Columns.Add(new WinForms.DataGridViewTextBoxColumn() { HeaderText = "Projection Transparency", Name = "colProjTransp" });
         dataGridView.Columns.Add(new WinForms.DataGridViewCheckBoxColumn() { HeaderText = "Cut Override", Name = "colCutOverride" });
         dataGridView.Columns.Add(new WinForms.DataGridViewButtonColumn() { HeaderText = "Cut Line Color", Name = "colCutColor" });
         dataGridView.Columns.Add(new WinForms.DataGridViewCheckBoxColumn() { HeaderText = "Halftone", Name = "colHalftone" });

         // Add a single row with default values.
         dataGridView.Rows.Add();
         var row = dataGridView.Rows[0];
         row.Cells[COL_ENABLE_FILTER].Value = true;
         row.Cells[COL_VISIBILITY].Value = true;
         row.Cells[COL_PROJ_OVERRIDE].Value = false;
         var projColorCell = row.Cells[COL_PROJ_COLOR] as WinForms.DataGridViewButtonCell;
         projColorCell.Value = "Click to pick";
         projColorCell.Style.BackColor = Drawing.Color.White;
         projColorCell.Tag = Drawing.Color.White;
         row.Cells[COL_PROJ_TRANSP].Value = "0";
         row.Cells[COL_CUT_OVERRIDE].Value = false;
         var cutColorCell = row.Cells[COL_CUT_COLOR] as WinForms.DataGridViewButtonCell;
         cutColorCell.Value = "Click to pick";
         cutColorCell.Style.BackColor = Drawing.Color.White;
         cutColorCell.Tag = Drawing.Color.White;
         row.Cells[COL_HALFTONE].Value = false;

         dataGridView.CellContentClick += DataGridView_CellContentClick;

         btnOK = new WinForms.Button();
         btnOK.Text = "OK";
         btnOK.DialogResult = WinForms.DialogResult.OK;
         btnOK.Location = new Drawing.Point(500, 90);
         btnOK.Size = new Drawing.Size(75, 30);

         btnCancel = new WinForms.Button();
         btnCancel.Text = "Cancel";
         btnCancel.DialogResult = WinForms.DialogResult.Cancel;
         btnCancel.Location = new Drawing.Point(585, 90);
         btnCancel.Size = new Drawing.Size(75, 30);

         this.Controls.Add(dataGridView);
         this.Controls.Add(btnOK);
         this.Controls.Add(btnCancel);

         this.AcceptButton = btnOK;
         this.CancelButton = btnCancel;
      }

      private void DataGridView_CellContentClick(object sender, WinForms.DataGridViewCellEventArgs e)
      {
         if (e.RowIndex >= 0 && (e.ColumnIndex == COL_PROJ_COLOR || e.ColumnIndex == COL_CUT_COLOR))
         {
            var cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex] as WinForms.DataGridViewButtonCell;
            using (var cd = new WinForms.ColorDialog())
            {
               cd.Color = (Drawing.Color)cell.Tag;
               if (cd.ShowDialog() == WinForms.DialogResult.OK)
               {
                  cell.Tag = cd.Color;
                  cell.Style.BackColor = cd.Color;
               }
            }
         }
      }

      // Expose user settings.
      public bool EnableFilter { get { return Convert.ToBoolean(dataGridView.Rows[0].Cells[COL_ENABLE_FILTER].Value); } }
      public bool Visibility { get { return Convert.ToBoolean(dataGridView.Rows[0].Cells[COL_VISIBILITY].Value); } }
      public bool ProjectionOverride { get { return Convert.ToBoolean(dataGridView.Rows[0].Cells[COL_PROJ_OVERRIDE].Value); } }
      public Drawing.Color ProjectionLineColor
      {
         get
         {
            var cell = dataGridView.Rows[0].Cells[COL_PROJ_COLOR] as WinForms.DataGridViewButtonCell;
            return (Drawing.Color)cell.Tag;
         }
      }
      public int ProjectionTransparency
      {
         get
         {
            int val = 0;
            int.TryParse(dataGridView.Rows[0].Cells[COL_PROJ_TRANSP].Value.ToString(), out val);
            return val;
         }
      }
      public bool CutOverride { get { return Convert.ToBoolean(dataGridView.Rows[0].Cells[COL_CUT_OVERRIDE].Value); } }
      public Drawing.Color CutLineColor
      {
         get
         {
            var cell = dataGridView.Rows[0].Cells[COL_CUT_COLOR] as WinForms.DataGridViewButtonCell;
            return (Drawing.Color)cell.Tag;
         }
      }
      public bool Halftone { get { return Convert.ToBoolean(dataGridView.Rows[0].Cells[COL_HALFTONE].Value); } }
   }

   //--------------------------------------------------------------------------
   // Command: AddFiltersToSelectedViewOrViewTemplates
   // Adds one or more project filters (with override settings) to the selected views (or their view templates).
   [Transaction(TransactionMode.Manual)]
   public class AddViewFiltersToSelectedViewOrViewTemplates : IExternalCommand
   {
      public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
      {
         UIDocument uiDoc = commandData.Application.ActiveUIDocument;
         Document doc = uiDoc.Document;

         List<View> selectedViews = FilterCommandHelper.GetSelectedViews(uiDoc, doc);
         if (selectedViews.Count == 0)
         {
            TaskDialog.Show("Error", "No valid views or view templates selected.");
            return Result.Failed;
         }

         // Get all project filters.
         FilteredElementCollector collector = new FilteredElementCollector(doc);
         List<ParameterFilterElement> projectFilters = collector
            .OfClass(typeof(ParameterFilterElement))
            .Cast<ParameterFilterElement>()
            .ToList();
         if (projectFilters.Count == 0)
         {
            TaskDialog.Show("Error", "No project filters available.");
            return Result.Failed;
         }

         // Build grid entries (with extra rule details).
         List<Dictionary<string, object>> gridEntries = FilterCommandHelper.BuildFilterGridEntries(doc, projectFilters);
         List<string> propertyNames = new List<string> { "Name", "Id", "Rule Category", "Rule Subject", "Rule Verb", "Rule Value" };

         // Show a DataGrid for filter selection (using your helper).
         List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(gridEntries, propertyNames, spanAllScreens: false);
         if (selectedEntries == null || selectedEntries.Count == 0)
         {
            TaskDialog.Show("Cancelled", "No filters selected.");
            return Result.Cancelled;
         }
         List<ElementId> selectedFilterIds = new List<ElementId>();
         foreach (var entry in selectedEntries)
         {
            if (entry.TryGetValue("Id", out object idObj) && int.TryParse(idObj.ToString(), out int intId))
               selectedFilterIds.Add(new ElementId(intId));
         }

         // Show the override options dialog.
         using (FilterOverrideDataGridForm overrideForm = new FilterOverrideDataGridForm())
         {
            if (overrideForm.ShowDialog() != WinForms.DialogResult.OK)
               return Result.Cancelled;

            // Build override settings.
            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            if (overrideForm.ProjectionOverride)
            {
               Autodesk.Revit.DB.Color projColor = new Autodesk.Revit.DB.Color(
                  overrideForm.ProjectionLineColor.R, overrideForm.ProjectionLineColor.G, overrideForm.ProjectionLineColor.B);
               ogs.SetProjectionLineColor(projColor);
               ogs.SetSurfaceTransparency(overrideForm.ProjectionTransparency);
               ogs.SetHalftone(overrideForm.Halftone);
            }
            if (overrideForm.CutOverride)
            {
               Autodesk.Revit.DB.Color cutColor = new Autodesk.Revit.DB.Color(
                  overrideForm.CutLineColor.R, overrideForm.CutLineColor.G, overrideForm.CutLineColor.B);
               ogs.SetCutLineColor(cutColor);
            }
            bool visibility = overrideForm.Visibility;
            bool enableFilter = overrideForm.EnableFilter;

            HashSet<ElementId> targetViewIds = FilterCommandHelper.GetTargetViewIds(selectedViews);

            using (Transaction trans = new Transaction(doc, "Add Filters with Overrides"))
            {
               trans.Start();
               foreach (ElementId targetId in targetViewIds)
               {
                  View targetView = doc.GetElement(targetId) as View;
                  if (targetView != null)
                  {
                     foreach (ElementId filterId in selectedFilterIds)
                     {
                        if (enableFilter)
                        {
                           if (!targetView.GetFilters().Contains(filterId))
                              targetView.AddFilter(filterId);
                           targetView.SetFilterVisibility(filterId, visibility);
                           targetView.SetFilterOverrides(filterId, ogs);
                        }
                        else
                        {
                           if (targetView.GetFilters().Contains(filterId))
                              targetView.RemoveFilter(filterId);
                        }
                     }
                  }
               }
               trans.Commit();
            }
         }
         TaskDialog.Show("Success", "Filters added with override settings.");
         return Result.Succeeded;
      }
   }

   //--------------------------------------------------------------------------
   // Command: SetFiltersOnSelectedViewsOrViewTemplates
   // Updates (or resets) override settings for filters already applied on the selected views (or their view templates).
   [Transaction(TransactionMode.Manual)]
   public class SetViewFiltersOnSelectedViewsOrViewTemplates : IExternalCommand
   {
      public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
      {
         UIDocument uiDoc = commandData.Application.ActiveUIDocument;
         Document doc = uiDoc.Document;

         List<View> selectedViews = FilterCommandHelper.GetSelectedViews(uiDoc, doc);
         if (selectedViews.Count == 0)
         {
            TaskDialog.Show("Error", "No valid views or view templates selected.");
            return Result.Failed;
         }
         HashSet<ElementId> targetViewIds = FilterCommandHelper.GetTargetViewIds(selectedViews);
         List<ParameterFilterElement> appliedFilters = FilterCommandHelper.GetAppliedFilterElements(doc, targetViewIds);
         if (appliedFilters.Count == 0)
         {
            TaskDialog.Show("Error", "No filters are applied on the selected view(s) or their view templates.");
            return Result.Failed;
         }

         List<Dictionary<string, object>> gridEntries = FilterCommandHelper.BuildFilterGridEntries(doc, appliedFilters);
         List<string> propertyNames = new List<string> { "Name", "Id", "Rule Category", "Rule Subject", "Rule Verb", "Rule Value" };
         List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(gridEntries, propertyNames, spanAllScreens: false);
         if (selectedEntries == null || selectedEntries.Count == 0)
         {
            TaskDialog.Show("Cancelled", "No filters selected for modification.");
            return Result.Cancelled;
         }
         List<ElementId> selectedFilterIds = new List<ElementId>();
         foreach (var entry in selectedEntries)
         {
            if (entry.TryGetValue("Id", out object idObj) && int.TryParse(idObj.ToString(), out int intId))
               selectedFilterIds.Add(new ElementId(intId));
         }

         using (FilterOverrideDataGridForm overrideForm = new FilterOverrideDataGridForm())
         {
            if (overrideForm.ShowDialog() != WinForms.DialogResult.OK)
               return Result.Cancelled;

            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            if (overrideForm.ProjectionOverride)
            {
               Autodesk.Revit.DB.Color projColor = new Autodesk.Revit.DB.Color(
                  overrideForm.ProjectionLineColor.R, overrideForm.ProjectionLineColor.G, overrideForm.ProjectionLineColor.B);
               ogs.SetProjectionLineColor(projColor);
               ogs.SetSurfaceTransparency(overrideForm.ProjectionTransparency);
               ogs.SetHalftone(overrideForm.Halftone);
            }
            if (overrideForm.CutOverride)
            {
               Autodesk.Revit.DB.Color cutColor = new Autodesk.Revit.DB.Color(
                  overrideForm.CutLineColor.R, overrideForm.CutLineColor.G, overrideForm.CutLineColor.B);
               ogs.SetCutLineColor(cutColor);
            }
            bool visibility = overrideForm.Visibility;
            bool enableFilter = overrideForm.EnableFilter;

            using (Transaction trans = new Transaction(doc, "Set Filter Overrides"))
            {
               trans.Start();
               foreach (ElementId targetId in targetViewIds)
               {
                  View targetView = doc.GetElement(targetId) as View;
                  if (targetView != null)
                  {
                     foreach (ElementId filterId in selectedFilterIds)
                     {
                        if (targetView.GetFilters().Contains(filterId))
                        {
                           if (enableFilter)
                           {
                              targetView.SetFilterVisibility(filterId, visibility);
                              targetView.SetFilterOverrides(filterId, ogs);
                           }
                           else
                           {
                              // Reset overrides by applying an empty OverrideGraphicSettings.
                              targetView.SetFilterOverrides(filterId, new OverrideGraphicSettings());
                           }
                        }
                     }
                  }
               }
               trans.Commit();
            }
         }
         TaskDialog.Show("Success", "Filter override settings updated.");
         return Result.Succeeded;
      }
   }

   //--------------------------------------------------------------------------
   // Command: DeleteFiltersFromSelectedViewsOrViewTemplates
   // Removes filters (that are currently applied) from the selected views (or their view templates).
   [Transaction(TransactionMode.Manual)]
   public class DeleteViewFiltersFromSelectedViewsOrViewTemplates : IExternalCommand
   {
      public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
      {
         UIDocument uiDoc = commandData.Application.ActiveUIDocument;
         Document doc = uiDoc.Document;

         List<View> selectedViews = FilterCommandHelper.GetSelectedViews(uiDoc, doc);
         if (selectedViews.Count == 0)
         {
            TaskDialog.Show("Error", "No valid views or view templates selected.");
            return Result.Failed;
         }
         HashSet<ElementId> targetViewIds = FilterCommandHelper.GetTargetViewIds(selectedViews);
         List<ParameterFilterElement> appliedFilters = FilterCommandHelper.GetAppliedFilterElements(doc, targetViewIds);
         if (appliedFilters.Count == 0)
         {
            TaskDialog.Show("Error", "No filters are applied on the selected view(s) or their view templates.");
            return Result.Failed;
         }

         List<Dictionary<string, object>> gridEntries = FilterCommandHelper.BuildFilterGridEntries(doc, appliedFilters);
         List<string> propertyNames = new List<string> { "Name", "Id", "Rule Category", "Rule Subject", "Rule Verb", "Rule Value" };
         List<Dictionary<string, object>> selectedEntries = CustomGUIs.DataGrid(gridEntries, propertyNames, spanAllScreens: false);
         if (selectedEntries == null || selectedEntries.Count == 0)
         {
            TaskDialog.Show("Cancelled", "No filters selected for deletion.");
            return Result.Cancelled;
         }
         List<ElementId> selectedFilterIds = new List<ElementId>();
         foreach (var entry in selectedEntries)
         {
            if (entry.TryGetValue("Id", out object idObj) && int.TryParse(idObj.ToString(), out int intId))
               selectedFilterIds.Add(new ElementId(intId));
         }

         using (Transaction trans = new Transaction(doc, "Delete Filters"))
         {
            trans.Start();
            foreach (ElementId targetId in targetViewIds)
            {
               View targetView = doc.GetElement(targetId) as View;
               if (targetView != null)
               {
                  foreach (ElementId filterId in selectedFilterIds)
                  {
                     if (targetView.GetFilters().Contains(filterId))
                        targetView.RemoveFilter(filterId);
                  }
               }
            }
            trans.Commit();
         }
         TaskDialog.Show("Success", "Selected filters have been deleted from the selected view(s) or view templates.");
         return Result.Succeeded;
      }
   }
}
