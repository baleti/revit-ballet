// AssignViewTemplateToSelectedViews.cs
//
// C# 7.3 command for Revit 2024 – assigns a single view template
// to every view or viewport currently selected.
//
// References: RevitAPI.dll, RevitAPIUI.dll

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

namespace YourCompany.YourAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AssignViewTemplateToSelectedViews : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc   = uiApp.ActiveUIDocument;
            Document   doc     = uiDoc.Document;

            /*------------------------------------------------------------------
             * 1) Collect target views from current selection.
             *-----------------------------------------------------------------*/
            HashSet<View> targetViews = new HashSet<View>();

            foreach (ElementId id in uiDoc.GetSelectionIds())
            {
                switch (doc.GetElement(id))
                {
                    case View v when !v.IsTemplate:
                        targetViews.Add(v);
                        break;

                    case Viewport vp:
                        if (doc.GetElement(vp.ViewId) is View v2 && !v2.IsTemplate)
                            targetViews.Add(v2);
                        break;
                }
            }

            if (targetViews.Count == 0)
            {
                TaskDialog.Show("Assign View Template",
                    "Select one or more views or viewports before running this command.");
                return Result.Cancelled;
            }

            /*------------------------------------------------------------------
             * 2) Retrieve all view templates.
             *-----------------------------------------------------------------*/
            List<View> templates = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate)
                .OrderBy(v => v.Name)
                .ToList();

            if (templates.Count == 0)
            {
                TaskDialog.Show("Assign View Template",
                    "This project contains no view templates.");
                return Result.Cancelled;
            }

            /*------------------------------------------------------------------
             * 3) Show CustomGUIs.DataGrid — only Name & View Type columns.
             *-----------------------------------------------------------------*/
            List<Dictionary<string, object>> templateRows = templates
                .Select(t => new Dictionary<string, object>
                {
                    { "Name",      t.Name },
                    { "View Type", t.ViewType.ToString() }
                })
                .ToList();

            List<string> columns = new List<string> { "Name", "View Type" };
            List<Dictionary<string, object>> chosen;

            do
            {
                chosen = CustomGUIs.DataGrid(templateRows, columns, false);

                if (chosen == null || chosen.Count == 0)           // user cancelled
                    return Result.Cancelled;

                if (chosen.Count > 1)
                    TaskDialog.Show("Assign View Template",
                        "Please select **exactly one** view template.");
            }
            while (chosen.Count != 1);

            string pickedName  = chosen[0]["Name"].ToString();
            View   templateSel = templates.First(t => t.Name == pickedName);

            /*------------------------------------------------------------------
             * 4) Apply the selected template to every target view.
             *-----------------------------------------------------------------*/
            using (Transaction tx = new Transaction(doc, "Assign View Template"))
            {
                tx.Start();

                foreach (View v in targetViews)
                {
                    if (v.ViewTemplateId != templateSel.Id)
                        v.ViewTemplateId = templateSel.Id;
                }

                tx.Commit();
            }

            return Result.Succeeded;
        }
    }
}
