using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

using TaskDialog = Autodesk.Revit.UI.TaskDialog;

public enum BubbleOperation { Hide, Show }
public enum BubbleOption    { End0, End1, Both }

public class BubbleOperationDialog : WinForms.Form
{
    public BubbleOperation SelectedOperation { get; private set; }
    public BubbleOption SelectedBubbleOption { get; private set; }

    private WinForms.RadioButton rbHide, rbShow, rbEnd0, rbEnd1, rbBoth;

    public BubbleOperationDialog()
    {
        Text = "Hide or Show Bubbles";
        Width = 350; Height = 300;
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        StartPosition = WinForms.FormStartPosition.CenterScreen;
        MaximizeBox = false; MinimizeBox = false;

        var grpOp = new WinForms.GroupBox { Text = "Operation", Left = 20, Top = 20, Width = 290, Height = 70 };
        rbHide = new WinForms.RadioButton { Text = "Hide Bubbles", Left = 20, Top = 30, Width = 120, Checked = true };
        rbShow = new WinForms.RadioButton { Text = "Show Bubbles", Left = 150, Top = 30, Width = 120 };
        grpOp.Controls.Add(rbHide);
        grpOp.Controls.Add(rbShow);

        var grpEnds = new WinForms.GroupBox { Text = "Bubble End(s)", Left = 20, Top = 110, Width = 290, Height = 100 };
        rbEnd0 = new WinForms.RadioButton { Text = "End0", Left = 20, Top = 30, Width = 80 };
        rbEnd1 = new WinForms.RadioButton { Text = "End1", Left = 110, Top = 30, Width = 80 };
        rbBoth = new WinForms.RadioButton { Text = "Both", Left = 200, Top = 30, Width = 80, Checked = true };
        grpEnds.Controls.Add(rbEnd0);
        grpEnds.Controls.Add(rbEnd1);
        grpEnds.Controls.Add(rbBoth);

        var btnOK     = new WinForms.Button { Text = "OK",     Left = 70,  Width = 80, Top = 230, DialogResult = WinForms.DialogResult.OK };
        var btnCancel = new WinForms.Button { Text = "Cancel", Left = 180, Width = 80, Top = 230, DialogResult = WinForms.DialogResult.Cancel };

        Controls.Add(grpOp);
        Controls.Add(grpEnds);
        Controls.Add(btnOK);
        Controls.Add(btnCancel);
        AcceptButton = btnOK;
        CancelButton = btnCancel;

        btnOK.Click += (s, e) =>
        {
            SelectedOperation  = rbHide.Checked ? BubbleOperation.Hide : BubbleOperation.Show;
            SelectedBubbleOption = rbEnd0.Checked ? BubbleOption.End0
                                 : rbEnd1.Checked ? BubbleOption.End1
                                 : BubbleOption.Both;
            DialogResult = WinForms.DialogResult.OK;
            Close();
        };
        btnCancel.Click += (s, e) => { DialogResult = WinForms.DialogResult.Cancel; Close(); };
    }
}

[Transaction(TransactionMode.Manual)]
[CommandMeta("Grid, Level")]
public class ToggleBubbles : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;
        Autodesk.Revit.DB.View activeView = doc.ActiveView;

        List<DatumPlane> selected = new List<DatumPlane>();

        foreach (ElementId id in uiDoc.GetSelectionIds())
        {
            Element elem = doc.GetElement(id);
            if (elem is Grid || elem is Level)
                selected.Add((DatumPlane)elem);
        }

        if (selected.Count == 0)
        {
            var items = new List<Dictionary<string, object>>();
            foreach (Grid g in new FilteredElementCollector(doc).OfClass(typeof(Grid)).Cast<Grid>())
                items.Add(new Dictionary<string, object> { { "Name", g.Name }, { "Type", "Grid" }, { "ElementIdObject", g.Id } });
            foreach (Level l in new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>())
                items.Add(new Dictionary<string, object> { { "Name", l.Name }, { "Type", "Level" }, { "ElementIdObject", l.Id } });

            CustomGUIs.SetCurrentUIDocument(uiDoc);
            var chosen = CustomGUIs.DataGrid(items, new List<string> { "Name", "Type" }, false);
            if (chosen == null || chosen.Count == 0) return Result.Cancelled;

            foreach (var row in chosen)
            {
                if (row.TryGetValue("ElementIdObject", out object idObj) && idObj is ElementId eid)
                {
                    if (doc.GetElement(eid) is DatumPlane dp)
                        selected.Add(dp);
                }
            }
            if (selected.Count == 0) return Result.Succeeded;
        }

        BubbleOperation chosenOp;
        BubbleOption chosenEnd;
        using (var dialog = new BubbleOperationDialog())
        {
            if (dialog.ShowDialog() != WinForms.DialogResult.OK)
                return Result.Cancelled;
            chosenOp  = dialog.SelectedOperation;
            chosenEnd = dialog.SelectedBubbleOption;
        }

        using (Transaction trans = new Transaction(doc, "Hide/Show Bubbles"))
        {
            trans.Start();
            foreach (DatumPlane dp in selected)
            {
                try
                {
                    switch (chosenEnd)
                    {
                        case BubbleOption.End0:
                            if (chosenOp == BubbleOperation.Hide) dp.HideBubbleInView(DatumEnds.End0, activeView);
                            else                                   dp.ShowBubbleInView(DatumEnds.End0, activeView);
                            break;
                        case BubbleOption.End1:
                            if (chosenOp == BubbleOperation.Hide) dp.HideBubbleInView(DatumEnds.End1, activeView);
                            else                                   dp.ShowBubbleInView(DatumEnds.End1, activeView);
                            break;
                        case BubbleOption.Both:
                            if (chosenOp == BubbleOperation.Hide)
                            {
                                dp.HideBubbleInView(DatumEnds.End0, activeView);
                                dp.HideBubbleInView(DatumEnds.End1, activeView);
                            }
                            else
                            {
                                dp.ShowBubbleInView(DatumEnds.End0, activeView);
                                dp.ShowBubbleInView(DatumEnds.End1, activeView);
                            }
                            break;
                    }
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException) { }
            }
            trans.Commit();
        }
        return Result.Succeeded;
    }
}
