using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using WinForms = System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CreateViewSetFromSelectedViews : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        UIDocument uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;
        
        // 1️⃣ Get currently selected elements using SelectionModeManager
        ICollection<ElementId> selectedIds = uidoc.GetSelectionIds();
        
        if (selectedIds == null || selectedIds.Count == 0)
        {
            TaskDialog.Show("Create View Set", "No elements are currently selected.");
            return Result.Cancelled;
        }
        
        // 2️⃣ Filter selection to get only views (including sheets)
        List<View> selectedViews = new List<View>();
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem is View view && !view.IsTemplate)
            {
                selectedViews.Add(view);
            }
        }
        
        if (selectedViews.Count == 0)
        {
            TaskDialog.Show("Create View Set", "No views or sheets found in the current selection.");
            return Result.Cancelled;
        }
        
        // 3️⃣ Get name for the new view set
        string newSetName;
        using (var dlg = new ViewSetNameForm())
        {
            if (dlg.ShowDialog() != WinForms.DialogResult.OK)
                return Result.Cancelled;
            newSetName = dlg.SetName.Trim();
        }
        
        if (string.IsNullOrEmpty(newSetName))
        {
            message = "You must supply a name for the new view set.";
            return Result.Failed;
        }
        
        // 4️⃣ Check if a set with this name already exists
        List<ViewSheetSet> allSets = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheetSet))
            .Cast<ViewSheetSet>()
            .ToList();
            
        ViewSheetSet existingSet = allSets.FirstOrDefault(s => 
            s.Name.Equals(newSetName, StringComparison.OrdinalIgnoreCase));
            
        bool overwriteExisting = false;
        if (existingSet != null)
        {
            WinForms.DialogResult res = WinForms.MessageBox.Show(
                $"A view set named \"{newSetName}\" already exists.\n\n" +
                "Do you want to overwrite it?",
                "Overwrite Existing Set?",
                WinForms.MessageBoxButtons.YesNo,
                WinForms.MessageBoxIcon.Question,
                WinForms.MessageBoxDefaultButton.Button2);
                
            if (res != WinForms.DialogResult.Yes)
                return Result.Cancelled;
                
            overwriteExisting = true;
        }
        
        // 5️⃣ Create ViewSet from selected views
        ViewSet viewSet = new ViewSet();
        foreach (View view in selectedViews)
        {
            viewSet.Insert(view);
        }
        
        // 6️⃣ Create or update the view set
        using (Transaction tx = new Transaction(doc, 
            overwriteExisting ? "Overwrite view set" : "Create view set from selection"))
        {
            tx.Start();
            
            PrintManager pm = doc.PrintManager;
            pm.PrintRange = PrintRange.Select;
            pm.PrintSetup.CurrentPrintSetting = pm.PrintSetup.InSession;
            ViewSheetSetting vss = pm.ViewSheetSetting;
            
            if (overwriteExisting)
            {
                vss.CurrentViewSheetSet = existingSet;
                vss.CurrentViewSheetSet.Views = viewSet;
                vss.Save();
            }
            else
            {
                vss.CurrentViewSheetSet.Views = viewSet;
                vss.SaveAs(newSetName);
            }
            
            tx.Commit();
        }
        
        // 7️⃣ Success message
        string action = overwriteExisting ? "updated" : "created";
        TaskDialog.Show("Create View Set", 
            $"Successfully {action} view set \"{newSetName}\" with {selectedViews.Count} view(s).");
        
        return Result.Succeeded;
    }
}

// ──────────────────────────────────────────────────────────
//  Simple WinForms dialog for view set name input
// ──────────────────────────────────────────────────────────
public class ViewSetNameForm : WinForms.Form
{
    private readonly WinForms.TextBox _txtName;
    
    public string SetName => _txtName.Text;
    
    public ViewSetNameForm()
    {
        // Form setup
        Text = "Create View Set";
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        StartPosition = WinForms.FormStartPosition.CenterScreen;
        MinimizeBox = false;
        MaximizeBox = false;
        Width = 330;
        Height = 140;
        
        // Label for name input
        var lblName = new WinForms.Label
        {
            Text = "Enter a name for the new view set:",
            AutoSize = true,
            Left = 10,
            Top = 15
        };
        Controls.Add(lblName);
        
        // TextBox for name input
        _txtName = new WinForms.TextBox
        {
            Left = 10,
            Top = 40,
            Width = 300
        };
        Controls.Add(_txtName);
        
        // OK button
        var btnOK = new WinForms.Button
        {
            Text = "OK",
            DialogResult = WinForms.DialogResult.OK,
            Left = 155,
            Width = 70,
            Top = 75
        };
        Controls.Add(btnOK);
        
        // Cancel button
        var btnCancel = new WinForms.Button
        {
            Text = "Cancel",
            DialogResult = WinForms.DialogResult.Cancel,
            Left = 235,
            Width = 70,
            Top = 75
        };
        Controls.Add(btnCancel);
        
        // Set form defaults
        AcceptButton = btnOK;
        CancelButton = btnCancel;
        
        // Focus on the text box
        _txtName.Select();
    }
}
