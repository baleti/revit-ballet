using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

namespace HideLevelBubbles
{
    [Transaction(TransactionMode.Manual)]
    public class ToggleBubblesOfSelectedGrids : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the active document and view.
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            if (uiDoc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }
            Document doc = uiDoc.Document;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            // Retrieve the currently selected elements (grids).
            ICollection<ElementId> selIds = uiDoc.GetSelectionIds();
            List<Grid> selectedGrids = new List<Grid>();

            foreach (ElementId id in selIds)
            {
                Element elem = doc.GetElement(id);
                if (elem is Grid grid)
                {
                    selectedGrids.Add(grid);
                }
            }

            if (selectedGrids.Count == 0)
            {
                message = "Please select one or more grid elements.";
                return Result.Failed;
            }

            // Display the dialog to capture user choices.
            BubbleOperation chosenOperation;
            BubbleOption chosenBubbleOption;
            using (BubbleOperationDialog dialog = new BubbleOperationDialog())
            {
                if (dialog.ShowDialog() != WinForms.DialogResult.OK)
                {
                    message = "Operation cancelled by the user.";
                    return Result.Cancelled;
                }
                chosenOperation = dialog.SelectedOperation;
                chosenBubbleOption = dialog.SelectedBubbleOption;
            }

            // Start a transaction.
            using (Transaction trans = new Transaction(doc, "Hide/Show Grid Bubbles"))
            {
                trans.Start();

                foreach (Grid grid in selectedGrids)
                {
                    DatumPlane dp = grid as DatumPlane;
                    if (dp != null)
                    {
                        try
                        {
                            // Process based on the chosen bubble option and operation.
                            switch (chosenBubbleOption)
                            {
                                case BubbleOption.End0:
                                    if (chosenOperation == BubbleOperation.Hide)
                                        dp.HideBubbleInView(DatumEnds.End0, activeView);
                                    else
                                        dp.ShowBubbleInView(DatumEnds.End0, activeView);
                                    break;
                                case BubbleOption.End1:
                                    if (chosenOperation == BubbleOperation.Hide)
                                        dp.HideBubbleInView(DatumEnds.End1, activeView);
                                    else
                                        dp.ShowBubbleInView(DatumEnds.End1, activeView);
                                    break;
                                case BubbleOption.Both:
                                    if (chosenOperation == BubbleOperation.Hide)
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
                        catch (Autodesk.Revit.Exceptions.ArgumentException)
                        {
                            // If the datum plane is not visible in this view, ignore the error.
                        }
                    }
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
}
