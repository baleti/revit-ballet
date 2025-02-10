using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[Transaction(TransactionMode.ReadOnly)]
public class DeselectRandomlySimplex : IExternalCommand
{
    private class SimplexNoise
    {
        private static readonly int[] Perm = {
            151,160,137,91,90,15,131,13,201,95,96,53,194,233,7,225,140,36,103,30,69,142,8,99,37,240,21,10,23,190,6,148,
            247,120,234,75,0,26,197,62,94,252,219,203,117,35,11,32,57,177,33,88,237,149,56,87,174,20,125,136,171,168,68,175,
            74,165,71,134,139,48,27,166,77,146,158,231,83,111,229,122,60,211,133,230,220,105,92,41,55,46,245,40,244,102,143,54,
            65,25,63,161,1,216,80,73,209,76,132,187,208,89,18,169,200,196,135,130,116,188,159,86,164,100,109,198,173,186,3,64,
            52,217,226,250,124,123,5,202,38,147,118,126,255,82,85,212,207,206,59,227,47,16,58,17,182,189,28,42,223,183,170,213,
            119,248,152,2,44,154,163,70,221,153,101,155,167,43,172,9,129,22,39,253,19,98,108,110,79,113,224,232,178,185,112,104,
            218,246,97,228,251,34,242,193,238,210,144,12,191,179,162,241,81,51,145,235,249,14,239,107,49,192,214,31,181,199,106,157,
            184,84,204,176,115,121,50,45,127,4,150,254,138,236,205,93,222,114,67,29,24,72,243,141,128,195,78,66,215,61,156,180
        };

        private static int FastFloor(double x)
        {
            return x > 0 ? (int)x : (int)x - 1;
        }

        public static double Noise(double x, double y, double z)
        {
            const double F3 = 1.0 / 3.0;
            const double G3 = 1.0 / 6.0;

            double n0, n1, n2, n3;

            double s = (x + y + z) * F3;
            int i = FastFloor(x + s);
            int j = FastFloor(y + s);
            int k = FastFloor(z + s);

            double t = (i + j + k) * G3;
            double X0 = i - t;
            double Y0 = j - t;
            double Z0 = k - t;
            double x0 = x - X0;
            double y0 = y - Y0;
            double z0 = z - Z0;

            int i1, j1, k1;
            int i2, j2, k2;

            if (x0 >= y0)
            {
                if (y0 >= z0) { i1 = 1; j1 = 0; k1 = 0; i2 = 1; j2 = 1; k2 = 0; }
                else if (x0 >= z0) { i1 = 1; j1 = 0; k1 = 0; i2 = 1; j2 = 0; k2 = 1; }
                else { i1 = 0; j1 = 0; k1 = 1; i2 = 1; j2 = 0; k2 = 1; }
            }
            else
            {
                if (y0 < z0) { i1 = 0; j1 = 0; k1 = 1; i2 = 0; j2 = 1; k2 = 1; }
                else if (x0 < z0) { i1 = 0; j1 = 1; k1 = 0; i2 = 0; j2 = 1; k2 = 1; }
                else { i1 = 0; j1 = 1; k1 = 0; i2 = 1; j2 = 1; k2 = 0; }
            }

            double x1 = x0 - i1 + G3;
            double y1 = y0 - j1 + G3;
            double z1 = z0 - k1 + G3;
            double x2 = x0 - i2 + 2.0 * G3;
            double y2 = y0 - j2 + 2.0 * G3;
            double z2 = z0 - k2 + 2.0 * G3;
            double x3 = x0 - 1.0 + 3.0 * G3;
            double y3 = y0 - 1.0 + 3.0 * G3;
            double z3 = z0 - 1.0 + 3.0 * G3;

            int ii = i & 255;
            int jj = j & 255;
            int kk = k & 255;

            double t0 = 0.6 - x0 * x0 - y0 * y0 - z0 * z0;
            if (t0 < 0.0) n0 = 0.0;
            else
            {
                t0 *= t0;
                n0 = t0 * t0 * Grad(Perm[(ii + Perm[(jj + Perm[kk & 255]) & 255]) & 255], x0, y0, z0);
            }

            double t1 = 0.6 - x1 * x1 - y1 * y1 - z1 * z1;
            if (t1 < 0.0) n1 = 0.0;
            else
            {
                t1 *= t1;
                n1 = t1 * t1 * Grad(Perm[(ii + i1 + Perm[(jj + j1 + Perm[(kk + k1) & 255]) & 255]) & 255], x1, y1, z1);
            }

            double t2 = 0.6 - x2 * x2 - y2 * y2 - z2 * z2;
            if (t2 < 0.0) n2 = 0.0;
            else
            {
                t2 *= t2;
                n2 = t2 * t2 * Grad(Perm[(ii + i2 + Perm[(jj + j2 + Perm[(kk + k2) & 255]) & 255]) & 255], x2, y2, z2);
            }

            double t3 = 0.6 - x3 * x3 - y3 * y3 - z3 * z3;
            if (t3 < 0.0) n3 = 0.0;
            else
            {
                t3 *= t3;
                n3 = t3 * t3 * Grad(Perm[(ii + 1 + Perm[(jj + 1 + Perm[(kk + 1) & 255]) & 255]) & 255], x3, y3, z3);
            }

            return 32.0 * (n0 + n1 + n2 + n3);
        }

        private static double Grad(int hash, double x, double y, double z)
        {
            int h = hash & 15;
            double u = h < 8 ? x : y;
            double v = h < 4 ? y : h == 12 || h == 14 ? x : z;
            return ((h & 1) == 0 ? u : -u) + ((h & 2) == 0 ? v : -v);
        }
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;
        Autodesk.Revit.DB.View activeView = doc.ActiveView;
        
        ICollection<ElementId> selectedElementIds = uiDoc.Selection.GetElementIds();
        if (selectedElementIds == null || selectedElementIds.Count == 0)
        {
            TaskDialog.Show("Error", "No elements are currently selected.");
            return Result.Cancelled;
        }

        double fractionToKeep = PromptForFraction();
        if (fractionToKeep < 0.0 || fractionToKeep > 1.0)
        {
            TaskDialog.Show("Error", "The entered value must be between 0.0 and 1.0.");
            return Result.Cancelled;
        }

        BoundingBoxXYZ totalBounds = null;
        foreach (ElementId elementId in selectedElementIds)
        {
            Element element = doc.GetElement(elementId);
            BoundingBoxXYZ elementBounds = element.get_BoundingBox(activeView);
            if (elementBounds != null)
            {
                if (totalBounds == null)
                {
                    totalBounds = elementBounds;
                }
                else
                {
                    totalBounds.Min = new XYZ(
                        Math.Min(totalBounds.Min.X, elementBounds.Min.X),
                        Math.Min(totalBounds.Min.Y, elementBounds.Min.Y),
                        Math.Min(totalBounds.Min.Z, elementBounds.Min.Z));
                    totalBounds.Max = new XYZ(
                        Math.Max(totalBounds.Max.X, elementBounds.Max.X),
                        Math.Max(totalBounds.Max.Y, elementBounds.Max.Y),
                        Math.Max(totalBounds.Max.Z, elementBounds.Max.Z));
                }
            }
        }

        if (totalBounds == null)
        {
            TaskDialog.Show("Error", "Could not calculate bounds of selected elements.");
            return Result.Cancelled;
        }

        var elementsWithNoise = new List<(ElementId, double)>();
        Random seedGenerator = new Random();
        double offsetX = seedGenerator.NextDouble() * 100.0;
        double offsetY = seedGenerator.NextDouble() * 100.0;
        double offsetZ = seedGenerator.NextDouble() * 100.0;

        double boxWidth = totalBounds.Max.X - totalBounds.Min.X;
        double boxHeight = totalBounds.Max.Y - totalBounds.Min.Y;
        double boxDepth = totalBounds.Max.Z - totalBounds.Min.Z;
        double maxDimension = Math.Max(Math.Max(boxWidth, boxHeight), boxDepth);

        double frequencyMultiplier = 50.0;

        foreach (ElementId elementId in selectedElementIds)
        {
            Element element = doc.GetElement(elementId);
            BoundingBoxXYZ bounds = element.get_BoundingBox(activeView);
            
            if (bounds != null)
            {
                double centerX = (bounds.Min.X + bounds.Max.X) * 0.5;
                double centerY = (bounds.Min.Y + bounds.Max.Y) * 0.5;
                double centerZ = (bounds.Min.Z + bounds.Max.Z) * 0.5;

                double normalizedX = (centerX - totalBounds.Min.X) / maxDimension;
                double normalizedY = (centerY - totalBounds.Min.Y) / maxDimension;
                double normalizedZ = (centerZ - totalBounds.Min.Z) / maxDimension;

                double noiseX = normalizedX * frequencyMultiplier + offsetX;
                double noiseY = normalizedY * frequencyMultiplier + offsetY;
                double noiseZ = normalizedZ * frequencyMultiplier + offsetZ;

                double noiseValue = (SimplexNoise.Noise(noiseX, noiseY, noiseZ) + 1) * 0.5;
                elementsWithNoise.Add((elementId, noiseValue));
            }
        }

        List<ElementId> elementsToKeep = elementsWithNoise
            .OrderByDescending(x => x.Item2)
            .Take((int)(fractionToKeep * selectedElementIds.Count))
            .Select(x => x.Item1)
            .ToList();

        uiDoc.Selection.SetElementIds(elementsToKeep);
        return Result.Succeeded;
    }

    private double PromptForFraction()
    {
        using (System.Windows.Forms.Form promptForm = new System.Windows.Forms.Form())
        {
            promptForm.Width = 300;
            promptForm.Height = 150;
            promptForm.Text = "Selection Parameters";
            promptForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            promptForm.MaximizeBox = false;
            promptForm.MinimizeBox = false;
            promptForm.StartPosition = FormStartPosition.CenterScreen;

            System.Windows.Forms.Label fractionLabel = new System.Windows.Forms.Label() { Left = 10, Top = 20, Text = "Fraction to keep (0.0 to 1.0):" };
            System.Windows.Forms.TextBox fractionBox = new System.Windows.Forms.TextBox() { Left = 10, Top = 40, Width = 260 };
            
            System.Windows.Forms.Button confirmation = new System.Windows.Forms.Button()
            {
                Text = "OK",
                Left = 190,
                Width = 80,
                Top = 80,
                DialogResult = DialogResult.OK
            };

            confirmation.Click += (sender, e) => { promptForm.Close(); };
            promptForm.Controls.AddRange(new System.Windows.Forms.Control[] { 
                fractionLabel, fractionBox, 
                confirmation 
            });
            promptForm.AcceptButton = confirmation;

            if (promptForm.ShowDialog() == DialogResult.OK)
            {
                if (double.TryParse(fractionBox.Text, out double fraction))
                {
                    return fraction;
                }
            }
            return -1;
        }
    }
}
