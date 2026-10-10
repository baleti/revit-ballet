var o = new XYZ(0, 0, 0);
double s = 0;
for (int i = 0; i < 50000; i++) s += new XYZ(i, i % 7, 1).DistanceTo(o);
return Math.Round(s);
