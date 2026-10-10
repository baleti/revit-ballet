var xs = new List<int>(300000);
for (int i = 0; i < 300000; i++) xs.Add((int)((i * 2654435761L) % 1000003));
xs.Sort((a, b) => b.CompareTo(a));
var s = string.Join(",", xs.Take(100000));
return s.Length;
