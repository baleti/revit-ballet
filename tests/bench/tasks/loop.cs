long s = 0;
for (int i = 0; i < 3000000; i++) s += (long)i * i % 7;
return s;
