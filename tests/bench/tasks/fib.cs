Func<int, int> fib = null;
fib = n => n < 2 ? n : fib(n - 1) + fib(n - 2);
return fib(27);
