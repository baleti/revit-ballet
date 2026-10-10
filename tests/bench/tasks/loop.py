try:
    xrange
except NameError:
    xrange = range
s = 0
for i in xrange(3000000):
    s += i * i % 7
result = s
