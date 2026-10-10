xs = [(i * 2654435761) % 1000003 for i in range(300000)]
xs.sort(reverse=True)
s = ",".join(str(x) for x in xs[:100000])
result = len(s)
