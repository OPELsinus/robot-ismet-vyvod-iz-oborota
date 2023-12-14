












import datetime

import urllib.parse


a = [1, 2, 3, 4, 5, 6]

start, end = 0, 2

while start < len(a):
    sublist = a[start:end]
    print(sublist)
    start = end
    end += 2