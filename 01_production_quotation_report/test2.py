a = ['~123.xlsx', 'a','b','~4.xlsx','~5.xlsx']



for i in range(0, len(a)):
    if '~' in a[i]:
        a.remove(a[i])


print(a)