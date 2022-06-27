from multiprocessing.dummy import active_children
import os
import openpyxl as px
import pandas as pd

xlist = list()

for f in os.listdir('./aset'):
    if 'xlsx' in f:
        xlist.append(f)


tb = px.load_workbook(f'./aset/{xlist[0]}')
ts = tb.active





