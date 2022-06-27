import os
import pandas as pd
import openpyxl as px

lf = os.listdir('./aset')
rf = os.listdir('./bset')

lb = px.load_workbook(f'./aset/{lf[0]}')
ls = lb.active

ls.delete_cols(1,3)


pd.
# rb = px.load_workbook(f'./aset/{rf[0]}')
# rs = rb.active