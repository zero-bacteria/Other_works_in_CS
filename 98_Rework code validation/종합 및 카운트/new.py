import os
import openpyxl as px
import pandas as pd
import datetime

creation_date = datetime.datetime.today()
creation_date = creation_date.strftime('%Y%m%d')[2:]

backdata_dir = './01_backdata/'

pre_data = pd.read_excel(backdata_dir + os.listdir('./01_backdata')[-2])
new_data = pd.read_excel(backdata_dir + os.listdir('./01_backdata')[-1])

print(os.listdir('./01_backdata')[-2])
print(os.listdir('./01_backdata')[-1])


