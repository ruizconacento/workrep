from datetime import datetime
from xml.etree.ElementTree import tostring
import pandas as pd
import os

#First, the vars was created
file = "OT " + datetime.now().strftime('%d.%m.%Y') + ".XLSX"
main_path = "C:\\Users\\mruiz\\Documents\\DB OT" 
current_month = datetime.now().strftime('%B')
main_path_ = os.listdir(main_path)
#print(main_path_)
#Then, we need to find the correct file
for item in main_path_:
    if item == file:
        item1 = item
        print(item1)
        break
for item in main_path_:
    if item == 'Master_OT.xlsx':
        item2 = item
        print(item2)
        break
#Then, just open it to start    
data1 = main_path + "\\" + item1
data2 = main_path + "\\" + item2
data1_ = pd.read_excel(data1, index_col=0, sheet_name="Sheet1", header=0)
data2_ = pd.read_excel(data2, index_col=0, sheet_name="Names", header=0)
print(data2_.to_string())