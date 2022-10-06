from datetime import date
from datetime import datetime
from xml.etree.ElementTree import tostring
import pandas as pd
import os

#First, the vars was created
file = "OT " + datetime.now().strftime('%d.%m.%Y') + ".XLSX"
main_path = "C:\\Users\\mruiz\\Documents\\DB OT" 
current_month = datetime.now().strftime('%B')
main_path_ = os.listdir(main_path)
#Then, we need to find the correct file
for item in main_path_:
    if item == file:
        item_ = item
        print(item_)
        break
#Then, just open it to start    
destination_path = main_path + "\\" + item_
data = pd.read_excel(destination_path, index_col=0, sheet_name="Sheet1", header=0)
print(data.info())
print(data)
