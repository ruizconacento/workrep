from xml.etree.ElementTree import tostring
from datetime import datetime
import pandas as pd
import os

##First, the vars was created
file = "OT " + datetime.now().strftime('%d.%m.%Y') + ".XLSX"
main_path = "C:\\Users\\mruiz\\Documents\\DB OT" 
current_month = datetime.now().strftime('%B')
main_path_ = os.listdir(main_path)
#print(main_path_)
##Then, we need to find the correct file
for item in main_path_:
    if item == file:
        item1 = item
        print("Los datos del día de hoy son : ", item1)
        break
for item in main_path_:
    if item == 'Master_OT.xlsx':
        item2 = item
        print("La base de datos de nombres es : ", item2)
        break
##Then, just open it to start    
data1 = main_path + "\\" + item1
data1excel = pd.read_excel(data1, index_col=0, sheet_name="Sheet1", header=0)
data1excel_ = pd.DataFrame(data1excel)
path_csv = main_path + "\\out.csv"
data1csv = data1excel_.to_csv(path_csv)
data1csv_ = pd.read_csv(path_csv)
print(data1csv_)

data2 = main_path + "\\" + item2
data2excel = pd.read_excel(data2, index_col=0, sheet_name="Names", header=0)
data2excel_ = pd.DataFrame(data2excel)
path_csv = main_path + "\\" + str(item2)
data2csv = data2excel_.to_csv(path_csv)
data2csv_ = pd.read_csv(path_csv)
print(data2csv_)
