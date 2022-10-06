from datetime import date
from datetime import datetime
import pandas as pd
import os

file = "OT " + datetime.now().strftime('%d.%m.%Y') + ".XLSX"
main_path = "C:\\Users\\mruiz\\Documents\\DB OT" 
current_month = datetime.now().strftime('%B')
current_path = main_path + "\\" + file

print(current_path)

main_path_ = os.listdir(main_path)

for item in main_path_:
    #print(item)
    if item == file:
        item_ = item
        print(item_)
        break
    
data_xls = pd.read_excel(current_path, index_col=0)
data = data_xls.to_csv(file + ".csv",encoding='utf-8', index=False)

data_f = pd.read_csv(data)







"""data_xls.to_csv(str(file), encoding='utf-8')
print(type(data_xls))"""