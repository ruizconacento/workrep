from asyncore import read
import pandas as pd 
import numpy as np 

df = pd.read_csv('OT.05.10.2022.csv')
print(df.info())
print(df["Pto.tbjo.responsable"].mode())
df["Pto.tbjo.responsable"].fillna("Proveedor",inplace=True)

df_ = pd.read_csv('')