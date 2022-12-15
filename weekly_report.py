import pandas as pd 
import numpy as np 
import os 
import glob 

list = glob.glob('/home/laxmi/Documents/NIBL_Task/nibllss/Automation_data_collection/data/*')
print(list)

print(len(list))

df = pd.read_csv(list[0])
df1 = pd.read_csv(list[1])
df2 = pd.read_csv(list[2])

print(df.shape)
print(df1.shape)
print(df2.shape)

df_m = pd.concat([df, df1])
print(df_m.shape)

df_ma = pd.concat([df_m, df2])
print(df_ma.shape)

