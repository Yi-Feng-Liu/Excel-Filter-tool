import copy
import json
import pandas as pd

df = pd.read_excel('C:\\Users\\acer\\Desktop\\source.xlsx', sheet_name='工作表1', engine='openpyxl')
df.drop(labels=['歸屬廠區', '單位', '工作班別'], axis=1, inplace=True)

df.insert(0, '年度代碼', value='112體檢')
df.insert(9, 'I_Score', value=0)
df.insert(10, 'I_Risk', value=0)
df.insert(18, 'J_Score', value=0)
df.insert(19, 'J_Risk', value=0)
df.insert(22, '工作型態內容說明', value='無')
df.insert(23, '工作負荷等級', value=0)

target_col = ['年度代碼', '姓名', '員工編號', 'I01', 'I02', 'I03', 'I04', 'I05', 'I06', 'I_Score', 'I_Risk', 'J01', 'J02', 'J03', 'J04','J05', 'J06', 'J07', 'J_Score', 'J_Risk', '月加班時數等級', '工作型態評估等級', '工作型態內容說明', '工作負荷等級']
