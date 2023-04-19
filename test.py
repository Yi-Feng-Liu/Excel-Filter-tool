import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.styles import Alignment, numbers
import copy

df = pd.read_excel('111年度健檢代謝症候群.xlsx', sheet_name='工作表1', engine='openpyxl')

print(df.iloc[0,:].tolist())
print(df.columns[0])