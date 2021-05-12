import pandas
from openpyxl import load_workbook

with pd.ExcelWriter(path, mode='a') as writer:
    s.to_Excel(writer, sheet_name='another sheet', index=False)