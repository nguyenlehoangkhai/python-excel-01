# Import `pandas`
import pandas as pd

# Import `load_workbook` module from `openpyxl`
from openpyxl import load_workbook
from itertools import islice

# Import `dataframe_to_rows`
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import *


path = 'C:\Khai\MyCode\python-excel\movies.xls'
path02 = 'C:\Khai\MyCode\python-excel\ALM_xxxx_CodeReview_ShortDescription.xlsx'
excel_file = path

# #### first way to read all sheet in a xlsx file
# movies_01 = pd.read_excel(excel_file, index_col=0, sheet_name=0)
# movies_02 = pd.read_excel(excel_file, index_col=0, sheet_name=1)
# movies_03 = pd.read_excel(excel_file, index_col=0, sheet_name=2)
#
# movies_skip_rows = pd.read_excel(excel_file, header=None, skiprows=1)
# movies_skip_rows.head(5)
#
# all_movies_01 = pd.concat([movies_01, movies_02, movies_03])
# print(all_movies_01.shape)
# print(movies_02)

#### second way to read all sheet in a xlxs file
all_sheets = pd.ExcelFile(excel_file)
movies_sheets = []
print(all_sheets.sheet_names)

for sheet in all_sheets.sheet_names:
    movies_sheets.append((all_sheets.parse(sheet)))
    all_movies_02 = pd.concat(movies_sheets)
print(all_movies_02.shape)

## print the last 10 items
# print(all_movies_02.tail(10))

# sort_by_gross = all_movies_02.sort_values(["Gross Earnings"], ascending= False)
# print(sort_by_gross["Gross Earnings"].head(10))
print(all_movies_02.describe())


all_movies_02.to_excel('output.xlsx', index=False)

#
# writer = pd.ExcelWriter('OUTPUT.xls', engine='xlsxwriter')
#
# all_movies_02.to_excel(writer, index=False, sheet_name=" report")
# workbook= writer.bookworksheet = writer.sheets['report']