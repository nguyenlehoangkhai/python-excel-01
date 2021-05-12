import pandas as pd
import tkinter as tk
from tkinter import filedialog


# root = tk.Tk()
# root.withdraw()
#
# file_path = filedialog.askopenfilename()
file_path = 'C:\Khai\MyCode\ALM-tool-excel\output.xlsx'

excel_file = file_path

#### second way to read all sheet in a xlxs file
read_excel = pd.ExcelFile(excel_file)
all_sheets_temp = []
print(read_excel.sheet_names)


for sheet in read_excel.sheet_names:
    all_sheets_temp.append((read_excel.parse(sheet)))
    all_sheets = pd.concat(all_sheets_temp)
# print(all_sheets.shape)
# print(type(all_sheets))
# print(all_sheets)

'''
# print header of the dataframe
print(list(all_sheets.columns.values))
'''
# print(all_sheets.values[0, 0])

def locate(data, query, value, output):
    df = pd.DataFrame(data)
    # create a list of values in the query (column)
    values = df[query].tolist()
    # print(values)

    row = 0
    if value in values:
        print(value)
        row = values.index(value)
        print(row)

    return df.loc[row, output]


dict_all_sheets = dict(all_sheets)

print(type(dict_all_sheets))
# print(dict_all_sheets)

print(locate(dict_all_sheets, query= 'Title', value= 'Metropolis'+'\xa0', output= 'Year'))


# a = all_sheets.query('Title== "Over the Hill to the Poorhouse"')['IMDB Score']
# print(a)



'''
## print the last 10 items
# print(all_movies_02.tail(10))

# sort_by_gross = all_movies_02.sort_values(["Gross Earnings"], ascending= False)
# print(sort_by_gross["Gross Earnings"].head(10))
# print(all_movies_02.describe())
'''

# all_movies_02.to_excel('output.xlsx', index=False)

