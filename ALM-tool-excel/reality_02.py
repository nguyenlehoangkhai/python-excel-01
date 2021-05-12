import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

############################################################################

'''
open directory or file with a window dialog using tkinter library 
'''
root = tk.Tk()
root.withdraw()
# file_path = filedialog.askopenfilename()
file_dir = filedialog.askdirectory()

dir = file_dir
# dir = 'C:/Khai/MyCode/Excel-demo-files'
print(os.listdir(dir), '\n')
###########################################################################
'''
main function
'''

i = 1

excel_output = {'Number': [], 'Project name': [], 'Result': [],
                'No Findings': [], 'Author': [], 'Reviewer': []}
print(excel_output)

for filename in os.listdir(dir):
    list_file = pd.ExcelFile(os.path.join(dir, filename))

    # file_path = 'C:/Khai/MyCode/ALM-tool-excel/ALM-sample.xlsx'
    file_path = list_file

    excel_file = file_path
    parse_excel = pd.ExcelFile(excel_file)

    ##############################################################################

    # parse one sheet of a excel file
    sheet_01 = pd.read_excel(parse_excel, sheet_name=1)
    # print(sheet_01)
    sheet_02 = pd.read_excel(parse_excel, sheet_name=2)
    # print(sheet_02)
    sheet_03 = pd.read_excel(parse_excel, sheet_name=3)
    # print(sheet_03)

    # print(sheet_01.shape)
    # print(type(sheet_01))
    # print(sheet_01)

    #####################################################################

    all_sheets_temp = []
    # print(parse_excel.sheet_names)

    # second way to read all sheet in a xlxs file
    for sheet in parse_excel.sheet_names:
        if (sheet == 'History'): continue
        all_sheets_temp.append((parse_excel.parse(sheet)))
        all_sheets = pd.concat(all_sheets_temp)

    #####################################################################

    # # print header of the dataframe
    # print(list(sheet_01.columns.values))
    # print(sheet_01.values[0, 1])
    # # print the last 10 items
    # print(all_movies_02.tail(10))

    ######################################################################
    '''
    this is a function that search and query data with 
    a coordinate in data frame 
    
    '''
    def locate(data, query, value, output):
        df = pd.DataFrame(data)
        # create a list of values in the query (column)
        values = df[query].tolist()
        # print(values)
        row = 0
        if value in values:
            # print(value)
            row = values.index(value)
            # print(row)

        return df.loc[row, output]


    '''ham rut gon '''
    # data = pd.DataFrame(sheet_01)
    # a = (data['General'].tolist()).index('Project / Component')
    # output = data.loc[a, 'Unnamed: 1']
    # print(output)
    #######################################################################
    '''
    write data to dictionary, and this dict is used to convert to xml or xlsx
    '''

    name = locate(sheet_01, 'General', 'Project / Component', 'Unnamed: 1')
    result = locate(sheet_01, 'General', 'Formally ok (automatic!)', 'Unnamed: 1')
    noFindings = locate(sheet_01, 'General', 'Number of findings', 'Unnamed: 2')
    author = locate(sheet_01, 'Unnamed: 2', 'Author', 'Unnamed: 1')
    reviewer = locate(sheet_01, 'Unnamed: 2', 'Reviewer', 'Unnamed: 1')

    excel_output['Number'].append(i)
    excel_output['Result'].append(result)
    excel_output['Project name'].append(name)
    excel_output['Author'].append(author)
    excel_output['No Findings'].append(noFindings)
    excel_output['Reviewer'].append(reviewer)

    i = i + 1
    #######################################################################

#######################################################################
'''
write data to Output
'''
print(type(excel_output))
print(excel_output)
df_excel_output = pd.DataFrame(excel_output)
df_excel_output.to_excel('excel_output_01.xlsx', index=False)
