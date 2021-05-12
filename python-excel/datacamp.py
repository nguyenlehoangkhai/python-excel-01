import os
import pandas as pd


# Retrieve current working directory
cwd = os.getcwd()
# print(cwd)

# List all files and directories in current directory
os.listdir(cwd)
# print(os.listdir(cwd))

# Assign spreadsheet filename to 'file'
file_01 = 'example01.xlsx'
file_02 = 'example02.xlsx'

# Load spreadsheet
load_excel_01 = pd.ExcelFile(file_02)
load_excel_02 = pd.ExcelFile(file_02)


# Print the sheet names
print(load_excel_01.sheet_names)
print(load_excel_02.sheet_names)

# Load a sheet into a DataFrame by name or index:
sheet_01_01 = load_excel_01.parse('Sheet1')
# print(sheet_01)
sheet_01_02 = load_excel_02.parse('Sheet1')

# Specify a writer
writer = pd.ExcelWriter('example03.xlsx')

# Write your dataframe to a file
# YourData is a dataFrame that you are interesting in writing as an excel file
yourData = pd.DataFrame(sheet_01_02)
yourData.to_excel(writer, 'Sheet2', index=False)

# Save the result
writer.save()
