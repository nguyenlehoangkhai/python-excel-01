# Import `xlrd`
import xlrd

# Open a workbook
workbook = xlrd.open_workbook('demo.xlsx')

# Loads only current sheets to memory
workbook = xlrd.open_workbook('demo.xlsx', on_demand = True)