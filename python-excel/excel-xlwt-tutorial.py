from xlwt import Workbook

wb = Workbook()
sheet1 = wb.add_sheet('sheet 1')
sheet2 = wb.add_sheet('sheet 2')
sheet3 = wb.add_sheet('sheet 3')

wb.save('demo1_xlwt.xls')


# Write data in sheet 1
sheet1.write(0, 0, 'STT')
sheet1.write(0, 1, 'Name')

# data
author_stdio = (
    [1, 'La Kiến Vinh'],
    [2, 'Vũ Quang Huy'],
    [3, 'Ryan Le'],
)

# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for item, cost in (author_stdio):
    sheet1.write(row, col, item)
    sheet1.write(row, col + 1, cost)
    row += 1

wb.save('demo1_xlwt.xls')