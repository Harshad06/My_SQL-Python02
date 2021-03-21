import xlrd

book = xlrd.open_workbook('wb.xlsx')
print(book.nsheets)

print(book.sheet_names())

first_sheet = book.sheet_by_index(0)
print(first_sheet.row_values(0))

cell = first_sheet.cell(0,0)
print(cell)

print(cell.value)
