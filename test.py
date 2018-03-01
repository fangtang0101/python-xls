from xlutils.copy import copy
import xlrd



book = xlrd.open_workbook('模版.xls',formatting_info=True)
sheets=book.sheets()

sheet_A37 = book.sheet_by_name('37#')

rows = sheet_A37.nrows
cols = sheet_A37.ncols

print('row',rows)
print('col',rows)

for row in range(rows-100):
	cell = sheet_A37.cell_value(row,3)
	print('cell value 哈哈:',cell)