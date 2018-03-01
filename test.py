from xlutils.copy import copy
import xlrd
import sys



# book = xlrd.open_workbook('模版.xls',formatting_info=True)
# sheets=book.sheets()

# sheet_A37 = book.sheet_by_name('37#')

# rows = sheet_A37.nrows
# cols = sheet_A37.ncols

# print('row',rows)
# print('col',rows)

# # row 列
# # col 行

# list_cell = []

# for row in range(rows):
# 	cell = sheet_A37.cell_value(row,3)
# 	if cell.strip() != '':
# 		cell_target  = {'row': row, 'col': 3, 'name': cell }
# 		list_cell.append(cell_target)

	    
# print(list_cell)


# anther sheet find value

book_source = xlrd.open_workbook('YP03销售量排名.xls',formatting_info=True,encoding_override=sys.getfilesystemencoding())

# sheet_B = book_source.sheet_by_index(0)
# print(sheet_B.name)



