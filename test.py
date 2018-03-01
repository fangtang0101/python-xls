
#-*-coding:utf-8 -*-
# from xlutils.copy import copy
import xlrd
import sys
import chardet

import sys
defaultencoding = 'utf-8'
if sys.getdefaultencoding() != defaultencoding:
    reload(sys)
    sys.setdefaultencoding(defaultencoding)



# step 1  解析模板

book = xlrd.open_workbook('模版.xls',formatting_info=True)
sheets=book.sheets()

sheet_A37 = book.sheet_by_name('37#')

rows = sheet_A37.nrows
cols = sheet_A37.ncols

print('row',rows)
print('col',rows)

# row 列
# col 行

list_cell = []

for row in range(rows):
	cell = sheet_A37.cell_value(row,3)
	if cell.strip() != '': # 需要判断shif
		cell_target  = {'row': row, 'col': 3, 'name': cell }
		list_cell.append(cell_target)
		# print(cell.encode('utf-8'))

		
	    
# print(list_cell)


#  step2    anther sheet find value



book_source = xlrd.open_workbook('./1704/RY01销售量排名.xls',formatting_info=True)

sheet_B = book_source.sheet_by_index(0)

# 拿到所有的  list

rows_sou = sheet_B.nrows
cols_sou = sheet_B.ncols

list_sou = []

for row in range(rows_sou):
	cell = sheet_B.cell_value(row,1)
	if cell.strip() != '': # 需要判断shif
		cell_target  = {'row': row, 'col': 1, 'name': cell ,'val1':sheet_B.cell_value(row,3), 'val2':sheet_B.cell_value(row,4)}

		# val1 = sheet_B.cell_value(row,3)
		# val2 = sheet_B.cell_value(row,4)
		# cell_target.val1 = val1
		# cell_target.val2 = val2
		list_sou.append(cell_target)

# print(list_sou)	


# step 3  合并 数组

list_all = []

for item_sou in list_sou:
    for item_target in list_cell:
 	    if item_sou['name'] == item_target['name']:
 		    item_target['val1'] = item_sou['val1']
 		    item_target['val2'] = item_sou['val2']
 		    list_all.append(item_target)
 		    break


print(list_all)






# cell = sheet_B.cell_by_name("原味华夫")
# print(cell)




# 注释：

# print(cell.__class__)  打印字符串 是 str 还是 unicode   http://in355hz.iteye.com/blog/1860787  http://wklken.me/posts/2013/08/31/python-extra-coding-intro.html

# 查看 文件的 编码格式
# f = open('YP03销售量排名.xls')
# data = f.read()
# print chardet.detect(data)
