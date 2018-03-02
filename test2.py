
#-*-coding:utf-8 -*-
# from xlutils.copy import copy
#   西餐厅    ./1704/西餐厅.xls
import xlrd
import sys
import chardet
from xlutils.copy import copy

import sys
defaultencoding = 'utf-8'
if sys.getdefaultencoding() != defaultencoding:
    reload(sys)
    sys.setdefaultencoding(defaultencoding)




# key_sheet_name = '西餐厅'   #模板中对应的表格的名字
# key_path_target = './1704/西餐厅.xls' # 数据源xls 的path


key_sheet_name = sys.argv[1]   #模板中对应的表格的名字
key_path_target = sys.argv[2] # 数据源xls 的path



print('第二个参数是',key_sheet_name)
print('第三个参数是',key_path_target)


# step 1  解析模板 ================================================

book = xlrd.open_workbook('模版.xls',formatting_info=True)
sheets=book.sheets()

sheet_A37 = book.sheet_by_name(key_sheet_name)

rows = sheet_A37.nrows
cols = sheet_A37.ncols

print('row',rows)
print('col',rows)

# row 列
# col 行

list_cell = []

for row in range(rows):
	cell = sheet_A37.cell_value(row,3)  # 小弄堂   3 要改成2
	if isinstance(cell,float):
		print('不对',row,type(cell))
	if cell.strip() != '': # 需要判断shif
		cell_target  = {'row': row, 'col': 3, 'name': cell }  # 小弄堂   3 要改成2
		list_cell.append(cell_target)
		# print(cell.encode('utf-8'))

		
	    
# print(list_cell)


#  step2    anther sheet find value  ================================================



book_source = xlrd.open_workbook(key_path_target,formatting_info=True)

sheet_B = book_source.sheet_by_index(0)

# 拿到所有的  list

rows_sou = sheet_B.nrows
cols_sou = sheet_B.ncols

list_sou = []

for row in range(rows_sou):
	cell = sheet_B.cell_value(row,2)
	if cell.strip() != '': # 需要判断shif
		cell_target  = {'row': row, 'col': 1, 'name': cell ,'val1':sheet_B.cell_value(row,5), 'val2':sheet_B.cell_value(row,7)}

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


# print(list_all)




#  step4    写入数据  ================================================



rb = xlrd.open_workbook('模版.xls',formatting_info=True)
wb = copy(rb)
ws = wb.get_sheet(key_sheet_name)

print(ws.name)

for item in list_all:
	ws.write(item['row'], item['col']+2, item['val1'])
	ws.write(item['row'], item['col']+3, item['val2'])

wb.save('模版.xls')




# 注释：

# print(cell.__class__)  打印字符串 是 str 还是 unicode   http://in355hz.iteye.com/blog/1860787  http://wklken.me/posts/2013/08/31/python-extra-coding-intro.html

# 查看 文件的 编码格式
# f = open('YP03销售量排名.xls')
# data = f.read()
# print chardet.detect(data)


#  ws = wb.get_sheet(0) ws.write(0, 0, 'changed!')   get_sheet 通过 这样有 write的属性 通过sheet_by_index()获取的sheet没有write()方法

