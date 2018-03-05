
#-*-coding:utf-8 -*-
# from xlutils.copy import copy
#   西餐厅    ./1704/西餐厅.xls
import xlrd
import sys
import chardet
from xlutils.copy import copy

import os  
import json
import xlwt

import sys
defaultencoding = 'utf-8'
if sys.getdefaultencoding() != defaultencoding:
	reload(sys)
	sys.setdefaultencoding(defaultencoding)

# key_sheet_name = '西餐厅'   #模板中对应的表格的名字
# key_path_target = './1704/西餐厅.xls' # 数据源xls 的path


# one ========= 拿到配置json 文件
list_map = []
def get_json_file(filename):
    f = open(filename, encoding='utf-8')  #设置以utf-8解码模式读取文件，encoding参数必须设置，否则默认以gbk模式读取文件，当文件中包含中文时，会报错
    setting = json.load(f)
    return setting

    



def body_func(map_content):
	key_sheet_name = map_content["key_sheet_name"]   #模板中对应的表格的名字
	key_path_target = map_content["path"]   # 数据源xls 的path
	print('模板中的sheet名字',key_sheet_name)
	print('数据源路劲表格',key_path_target)


	col_target = map_content["col_target"]  # 小弄堂   3 要改成2
	col_source = map_content["col_source"]  # 可能会改动的
	key_path_target = map_content["path"] 
	

	# step 1  解析模板 ================================================  start  
	book = xlrd.open_workbook('./source/模版.xls',formatting_info=True)
	sheets=book.sheets()
	sheet_A37 = book.sheet_by_name(key_sheet_name)
	rows = sheet_A37.nrows
	cols = sheet_A37.ncols
	list_cell = []
	for row in range(rows):
		cell = sheet_A37.cell_value(row,col_target)  
		if isinstance(cell,float):
			print('模板中的',key_sheet_name,row,type(cell))
		if cell.strip() != '': # 需要判断shif
			cell_target  = {'row': row, 'col': col_target, 'name': cell }  # 小弄堂   3 要改成2
			list_cell.append(cell_target)
	# print("模板中的数据",list_cell)	


	# step 2  数据源去数据 ================================================  start  
	book_source = xlrd.open_workbook(key_path_target,formatting_info=True)
	sheet_B = book_source.sheet_by_index(0)
	rows_sou = sheet_B.nrows
	cols_sou = sheet_B.ncols
	list_sou = []
	for row in range(rows_sou):
		cell = sheet_B.cell_value(row,col_source)
		if cell.strip() != '': 
			cell_target  = {'row': row, 'col': col_source, 'name': cell ,'val1':sheet_B.cell_value(row,col_source+3), 'val2':sheet_B.cell_value(row,col_source+5)} #销量(往后移动3位)  + 销售额(往后移动5位)
			list_sou.append(cell_target)
	# print("数据源中的数据",list_sou)	


	# step 3  对比数据 筛选数据 ================================================  start  
	list_all = []
	for item_sou in list_sou:
		for item_target in list_cell:
			if item_sou['name'] == item_target['name']:
				item_target['val1'] = item_sou['val1']
				item_target['val2'] = item_sou['val2']
				item_target["row_source"] = item_sou['row']
				item_target["col_source"] = item_sou['col']
				list_all.append(item_target)
				break
	# print("筛选对比之后的数据",list_all)

	# step 4  填写数据 保存表格 ================================================  start  
	rb = xlrd.open_workbook('./source/模版.xls',formatting_info=True)
	wb = copy(rb)
	ws = wb.get_sheet(key_sheet_name)
	for item in list_all:
		ws.write(item['row'], item['col']+2, item['val1'])
		ws.write(item['row'], item['col']+3, item['val2'])
	wb.save('模版.xls')


	# 重写 souce 的单元格

	rb = xlrd.open_workbook(key_path_target,formatting_info=True)
	wb = copy(rb)
	ws = wb.get_sheet(0)

    # 设置单元格颜色
	pattern = xlwt.Pattern() # Create the Pattern
	pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
	pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
	style = xlwt.XFStyle() # Create the Pattern
	style.pattern = pattern # Add Pattern to Style

	for item in list_all:
		# cell = ws.cell_value(item['row_source'],item['col_source'])
		ws.write(item['row_source'], item['col_source']+3,item['val1'],style)
		ws.write(item['row_source'], item['col_source']+5,item['val2'],style)
    save_path = './source/model/'+key_sheet_name+'.xls'
	wb.save(save_path)














map_key =  get_json_file("key.json")
list_map = map_key["content"]
print("配置文件是：",list_map)
for item in list_map:
	body_func(item)


# # step 1  解析模板 ================================================

# book = xlrd.open_workbook('模版.xls',formatting_info=True)
# sheets=book.sheets()

# sheet_A37 = book.sheet_by_name(key_sheet_name)

# rows = sheet_A37.nrows
# cols = sheet_A37.ncols

# print('row',rows)
# print('col',rows)

# # row 列
# # col 行

# list_cell = []

# for row in range(rows):
# 	cell = sheet_A37.cell_value(row,2)  # 小弄堂   3 要改成2
# 	if isinstance(cell,float):
# 		print('模板中的',key_sheet_name,row,type(cell))
# 	if cell.strip() != '': # 需要判断shif
# 		cell_target  = {'row': row, 'col': 2, 'name': cell }  # 小弄堂   3 要改成2
# 		list_cell.append(cell_target)
# 		# print(cell.encode('utf-8'))



# # print(list_cell)


# #  step2    anther sheet find value  ================================================



# book_source = xlrd.open_workbook(key_path_target,formatting_info=True)

# sheet_B = book_source.sheet_by_index(0)

# # 拿到所有的  list

# rows_sou = sheet_B.nrows
# cols_sou = sheet_B.ncols

# list_sou = []

# for row in range(rows_sou):
# 	cell = sheet_B.cell_value(row,2)
# 	if cell.strip() != '': # 需要判断shif
# 		cell_target  = {'row': row, 'col': 1, 'name': cell ,'val1':sheet_B.cell_value(row,5), 'val2':sheet_B.cell_value(row,7)}

# 		# val1 = sheet_B.cell_value(row,3)
# 		# val2 = sheet_B.cell_value(row,4)
# 		# cell_target.val1 = val1
# 		# cell_target.val2 = val2
# 		list_sou.append(cell_target)

# # print(list_sou)	


# # step 3  合并 数组

# list_all = []

# for item_sou in list_sou:
#     for item_target in list_cell:
#  	    if item_sou['name'] == item_target['name']:
#  		    item_target['val1'] = item_sou['val1']
#  		    item_target['val2'] = item_sou['val2']
#  		    list_all.append(item_target)
#  		    break


# # print(list_all)




# #  step4    写入数据  ================================================



# rb = xlrd.open_workbook('模版.xls',formatting_info=True)
# wb = copy(rb)
# ws = wb.get_sheet(key_sheet_name)

# print(ws.name)

# for item in list_all:
# 	ws.write(item['row'], item['col']+2, item['val1'])
# 	ws.write(item['row'], item['col']+3, item['val2'])

# wb.save('模版.xls')




# 注释：

# print(cell.__class__)  打印字符串 是 str 还是 unicode   http://in355hz.iteye.com/blog/1860787  http://wklken.me/posts/2013/08/31/python-extra-coding-intro.html

# 查看 文件的 编码格式
# f = open('YP03销售量排名.xls')
# data = f.read()
# print chardet.detect(data)


#  ws = wb.get_sheet(0) ws.write(0, 0, 'changed!')   get_sheet 通过 这样有 write的属性 通过sheet_by_index()获取的sheet没有write()方法

