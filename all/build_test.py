
#-*-coding:utf-8 -*-
# from xlutils.copy import copy
#   西餐厅    ./1704/西餐厅.xls
import xlrd
import sys
import chardet
from xlutils.copy import copy

import os  
import json

import sys
defaultencoding = 'utf-8'
if sys.getdefaultencoding() != defaultencoding:
    reload(sys)
    sys.setdefaultencoding(defaultencoding)




# key_sheet_name = '西餐厅'   #模板中对应的表格的名字
# key_path_target = './1704/西餐厅.xls' # 数据源xls 的path


# one ========= 拿到配置json 文件

def get_json_file(filename):
    f = open(filename, encoding='utf-8')  #设置以utf-8解码模式读取文件，encoding参数必须设置，否则默认以gbk模式读取文件，当文件中包含中文时，会报错
    setting = json.load(f)
    return setting

map_key =  get_json_file("key.json")
list_map = map_key["content"]
print("配置文件是：",list_map)


# two ========= 根据配置 解析


# 获取需要改变的 数据
def get_list_target(key_sheet_name,col):
	book = xlrd.open_workbook('./source/模版.xls',formatting_info=True)
	sheets=book.sheets()
	sheet_A37 = book.sheet_by_name(key_sheet_name)
	rows = sheet_A37.nrows
	cols = sheet_A37.ncols
	list_cell = []

	for row in range(rows):
	    cell = sheet_A37.cell_value(row,col)  # 小弄堂   3 要改成2
	    if "小   类" == cell:
	    	continue
	    if isinstance(cell,float):
		    print('模板中的',key_sheet_name,row,type(cell))
	    if cell.strip() != '': # 需要判断shif
		    cell_target  = {'row': row, 'col': col, 'name': cell }  # 小弄堂   3 要改成2
		    list_cell.append(cell_target)
	return list_cell



# 获取数据源的数据
def get_list_source(key_path_target,col):
	print("path==",key_path_target)
	# col = 3  # 默认是 3  小弄堂 为2
	book_source = xlrd.open_workbook(key_path_target,formatting_info=True)
	sheet_B = book_source.sheet_by_index(0)
	rows_sou = sheet_B.nrows
	cols_sou = sheet_B.ncols
	list_sou = []

	# 检查 对应的列 
	cell_first = sheet_B.cell_value(0,col)
	if cell_first != "名称2":
		print("出错啦、、、、当前cell的value 1",cell_first,key_path_target,"对应不上")
		return

	cell_first = sheet_B.cell_value(0,col+3)
	print("出错啦、、、、当前cell的value 2",cell_first,key_path_target,"对应不上")
	if cell_first != "销量":
		return

	cell_first = sheet_B.cell_value(0,col+5)
	print("出错啦、、、、当前cell的value 3",cell_first,key_path_target,"对应不上")
	if cell_first != "销售额":
		return

	for row in range(rows_sou):
		cell = sheet_B.cell_value(row,col)
		print(cell,row,col)	
		if cell.strip() != '':
			cell_target  = {'row': row, 'col': col, 'name': cell ,'val1':sheet_B.cell_value(row,col+3), 'val2':sheet_B.cell_value(row,col+5)} #销量(往后移动3位)  + 销售额(往后移动5位)
			list_sou.append(cell_target)
	return list_sou
	


# 合并数据 + 把(source)颜色变成 红色

def merge_list(path,list_sou,list_cell):
	list_cell = []
	for row in range(rows):
		cell = sheet_A37.cell_value(row,2)  # 小弄堂   3 要改成2
		if isinstance(cell,float):
			print('模板中的',key_sheet_name,row,type(cell))
		if cell.strip() != '': 
			cell_target  = {'row': row, 'col': 2, 'name': cell }  # 小弄堂   3 要改成2
			list_cell.append(cell_target)
	return list_cell

    
	    


for item in list_map:
	list_cell_item = get_list_target(item["key_sheet_name"],item["col_target"])
	# print("最终数据",list_cell_item)
	lists_s = get_list_source(item["path"],item["col_source"])
	list_a = merge_list(item["path"],lists_s,list_cell_item)
	print("最终数据==",list_a)











# key_sheet_name = sys.argv[1]   #模板中对应的表格的名字
# key_path_target = sys.argv[2] # 数据源xls 的path



# print('第二个参数是',key_sheet_name)
# print('第三个参数是',key_path_target)


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

