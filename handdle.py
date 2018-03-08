#-*-coding:utf-8 -*-
import os

import string
import shutil

import xlrd
import sys
import chardet
from xlutils.copy import copy
import json
import xlwt
import sys


#2. 将对应的数据放到 对应的 文档里面


path_or = "source_or"
path_tar = "source"
list_info ={
  "content": [
   {
    "path":"./source/model/西餐厅.xls",
    "key_sheet_name":"西餐厅",
     "col_source":2,
     "col_target":3
   },
   {
    "path":"./source/model/茶餐厅.xls",
    "key_sheet_name":"茶餐厅",
     "col_source":2,
     "col_target":3
   },
   {
    "path":"./source/model/D5.xls",
    "key_sheet_name":"D5",
     "col_source":2,
     "col_target":3
   },
   {
    "path":"./source/model/D4.xls",
    "key_sheet_name":"D4",
     "col_source":2,
     "col_target":3
   },
   {
    "path":"./source/model/小弄堂.xls",
    "key_sheet_name":"小弄堂",
     "col_source":2,
     "col_target":2
   },
   {
    "path":"./source/model/A2.xls",
    "key_sheet_name":"A2",
     "col_source":2,
     "col_target":3
   },
   {
    "path":"./source/model/37#.xls",
    "key_sheet_name":"37#",
     "col_source":2,
     "col_target":3
   },
   {
    "path":"./source/model/A3.xls",
    "key_sheet_name":"A3",
     "col_source":2,
     "col_target":3
   }
  ]
}



   


def get_string_split(str_):
	str_ = ''.join(str_.split())
	return str_



def copy_file(path_form,path_to):
	#if target path is exist then delete it
    path_to = path_to.strip()
    path_to = path_to.rstrip("\\")
    isExists = os.path.exists(path_to)
    print("file exist",isExists)

    if isExists: 
    	shutil.rmtree(path_to) #delete file
    	print("success delete file name is ",path_to)

    shutil.copytree("source_or", "source")
    print("success copy file ...",)



def fix_data():
	pass
# get data  in  corresponding xls 
def get_data_corresponding(item):
	key_sheet_name = item["key_sheet_name"]
	list_temp = []
	path = "对应数据_名称.xls"
	book = xlrd.open_workbook(path,formatting_info=True)
	sheets=book.sheets()
	sheet_item = book.sheet_by_name(key_sheet_name)

	rows = sheet_item.nrows
	cols = sheet_item.ncols
	list_cell = []

	col_target = 2

	for row in range(1,rows):
		cell_temp = sheet_item.cell_value(row,col_target)  
		cell_data = sheet_item.cell_value(row,col_target+1)  
		if (cell_temp.strip() != '' and cell_data.strip() != '') :
			cell_target  = {'cell_temp': cell_temp, 'cell_data': cell_data} 
			list_cell.append(cell_target)
	# print("list_cell....",list_cell)
	return list_cell


def filled_data(item,list_all):
	target_path_model = item["path"]
	key_sheet_name = item["key_sheet_name"]
	rb = xlrd.open_workbook(target_path_model,formatting_info=True)
	wb = copy(rb)
	ws = wb.get_sheet(0)

	sheet_readonly = rb.sheet_by_index(0)

    # note...... must use sheet_by_index
	rows = sheet_readonly.nrows
	cols = sheet_readonly.ncols
	col_target = 2

	# 设置单元格颜色
	pattern = xlwt.Pattern() # Create the Pattern
	pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
	pattern.pattern_fore_colour = 2 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
	style = xlwt.XFStyle() # Create the Pattern
	style.pattern = pattern # Add Pattern to Style

	for data in list_all:
		for row in range(1,rows):
			cell = sheet_readonly.cell_value(row,col_target) 
			if get_string_split(cell) == get_string_split(data["cell_data"]):
				data["row"] = row
				data["col"] = col_target
				break
    
	for obj in list_all:
		if 'row' in obj :
			# ws.write(obj['row'], obj['col'], obj['cell_temp'],style)
			ws.write(obj['row'], obj['col'], obj['cell_temp'],style)
		else:
			print(item["key_sheet_name"],"connt find ... ",obj)
	wb.save(target_path_model)

def main_method():
  copy_file(path_or,path_tar)

  for item in list_info["content"]:
    list_data_item = get_data_corresponding(item)
    filled_data(item,list_data_item)
    print("finished ...",item["key_sheet_name"])


main_method()

    



# step 1 copy file
# copy_file(path_or,path_tar)

# step 2 fix data

# for item in list_info["content"]:
# 	 list_data_item = get_data_corresponding(item)
# 	 filled_data(item,list_data_item)
# 	 print("finished ...",item["key_sheet_name"])







# a=' hello world '
# a = ''.join(a.split())
# print(a)










