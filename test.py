
#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd

data = xlrd.open_workbook('4月分析表.xls')

tables = data.sheets()

# for tab in tables :
#     print(tab.name)


# 获取第一个表格 A3
table1 = data.sheet_by_name(u'A3')
print(table1.name)

#获取整行和整列的值（数组）  第几行 的所有数据
rows = table1.row_values(1)

nrows = table1.nrows
print(nrows)

for row in rows :
    print(row)

# 单元格
cell_A1 = table1.cell(3,3).value
print(cell_A1)

table1.cell(3,3).value = 'haha'


ctype = 1 
value = '测试的值'
xf = 0 # 扩展的格式化
table1.put_cell(3, 3, ctype, value, xf)
print(cell_A1)




