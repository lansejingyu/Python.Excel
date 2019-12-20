#使用xlrd模块对xls文件进行读操作
#2003以前的表格以.xls后缀，用xlwt来写表格,用xlrd来读取表格,搭配xlutils修改表格。

#1.1 获取工作簿对象
#%%
# coding=gbk
import xlrd  #引入模块

#打开文件，获取excel文件的workbook（工作簿）对象
workbook=xlrd.open_workbook("D:\\Python.Excel处理\\1.xlsx")  #文件路径

#1.2 获取工作表对象
'''对工作簿”(workbook)workbook对象进行操作'''
names = workbook.sheet_names()    #获取所有sheet的名字
print(names)

worksheet = workbook.sheet_by_index(0)      #通过sheet索引获得sheet对象
print(worksheet)

worksheet = workbook.sheet_by_name("Sheet1")    #通过sheet名获得sheet对象
print(worksheet)

sheet0_name = workbook.sheet_names()[0]     #由上可知，workbook.sheet_names() 返回一个list对象，可以对这个list对象进行操作
print(sheet0_name)                          #通过sheet索引获取sheet名称

#%%
#1.3 获取工作表的基本信息
'''对工作表（sheet）对象进行操作'''
nrows = worksheet.nrows       #获取该表总行数
print(nrows)

ncols = worksheet.ncols       #获取该表总列数
print(ncols)

#%%
#1.4 按行或列方式获得工作表的数据

for i in range(nrows):        #循环打印每一行
	print(worksheet.row_values(i))
col_data = worksheet.col_values(0)
print(col_data)               #获取第一列的内容

#%%
#1.5 获取某一个单元格的数据(在xlrd模块中，工作表的行和列都是从0开始计数的。)
#通过坐标读取表格中的数据
cell_value1=worksheet.cell_value(0,0)
cell_value2=worksheet.cell_value(1,1)
print(cell_value1)
print(cell_value2)

#%%
cell_value1 =worksheet.cell(0,0).value
print(cell_value1)
cell_value2 =worksheet.row(0)[0].value
print(cell_value2)

