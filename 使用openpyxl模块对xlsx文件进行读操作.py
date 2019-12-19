#使用openpyxl模块对xlsx文件进行读操作

'''
xlrd和xlwt都是针对Excel97-2003操作的，也就是以xls结尾的文件。
Excel2007以上的版本，以xlsx为后缀。要对这种类型的Excel文件进行操作要使用openpyxl，该模块既可以进行“读”操作，也可以进行“写”操作，还可以对已经存在的文件做修改。
'''

import openpyxl

#获取工作簿对象
workbook = openpyxl.load_workbook("C:\\Users\\EDZ\\Desktop\\1.xlsx")
#与xlrd 模块的区别
#wokrbook=xlrd.open_workbook("C:\\Users\\EDZ\\Desktop\\1.xlsx")

#%%
#获取所有工作表名
#获取工作簿 workbook的所有工作表
shenames = workbook.get_sheet_names()
#在xlrd模块中为 sheetnames=workbook.sheet_names()
print(shenames)
#使用上述语(shenames = workbook.get_sheet_names())句会发出警告：DeprecationWarning: Call to deprecated function get_sheet_names (Use wb.sheetnames).
#说明 get_sheet_names已经被弃用 可以改用 wb.sheetnames 方法(25行)

#%%
shenames = workbook.sheetnames
print(shenames)