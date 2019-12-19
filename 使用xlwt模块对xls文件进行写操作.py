import xlwt  #2003以前的表格以.xls后缀，用xlwt来写表格,用xlrd来读取表格,搭配xlutils修改表格。
             #2007的表格以.xlsx后缀，用openpyxl来读写表格。

#创建工作簿
#创建一个Workbook对象，相当于创建了一个Excel文件
book = xlwt.Workbook(encoding="utf-8",style_compression=0)

'''
Workbook类初始化时有encoding和style_compression参数
encoding:设置字符编码，一般要这样设置：w = Workbook(encoding='utf-8')，就可以在excel中输出中文了。默认是ascii。
style_compression:表示是否压缩，不常用。
'''

#创建工作表
# 创建一个sheet对象，一个sheet对象对应Excel文件中的一张表格。
sheet = book.add_sheet('test01',cell_overwrite_ok=True)
# 其中的test01是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False

#按单元格的方式向工作表中添加数据
sheet.write(0,0,'各省市')   # 其中的'0-行, 0-列'指定表中的单元，'各省市'是向该单元写入的内容
sheet.write(0,1,'薪资收入')


book.save('C:\\Users\\EDZ\\Desktop\\zhanglei.xls')