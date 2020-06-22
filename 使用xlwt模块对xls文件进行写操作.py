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
sheet1 = book.add_sheet('test01',cell_overwrite_ok=True)  #创建的第一张sheet1工作表为test01,隶属于同一个工作薄book(同一个Excel文件)
sheet2 = book.add_sheet('test02',cell_overwrite_ok=True)  #创建的第一张sheet2工作表为test02,隶属于同一个工作薄book(同一个Excel文件)
# 其中的test01是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False

#按单元格的方式向工作表中添加数据
sheet1.write(0,0,'各省市')   # 其中的'0-行, 0-列'指定表中的单元，'各省市'是向该单元写入的内容
sheet1.write(0,1,'薪资收入')

#也可以这样添加数据
txt1 = '杭州市'
txt2 = '5047.41'
sheet1.write(1,0,txt1)  #向工作表sheet的第二行，第一列插入txt1数据
sheet1.write(1,1,txt2)  #向工作表sheet的第二行，第二列插入txt2数据

#按行或列方式向工作表中添加数据(数据添加到了第二张sheet表中“test02”)
 #(第17行)为了验证这个功能，我们在工作簿中，再创建一个工作表，上个工作表叫“test01”，那么这个工作表命名为“test02”，都隶属于同一个工作簿。在下面代码中test02是表名，sheet2才是可供操作的工作表对象。
Province = ['北京市', '天津市', '河北省', '山西省', '内蒙古自治区', '辽宁省',
			'吉林省', '黑龙江省', '上海市', '江苏省', '浙江省', '安徽省', '福建省',
			'江西省', '山东省', '河南省', '湖北省', '湖南省', '广东省', '广西壮族自治区',
			'海南省', '重庆市', '四川省', '贵州省', '云南省', '西藏自治区', '陕西省', '甘肃省',
			'青海省', '宁夏回族自治区', '新疆维吾尔自治区']

Income = ['5047.4', '3247.9', '1514.7', '1374.3', '590.7', '1499.5', '605.1', '654.9',
		  '6686.0', '3104.8', '3575.1', '1184.1', '1855.5', '1441.3', '1671.5', '1022.7',
		  '1199.2', '1449.6', '2906.2', '972.3', '555.7', '1309.9', '1219.5', '715.5', '441.8',
		  '568.4', '848.3', '637.4', '653.3', '823.1', '254.1']
Project = ['各省市', '工资性收入', '家庭经营纯收入', '财产性收入', '转移性收入', '食品', '衣着',
		   '居住', '家庭设备及服务', '交通和通讯', '文教、娱乐用品及服务', '医疗保健', '其他商品及服务']

#填入第一列
for i in range(0,len(Province)):
	sheet2.write(i+1,0,Province[i])

#填入第二列
for i in range(0,len(Income)):
	sheet2.write(i+1,1,Income[i])

#填入第一行
for i in range(0,len(Project)):
	sheet2.write(0,i,Project[i])


# 最后，将以上操作保存到指定的Excel文件中(先创建一个Excel文件。)
book.save('C:\\Users\\EDZ\\Desktop\\123.xls')