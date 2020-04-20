import xlwt

book = xlwt.Workbook(encoding="utf-8", style_compression=0)

sheet1 = book.add_sheet('test01', cell_overwrite_ok=True)

list1 = list(range(1,1001))
print(list1)
list2 = list1*3
list2.sort()
print(list2)

target_list = []
for i in list2:
	target_list.append('hoy%s' %(i))
print(target_list)

Province = list(target_list)

for i in range(0, len(Province)):
	sheet1.write(i + 1, 0, Province[i])

book.save('C:\\Users\\EDZ\\Desktop\\123.xls')