#ʹ��xlrdģ���xls�ļ����ж�����
#2003��ǰ�ı����.xls��׺����xlwt��д���,��xlrd����ȡ���,����xlutils�޸ı��

#1.1 ��ȡ����������
#%%
# coding=gbk
import xlrd  #����ģ��

#���ļ�����ȡexcel�ļ���workbook��������������
workbook=xlrd.open_workbook("D:\\Python.Excel����\\1.xlsx")  #�ļ�·��

#1.2 ��ȡ���������
'''�Թ�������(workbook)workbook������в���'''
names = workbook.sheet_names()    #��ȡ����sheet������
print(names)

worksheet = workbook.sheet_by_index(0)      #ͨ��sheet�������sheet����
print(worksheet)

worksheet = workbook.sheet_by_name("Sheet1")    #ͨ��sheet�����sheet����
print(worksheet)

sheet0_name = workbook.sheet_names()[0]     #���Ͽ�֪��workbook.sheet_names() ����һ��list���󣬿��Զ����list������в���
print(sheet0_name)                          #ͨ��sheet������ȡsheet����

#%%
#1.3 ��ȡ������Ļ�����Ϣ
'''�Թ�����sheet��������в���'''
nrows = worksheet.nrows       #��ȡ�ñ�������
print(nrows)

ncols = worksheet.ncols       #��ȡ�ñ�������
print(ncols)

#%%
#1.4 ���л��з�ʽ��ù����������

for i in range(nrows):        #ѭ����ӡÿһ��
	print(worksheet.row_values(i))
col_data = worksheet.col_values(0)
print(col_data)               #��ȡ��һ�е�����

#%%
#1.5 ��ȡĳһ����Ԫ�������(��xlrdģ���У���������к��ж��Ǵ�0��ʼ�����ġ�)
#ͨ�������ȡ����е�����
cell_value1=worksheet.cell_value(0,0)
cell_value2=worksheet.cell_value(1,1)
print(cell_value1)
print(cell_value2)

#%%
cell_value1 =worksheet.cell(0,0).value
print(cell_value1)
cell_value2 =worksheet.row(0)[0].value
print(cell_value2)

