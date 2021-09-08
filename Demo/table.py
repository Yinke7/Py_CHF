import openpyxl
import xlrd


import openpyxl
import os

from openpyxl.styles import PatternFill
from openpyxl.styles import Font

temp = 'template.xlsx'
output = "./output.xlsx"
excelfile = './biao.xlsx'


list_all = []
rows = 0
cols = 0

def read_excel():
	"""
	从xlsx文件中读取数据
	"""
	global rows
	global cols
	global list_all

	workbook = openpyxl.load_workbook(filename=temp)
	print(workbook)
	# 可以使用workbook对象的sheetnames属性获取到excel文件中哪些表有数据
	print(workbook.sheetnames)
	table = workbook.active
	print(table)

	rows = table.max_row
	cols = table.max_column


	# context = table.cell(2, 3).value
	# ttpye = type(context)
	# print(ttpye)

	# if ttpye == 'str':
	# 	print("type ok")
	# else:
	# 	print("[err]")

	# list1 = [1,2,3]
	# list1.append(4)
	# print(rows)


	list_Fapiao_num = []
	for row in range(2, rows):
		list_per_row = []


		for col in range(1, cols):
			cell_context = table.cell(row, col).value
			# print('%s' % cell_context)
			list_per_row.append(cell_context)

		# context = list_per_row[2]
		# print('%s' % context)
		if isinstance(list_per_row[2], str):
			list_per_row[2] = list_per_row[2].zfill(8)
		else:
			list_per_row[2] = "%08d" % list_per_row[2]
		# print(context)

		list_all.append(list_per_row)
	# print(list_all)

	# 按第三列排序
	list_all = sorted(list_all, key=(lambda x: x[2]), reverse=False)
	print("list_Fapiao_num %s " % list_all)



	return

def write_excel():
	# excel文件的写入

	wb = openpyxl.Workbook()
	# 新建了一个工作表，尚未保存

	print(wb.get_sheet_names())
	# 默认提供Sheet的表，office 2016 默认新建Sheet1

	# 直接赋值可以改工作表的名称
	sheet = wb.active
	sheet.title = 'Sheet1'

	# 可以对工作表Sheet1 Sheet2 建立索引

	# wb.create_sheet('Data', index=1)

	# 这样新建第二个工作表：名称为Data

	# 删除某个工作表
	# remove 按照变量删除
	# wb.remove(sheet)

	# del 按表格名称删除
	# del wb['Data']

	# 新建一个工作表

	# wb = Workbook()
	# grab the active worksheet
	# sheet = wb.active

	# 直接给单元格赋值
	sheet['A1'] = '序号'
	sheet['B1'] = '发票代码'
	sheet['C1'] = '发票号码（定额票也要填）'
	sheet['D1'] = '发票开具单位名称'
	sheet['E1'] = '发票金额'
	sheet['F1'] = '报销部门(项目）'
	sheet['G1'] = '发票登记日期'
	sheet['H1'] = '报销人'

	# append函数
	# 添加一行
	for row in range(rows - 2):
		sheet.append(list_all[row])
	# wb.save('test.xlsx')


	# for i in range(17):
	# 	sheet.cell(row=i+2, column=1, value=i+1)

	# Save the file
	path = output
	# path = filer + "properties.xlsx"
	if os.path.exists(path):
		print("delete file %s" % path)

	wb.save(path)


	return



def check():

	workbook = openpyxl.load_workbook(filename=output)
	print(workbook)
	# 可以使用workbook对象的sheetnames属性获取到excel文件中哪些表有数据
	print(workbook.sheetnames)
	table = workbook.active

	# font = Font(u'微软雅黑', size=11, bold=True, italic=False, strike=False, color=123456"")  # 设置字体样式

	global list_all

	list_flag = []
	for row in range(rows - 2 - 9):
		if int(list_all[row + 9][2]) - 9 == int(list_all[row][2]):
			list_flag.append(row + 9)

	# print(list_flag)


	fille = PatternFill('solid', fgColor='FF0000')  # 设置填充颜色为 橙色
	for row in range(rows - 2):
		if row in list_flag:
			table.cell(row + 2, 3).fill = fille

	workbook.save(output)








	return