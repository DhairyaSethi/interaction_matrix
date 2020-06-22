import xlsxwriter

workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()

data = ["car body", "windshield", "motor", "door", "rear view mirror"]
row = 1
column = 1
i = 0

length_list = [len(x) for x in data]
max_width = max(length_list)

cell_bold = workbook.add_format({'bold': True, 'center_across': True, 'bg_color': '#f0f0c7'})
cell_grey = workbook.add_format()
cell_grey.set_bg_color('#D3D3D3')

for item in data:
	worksheet.write(row, 0, item, cell_bold)
	worksheet.write(0, column, item, cell_bold)
	worksheet.write(row, column,'' , cell_grey)
	worksheet.set_column(i, i, max_width)
	row+=1
	column+=1
	i+=1
	print('added ' + item + ' at R:C', row, column, i)
worksheet.set_column(len(data), len(data), max_width)


workbook.close()