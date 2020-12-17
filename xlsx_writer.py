import xlsxwriter

workbook = xlsxwriter.Workbook('arrays.xlsx')
worksheet = workbook.add_worksheet()

array = [['a1', 'a2', 'a3'],
         ['a4', 'a5', 'a6'],
         ['a7', 'a8', 'a9'],
         ['a10', 'a11', 'a12', 'a13', 'a14']]

col = 0

worksheet.write(0, 1, "data")
data_format1 = workbook.add_format({'bg_color': '#FFC7CE'})


# for row, data in enumerate(array):
#     # print(col)
#     # print(data)
#     worksheet.write_row(row, col, data)

for i in range(20):
    index = 4*i-3
    if (i % 2) == 0:
        worksheet.set_column(index,index+3, cell_format=data_format1)

workbook.close()