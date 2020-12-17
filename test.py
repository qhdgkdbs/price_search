import xlrd
import xlsxwriter

from os.path import expanduser
home = expanduser("~")

# this writes test data to an excel file
wb = xlsxwriter.Workbook("test.xlsx")
sheet1 = wb.add_worksheet()
for row in range(10):
    for col in range(20):
        sheet1.write(row, col, "test ({}, {})".format(row, col))
wb.close()

# open the file for reading
wbRD = xlrd.open_workbook("test.xlsx")
sheets = wbRD.sheets()

# open the same file for writing (just don't write yet)
wb = xlsxwriter.Workbook("test.xlsx")

# run through the sheets and store sheets in workbook
# this still doesn't write to the file yet
for sheet in sheets: # write data from old file
    newSheet = wb.add_worksheet(sheet.name)
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            newSheet.write(row, col, sheet.cell(row, col).value)

for row in range(10, 20): # write NEW data
    for col in range(20):
        newSheet.write(row, col, "test ({}, {})".format(row, col))
wb.close() # THIS writes