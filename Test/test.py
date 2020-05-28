"""
import xlrd
path = "..\\data\\TestData.xlsx"

inputXl = xlrd.open_workbook(path)
inputSheet = inputXl.sheet_by_index(0)

rows = inputSheet.nrows
cols = inputSheet.ncols

for i in range (1,rows):
    date = []
    date.append(inputSheet.cell_value(i,0))
    print(date)


import xlwt
wb = xlwt.Workbook()
ws = wb.add_sheet("Test")
ws.write(0,0,"abc")

wb.save("..\\output\\write.xls")
"""
import pyexcel as p
p.save_as(file_name='..\\output\\Result.xlsx', dest_file_name='..\\output\\Results.html')













