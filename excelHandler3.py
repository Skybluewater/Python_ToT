import xlwt
from xlrd import open_workbook

book1 = open_workbook("firstVerifyList.xlsx")
book2 = open_workbook("enrollBase.xlsx")
sheet1 = book1.sheet_by_index(0)
sheet2 = book2.sheet_by_index(0)
bookWrite = xlwt.Workbook()
sheetWrite = bookWrite.add_sheet("sheetNow")

d = {}

for row in range(1, sheet2.nrows):
    ct1 = sheet2.row_values(row)
    if ct1[2] not in d.keys():
        d[ct1[2]] = ct1[1]
    else:
        raise Exception

for row in range(2, sheet1.nrows):
    ct1 = sheet1.row_values(row)
    if len(ct1) > 4 and ct1[4] in d.keys():
        ct1.insert(4, d[ct1[4]])
    for i in range(len(ct1)):
        sheetWrite.write(row, i, ct1[i])

bookWrite.save("bookWrite.xlsx")