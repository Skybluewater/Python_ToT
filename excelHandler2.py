import xlwt
from xlrd import open_workbook

book1 = open_workbook("sheet1.xlsx")
book2 = open_workbook("sheet2.xlsx")
sheet1 = book1.sheet_by_index(0)
sheet2 = book2.sheet_by_index(0)
bookWrite = xlwt.Workbook()
sheetWrite = bookWrite.add_sheet("sheet0")
leader = (3.6, 3.6, "团长")
worker = (3, 3, "普通队员")
excellentLeader = (5.4, 5.4, "校级暑期社会实践优秀团队或个人/团长")
excellentWorker = (4.5, 4.5, "校级暑期社会实践优秀团队或个人")

excellentPeople = []

for row in range(sheet2.nrows):
    content = sheet2.row_values(row)
    excellentPeople.append(content[0])

for row in range(2, sheet1.nrows):
    ct1 = sheet1.row_values(row)
    ifIsLeader = 0
    ifIsExcellent = 0
    if ct1[3] == ct1[10]:
        ifIsLeader = 1
    if ct1[3] in excellentPeople:
        ifIsExcellent = 1
    if ifIsExcellent == 1 and ifIsLeader == 1:
        ct1[7], ct1[8], ct1[9] = excellentLeader
    elif ifIsLeader == 1:
        ct1[7], ct1[8], ct1[9] = leader
    elif ifIsExcellent == 1:
        ct1[7], ct1[8], ct1[9] = excellentWorker
    else:
        ct1[7], ct1[8], ct1[9] = worker
    for i in range(len(ct1)):
        sheetWrite.write(row, i, ct1[i])

bookWrite.save("bookWrite.xlsx")

