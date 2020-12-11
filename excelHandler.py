import xlwt
from xlrd import open_workbook

leader = (3.6, 3.6, "团长")
worker = (3, 3, "普通队员")
book = open_workbook('2.xlsx')
sheet0 = book.sheet_by_index(0)
bookWrite = xlwt.Workbook()
sheet1 = bookWrite.add_sheet("sheet0")
flag = 0
print(sheet0.nrows, sheet0.ncols)
for row in range(2, sheet0.nrows - 1):
    st1 = sheet0.row_values(row)
    st2 = sheet0.row_values(row + 1)
    # print(st1)
    if st1[2] == st2[2] and flag == 0:
        st1[8], st1[7], st1[9] = leader
        flag = 1
    elif flag == 1 and st1[2] == st2[2]:
        st1[8], st1[7], st1[9] = worker
    elif st1[2] != st2[2] and flag == 1:
        st1[8], st1[7], st1[9] = worker
        flag = 0
    elif st1[2] != st2[2] and flag == 0:
        st1[8], st1[7], st1[9] = leader
    for i in range(len(st1)):
        sheet1.write(row, i, st1[i])

st1 = sheet0.row_values(sheet0.nrows - 1)
if flag == 0:
    st1[8], st1[7], st1[9] = leader
else:
    st1[8], st1[7], st1[9] = worker
for i in range(len(st1)):
    sheet1.write(sheet0.nrows - 1, i, st1[i])
bookWrite.save("writeBook.xlsx")