import xlrd
import xlwt

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
fm_info = xlrd.open_workbook("退耕还林.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
farming_sheet = fm_info.sheet_by_index(1)

book_write = xlwt.Workbook()
sheet_to_write = book_write.add_sheet("sheet0")

dic_farming_info = {}

count = 0

for i in range(0, farming_sheet.nrows):
    content = farming_sheet.row_values(i)
    name = content[1]
    dic_farming_info[name] = i


for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    name = content[2]
    if name not in dic_farming_info:
        continue
    row_value = dic_farming_info[name]
    row_content = farming_sheet.row_values(row_value)
    for k in range(0, len(row_content)):
        sheet_to_write.write(count, k, row_content[k])
    count += 1

book_write.save("退耕.xls")
