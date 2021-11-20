import xlrd
import xlwt

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
credit_info = xlrd.open_workbook("小额信贷.xlsx.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
credit_sheet = credit_info.sheet_by_index(1)

book_write = xlwt.Workbook()
sheet_to_write = book_write.add_sheet("sheet0")

dic_credit_info = {}

count = 0

for i in range(0, credit_sheet.nrows):
    content = credit_sheet.row_values(i)
    county = content[1]
    name = content[3]
    dic_credit_info[(county, name)] = i


for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    county = content[3]
    name = content[2]
    if (county, name) not in dic_credit_info:
        continue
    row_value = dic_credit_info[(county, name)]
    row_content = credit_sheet.row_values(row_value)
    for k in range(0, len(row_content)):
        sheet_to_write.write(count, k, row_content[k])
    count += 1

book_write.save("小额信贷.xls")
