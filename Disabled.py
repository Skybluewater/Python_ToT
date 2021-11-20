import xlrd
import xlwt

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
disabled_book = xlrd.open_workbook("残疾人.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
disable_sheet = disabled_book.sheet_by_index(0)

dic_disable = {}

book_write = xlwt.Workbook()
disable_write = book_write.add_sheet("sheet0")
disable_count = 0

for i in range(0, disable_sheet.nrows):
    content = disable_sheet.row_values(i)
    persona_name = content[0]
    dic_disable[persona_name] = i


for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    persona_name = content[2]
    villa_name = content[3]

    def handle_diff_sheet(sheet_write, line_content, count) -> int:
        for k in range(0, len(line_content)):
            sheet_write.write(count, k, line_content[k])
        return count + 1

    if persona_name in dic_disable:
        co = disable_sheet.row_values(dic_disable[persona_name])
        disable_count = handle_diff_sheet(disable_write, co, disable_count)

book_write.save("残疾人.xls")
