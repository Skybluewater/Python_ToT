import xlrd
import xlwt

agricultural_book = xlrd.open_workbook("圪洞2021年一次性补贴花名表.xls")
housing_book = xlrd.open_workbook("住户信息.xlsx.xlsx")
housing_info = housing_book.sheet_by_index(0)

book_to_write = xlwt.Workbook()


def handle_dictionary(sheet):
    dic_temp = {}
    for k in range(2, sheet.nrows):
        content_temp = sheet.row_values(k)
        name_temp = content_temp[1]
        villa_temp = content_temp[7].split(" ")[1].split("村")[0]
        dic_temp[name_temp] = k
    return dic_temp


for i in range(0, agricultural_book.nsheets):
    current_sheet = agricultural_book.sheet_by_index(i)

    dic_person = handle_dictionary(current_sheet)
    sheet_to_write = book_to_write.add_sheet(current_sheet.name)
    count = 0


    def filling_sheet(sheet, co):
        for index in range(0, len(co)):
            sheet.write(count, index, co[index])


    for k in range(0, housing_info.nrows):
        content = housing_info.row_values(k)
        personnal_name = content[2]
        villa_name = content[3]
        if personnal_name in dic_person:
            filling_sheet(sheet_to_write, current_sheet.row_values(dic_person[personnal_name]))
            count += 1

book_to_write.save("一次性补贴.xls")
