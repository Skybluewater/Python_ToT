import xlrd
import xlwt

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
medic_helper = xlrd.open_workbook("护工xlsx.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
medic_helper_sheet_1 = medic_helper.sheet_by_index(0)

book_write = xlwt.Workbook()
medic_helper_write_1 = book_write.add_sheet("sheet1")
medic_helper_count_1 = 0


# s = set()
#
# for x in range(0, housing_sheet.nrows):
#     c = housing_sheet.row_values(x)
#     s.add(c[3])
#
# print(s)


def handle_dictionary(sheet):
    dic = {}
    for x in range(0, sheet.nrows):
        content_t = sheet.row_values(x)
        villa_name_t = ""
        persona_name_t = content_t[1]
        if (villa_name_t, persona_name_t) not in dic:
            dic[(villa_name_t, persona_name_t)] = x
    return dic


dic_medic_helper_1 = handle_dictionary(medic_helper_sheet_1)

for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    persona_name = content[2]
    villa_name = ""
    # villa_name = content[3]

    def handle_diff_sheet(sheet_write, line_content, count) -> int:
        for k in range(0, len(line_content)):
            sheet_write.write(count, k, line_content[k])
        return count + 1


    if (villa_name, persona_name) in dic_medic_helper_1:
        co = medic_helper_sheet_1.row_values(dic_medic_helper_1[(villa_name, persona_name)])
        medic_helper_count_1 = handle_diff_sheet(medic_helper_write_1, co, medic_helper_count_1)

book_write.save("马坊镇护工.xls")
