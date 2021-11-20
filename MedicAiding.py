import xlrd
import xlwt

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
medic_aid = xlrd.open_workbook("医疗名单.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
medic_aid_sheet_1 = medic_aid.sheet_by_index(0)
medic_aid_sheet_2 = medic_aid.sheet_by_index(1)
medic_aid_sheet_3 = medic_aid.sheet_by_index(2)

book_write = xlwt.Workbook()
medic_aid_write_1 = book_write.add_sheet("sheet1")
medic_aid_write_2 = book_write.add_sheet("sheet2")
medic_aid_write_3 = book_write.add_sheet("sheet3")
medic_aid_count_1, medic_aid_count_2, medic_aid_count_3 = 0, 0, 0


def handle_dictionary(sheet):
    dic = {}
    for x in range(0, sheet.nrows):
        content_t = sheet.row_values(x)
        villa_name_t = content_t[1][0: -1]
        persona_name_t = content_t[2]
        dic[(villa_name_t, persona_name_t)] = x
    return dic


dic_medic_aid_1 = handle_dictionary(medic_aid_sheet_1)
dic_medic_aid_2 = handle_dictionary(medic_aid_sheet_2)
dic_medic_aid_3 = handle_dictionary(medic_aid_sheet_3)

for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    persona_name = content[2]
    villa_name = content[3]


    def handle_diff_sheet(sheet_write, line_content, count) -> int:
        for k in range(0, len(line_content)):
            sheet_write.write(count, k, line_content[k])
        return count + 1


    if (villa_name, persona_name) in dic_medic_aid_1:
        co = medic_aid_sheet_1.row_values(dic_medic_aid_1[(villa_name, persona_name)])
        medic_aid_count_1 = handle_diff_sheet(medic_aid_write_1, co, medic_aid_count_1)
    if (villa_name, persona_name) in dic_medic_aid_2:
        co = medic_aid_sheet_2.row_values(dic_medic_aid_2[(villa_name, persona_name)])
        medic_aid_count_2 = handle_diff_sheet(medic_aid_write_2, co, medic_aid_count_2)
    if (villa_name, persona_name) in dic_medic_aid_3:
        co = medic_aid_sheet_3.row_values(dic_medic_aid_3[(villa_name, persona_name)])
        medic_aid_count_3 = handle_diff_sheet(medic_aid_write_3, co, medic_aid_count_3)

book_write.save("医疗救助.xls")
