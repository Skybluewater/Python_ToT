import xlrd
import xlwt
import csv

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
medic_helper = xlrd.open_workbook("马坊镇医疗名单.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
medic_aid_sheet_1 = medic_helper.sheet_by_index(0)
medic_aid_sheet_2 = medic_helper.sheet_by_index(1)
medic_aid_sheet_3 = medic_helper.sheet_by_index(2)

book_write = xlwt.Workbook()
medic_aid_write_1 = book_write.add_sheet("sheet1")
medic_aid_count_1 = 0
medic_aid_write_2 = book_write.add_sheet("sheet2")
medic_aid_count_2 = 0
medic_aid_write_3 = book_write.add_sheet("sheet3")
medic_aid_count_3 = 0

dic_p = set()

for i in range(0, housing_sheet.nrows):
    dic_p.add(housing_sheet.row_values(i)[3])


# print(dic_p)

def handle_dictionary(sheet):
    dic_t = {}
    set_t = set()
    for x in range(0, sheet.nrows):
        content_t = sheet.row_values(x)
        # villa_name_t = ""
        villa_name_t = content_t[1].split("村")[0]
        persona_name_t = content_t[2]
        dic_t[persona_name_t] = x
        set_t.add(villa_name_t)
    return dic_t, set_t


dic_medic_aid_1, set_medic_aid_1 = handle_dictionary(medic_aid_sheet_1)
dic_medic_aid_2, set_medic_aid_2 = handle_dictionary(medic_aid_sheet_2)
dic_medic_aid_3, set_medic_aid_3 = handle_dictionary(medic_aid_sheet_3)

for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    persona_name = content[2]
    villa_name = content[3]


    def handle_diff_sheet(sheet_write, line_content, count) -> int:
        for k in range(0, len(line_content)):
            sheet_write.write(count, k, line_content[k])
        return count + 1


    if villa_name is "":
        if persona_name in dic_medic_aid_1:
            co = medic_aid_sheet_1.row_values(dic_medic_aid_1[persona_name])
            medic_aid_count_1 = handle_diff_sheet(medic_aid_write_1, co, medic_aid_count_1)

        if persona_name in dic_medic_aid_2:
            co = medic_aid_sheet_2.row_values(dic_medic_aid_2[persona_name])
            medic_aid_count_2 = handle_diff_sheet(medic_aid_write_2, co, medic_aid_count_2)

        if persona_name in dic_medic_aid_3:
            co = medic_aid_sheet_3.row_values(dic_medic_aid_3[persona_name])
            medic_aid_count_3 = handle_diff_sheet(medic_aid_write_3, co, medic_aid_count_3)

    else:
        flag1 = False
        if villa_name in set_medic_aid_1:
            flag1 = True
        if flag1 and persona_name in dic_medic_aid_1:
            co = medic_aid_sheet_1.row_values(dic_medic_aid_1[persona_name])
            medic_aid_count_1 = handle_diff_sheet(medic_aid_write_1, co, medic_aid_count_1)

        flag2 = False
        if villa_name in set_medic_aid_2:
            flag2 = True
        if flag2 and persona_name in dic_medic_aid_2:
            co = medic_aid_sheet_2.row_values(dic_medic_aid_2[persona_name])
            medic_aid_count_2 = handle_diff_sheet(medic_aid_write_2, co, medic_aid_count_2)

        flag3 = False
        # for key in set_medic_aid_3:
        #     if key.find(villa_name) is not -1:
        #         flag3 = True
        #         break
        if villa_name in set_medic_aid_3: flag3 = True
        if flag3 and persona_name in dic_medic_aid_3:
            co = medic_aid_sheet_3.row_values(dic_medic_aid_3[persona_name])
            medic_aid_count_3 = handle_diff_sheet(medic_aid_write_3, co, medic_aid_count_3)

book_write.save("医疗名单.xls")
