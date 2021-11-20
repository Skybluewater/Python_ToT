import xlrd
import xlwt

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
medic_helper = xlrd.open_workbook("小额信贷.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
small_credit_sheet_1 = medic_helper.sheet_by_index(1)

book_write = xlwt.Workbook()
small_credit_write_1 = book_write.add_sheet("sheet1")
small_credit_count_1 = 0


# book_write = xlwt.Workbook()
# small_credit_write_1 = book_write.add_sheet("sheet1")
# small_credit_count_1 = 0


def handle_dictionary(sheet):
    dic_t = {}
    set_t = set()
    for x in range(0, sheet.nrows):
        content_t = sheet.row_values(x)
        # villa_name_t = ""
        villa_name_t = content_t[1].split("村")[0]
        persona_name_t = content_t[3]
        dic_t[persona_name_t] = x
        set_t.add(villa_name_t)
    return dic_t, set_t


dic_small_credit_1, set_small_credit_1 = handle_dictionary(small_credit_sheet_1)

for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    persona_name = content[2]
    villa_name = content[3]


    def handle_diff_sheet(sheet_write, line_content, count) -> int:
        for k in range(0, len(line_content)):
            sheet_write.write(count, k, line_content[k])
        return count + 1


    if villa_name is "":
        if persona_name in dic_small_credit_1:
            co = small_credit_sheet_1.row_values(dic_small_credit_1[persona_name])
            small_credit_count_1 = handle_diff_sheet(small_credit_write_1, co,
                                                     small_credit_count_1)

    else:
        if villa_name in set_small_credit_1 and persona_name in dic_small_credit_1:
            co = small_credit_sheet_1.row_values(dic_small_credit_1[persona_name])
            small_credit_count_1 = handle_diff_sheet(small_credit_write_1, co,
                                                     small_credit_count_1)

        # flag2 = False
        # if villa_name in set_small_credit_2:
        #     flag2 = True
        # if flag2 and persona_name in dic_small_credit_2:
        #     co = small_credit_sheet_2.row_values(dic_small_credit_2[persona_name])
        #     small_credit_count_2 = handle_diff_sheet(small_credit_write_2, co, small_credit_count_2)
        # 
        # flag3 = False
        # if villa_name in set_small_credit_3:
        #     flag3 = True
        # if flag3 and persona_name in dic_small_credit_3:
        #     co = small_credit_sheet_3.row_values(dic_small_credit_3[persona_name])
        #     small_credit_count_3 = handle_diff_sheet(small_credit_write_3, co, small_credit_count_3)

book_write.save("小额信贷.xls")
