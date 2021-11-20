import xlrd
import xlwt

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
agricultural_credit_union = xlrd.open_workbook("石站头.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
agricultural_credit_union_sheet_1 = agricultural_credit_union.sheet_by_index(0)

# book_write = xlwt.Workbook()
# agricultural_credit_union_write_1 = book_write.add_sheet("sheet1")
# agricultural_credit_union_count_1 = 0

book_write = xlwt.Workbook()
agricultural_credit_union_write_1 = book_write.add_sheet("sheet1")
agricultural_credit_union_count_1 = 0


def handle_dictionary(sheet):
    dic_t = {}
    set_t = set()
    for x in range(0, sheet.nrows):
        content_t = sheet.row_values(x)
        # villa_name_t = ""
        villa_name_t = content_t[2].split("村")[0]
        persona_name_t = content_t[5]
        dic_t[persona_name_t] = x
        set_t.add(villa_name_t)
    return dic_t, set_t


dic_agricultural_credit_union_1, set_agricultural_credit_union_1 = handle_dictionary(agricultural_credit_union_sheet_1)

for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    persona_name = content[2]
    villa_name = content[3]


    def handle_diff_sheet(sheet_write, line_content, count) -> int:
        for k in range(0, len(line_content)):
            sheet_write.write(count, k, line_content[k])
        return count + 1


    if villa_name is "":
        if persona_name in dic_agricultural_credit_union_1:
            co = agricultural_credit_union_sheet_1.row_values(dic_agricultural_credit_union_1[persona_name])
            agricultural_credit_union_count_1 = handle_diff_sheet(agricultural_credit_union_write_1, co,
                                                                  agricultural_credit_union_count_1)

    else:
        flag1 = False
        if villa_name in set_agricultural_credit_union_1:
            flag1 = True
        if flag1 and persona_name in dic_agricultural_credit_union_1:
            co = agricultural_credit_union_sheet_1.row_values(dic_agricultural_credit_union_1[persona_name])
            agricultural_credit_union_count_1 = handle_diff_sheet(agricultural_credit_union_write_1, co,
                                                                  agricultural_credit_union_count_1)

        # flag2 = False
        # if villa_name in set_agricultural_credit_union_2:
        #     flag2 = True
        # if flag2 and persona_name in dic_agricultural_credit_union_2:
        #     co = agricultural_credit_union_sheet_2.row_values(dic_agricultural_credit_union_2[persona_name])
        #     agricultural_credit_union_count_2 = handle_diff_sheet(agricultural_credit_union_write_2, co, agricultural_credit_union_count_2)
        # 
        # flag3 = False
        # if villa_name in set_agricultural_credit_union_3:
        #     flag3 = True
        # if flag3 and persona_name in dic_agricultural_credit_union_3:
        #     co = agricultural_credit_union_sheet_3.row_values(dic_agricultural_credit_union_3[persona_name])
        #     agricultural_credit_union_count_3 = handle_diff_sheet(agricultural_credit_union_write_3, co, agricultural_credit_union_count_3)

book_write.save("石站头.xls")
