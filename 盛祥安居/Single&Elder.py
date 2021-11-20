import xlrd
import xlwt

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
old_age_info = xlrd.open_workbook("高龄孤寡.xlsx.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
old_age_sheet = old_age_info.sheet_by_index(0)

book_write = xlwt.Workbook()
old_age_write = book_write.add_sheet("高龄")
single_old_write = book_write.add_sheet("孤寡")
old_age_count, single_old_count = 0, 0


def handle_dictionary(sheet):
    dic_t = {}
    set_t = set()
    next_different, count = 0, 0
    for x in range(0, sheet.nrows):
        content_t = sheet.row_values(x)
        # villa_name_t = ""
        villa_name_t = content_t[3].split("村")[0]
        persona_name_t = content_t[4]
        birth_year = content_t[5][6:10]
        is_single = False
        if x == next_different:
            count = 1
            for x_t in range(x + 1, sheet.nrows):
                if sheet.row_values(x_t)[6] != "户主":
                    continue
                else:
                    next_different = x_t
                    count = x_t - x
                    break

        if count == 1:
            is_single = True
        if int(birth_year) <= 1941:
            dic_t[persona_name_t] = [x, villa_name_t, birth_year, is_single]

    return dic_t, set_t


dic_old_age, set_old_age = handle_dictionary(old_age_sheet)

for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    persona_name = content[2]
    villa_name = content[3]


    def handle_diff_sheet(sheet_write, line_content, count) -> int:
        for k in range(0, len(line_content)):
            sheet_write.write(count, k, line_content[k])
        return count + 1

    # if villa_name is "":
    if persona_name in dic_old_age:
        co = old_age_sheet.row_values(dic_old_age[persona_name][0])
        old_age_count = handle_diff_sheet(old_age_write, co, old_age_count)
        if dic_old_age[persona_name][3] is True:
            single_old_count = handle_diff_sheet(single_old_write, co, single_old_count)
    # else:
    #     if persona_name in dic_old_age and villa_name == dic_old_age[persona_name][1]:
    #         co = old_age_sheet.row_values(dic_old_age[persona_name][0])
    #         old_age_count = handle_diff_sheet(old_age_write, co, old_age_count)
    #         if dic_old_age[persona_name][3] is True:
    #             single_old_count = handle_diff_sheet(single_old_write, co, single_old_count)

book_write.save("马坊高龄孤寡.xls")
