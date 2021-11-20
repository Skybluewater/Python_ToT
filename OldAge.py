import xlrd
import xlwt

housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
old_age_info = xlrd.open_workbook("城乡低保失能高龄.xlsx")
housing_sheet = housing_info.sheet_by_index(0)
old_age_sheet = old_age_info.sheet_by_index(1)
disable_sheet = old_age_info.sheet_by_index(0)

book_write = xlwt.Workbook()
old_age_write = book_write.add_sheet("高龄")
disable_write = book_write.add_sheet("失能")
old_age_count, disable_count = 0, 0

# for i in range(0, old_age_sheet.nrows):
#     content = old_age_sheet.row_values(i)
#     villa_name = content[0][0: -3]
#     persona_name = content[3]
#     dic_old_age[(villa_name, persona_name)] = i
#
# for i in range(0, disable_sheet.nrows):
#     content = disable_sheet.row_values(i)
#     villa_name = content[0][0: -3]
#     persona_name = content[3]
#     dic_disable[(villa_name, persona_name)] = i
#
# for i in range(0, housing_sheet.nrows):
#     content = housing_sheet.row_values(i)
#     persona_name = content[2]
#     villa_name = content[3]
#
#     def handle_diff_sheet(sheet_write, line_content, count) -> int:
#         for k in range(0, len(line_content)):
#             sheet_write.write(count, k, line_content[k])
#         return count + 1
#
#     if (villa_name, persona_name) in dic_old_age:
#         co = old_age_sheet.row_values(dic_old_age[(villa_name, persona_name)])
#         old_age_count = handle_diff_sheet(old_age_write, co, old_age_count)
#     if (villa_name, persona_name) in dic_disable:
#         co = disable_sheet.row_values(dic_disable[(villa_name, persona_name)])
#         disable_count = handle_diff_sheet(disable_write, co, disable_count)


def handle_dictionary(sheet):
    dic_t = {}
    set_t = set()
    for x in range(0, sheet.nrows):
        content_t = sheet.row_values(x)
        # villa_name_t = ""
        villa_name_t = content_t[0].split("村")[0]
        persona_name_t = content_t[3]
        dic_t[persona_name_t] = x
        set_t.add(villa_name_t)
    return dic_t, set_t


dic_old_age, set_old_age = handle_dictionary(old_age_sheet)
dic_disable, set_disable = handle_dictionary(disable_sheet)

for i in range(0, housing_sheet.nrows):
    content = housing_sheet.row_values(i)
    persona_name = content[2]
    villa_name = content[3]


    def handle_diff_sheet(sheet_write, line_content, count) -> int:
        for k in range(0, len(line_content)):
            sheet_write.write(count, k, line_content[k])
        return count + 1


    if villa_name is "":
        if persona_name in dic_old_age:
            co = old_age_sheet.row_values(dic_old_age[persona_name])
            old_age_count = handle_diff_sheet(old_age_write, co, old_age_count)
        if persona_name in dic_disable:
            co = disable_sheet.row_values(dic_disable[persona_name])
            disable_count = handle_diff_sheet(disable_write, co, disable_count)
    else:
        if villa_name in set_old_age and persona_name in dic_old_age:
            co = old_age_sheet.row_values(dic_old_age[persona_name])
            old_age_count = handle_diff_sheet(old_age_write, co, old_age_count)
        if villa_name in set_disable and persona_name in dic_disable:
            co = disable_sheet.row_values(dic_disable[persona_name])
            disable_count = handle_diff_sheet(disable_write, co, disable_count)


book_write.save("高龄、失能.xls")
