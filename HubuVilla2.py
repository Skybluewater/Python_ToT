import xlrd
import xlwt

OutProvince = {"北京", "大理", "杭州", "江苏", "陕西", "上海", "省外", "新疆"}

workbook = xlrd.open_workbook("就业信息表.xlsx.xlsx")
sheet = workbook.sheet_by_index(0)

info = xlrd.open_workbook("脱贫户信息统计表.xls.xlsx")
info_sheet = info.sheet_by_index(0)

party_member_book = xlrd.open_workbook("党员.xls")
party_member = party_member_book.sheet_by_index(1)

dic_info = {}
dic_party = set()

for i in range(0, info_sheet.nrows):
    values = info_sheet.row_values(i)
    name = values[7]
    if name is "":
        break
    else:
        dic_info[name] = i

for i in range(0, party_member.nrows):
    values = party_member.row_values(i)
    name = values[1]
    dic_party.add(name)

inProvince = xlwt.Workbook()
inProvince_sheet = inProvince.add_sheet("sheet0")

outProvince = xlwt.Workbook()
outProvince_sheet = outProvince.add_sheet("sheet0")

inCounty = xlwt.Workbook()
inCounty_sheet = inCounty.add_sheet("sheet0")

inProvince_count, inCounty_count, outProvince_count = 0, 0, 0
for i in range(0, sheet.nrows):
    values = sheet.row_values(i)
    if values[1] is "":
        break
    if values[9] in {"务农", "上学", "上大学", "在家养病", "技术工程学校", "", "服役"}:
        continue

    name = values[1]
    sex = values[2]
    birth_date = values[3][6:12]
    edu = values[4]
    work_loc = values[8]
    work_type = values[9]
    start_time = values[10]
    monthly_income = values[11]

    def filling_values(book, count):
        if name in dic_info:
            line_content = info_sheet.row_values(dic_info[name])
            work_time, income_per_person = line_content[16], line_content[21]
        else:
            work_time, income_per_person = "", ""
        if name in dic_party:
            party = "党员"
        else: party = "群众"
        line_to_fill = [count + 1, name, sex, birth_date, edu, party, start_time, work_loc, work_type,
                        monthly_income, work_time]
        if work_time is "" or monthly_income is "":
            line_to_fill.append(income_per_person)
        else:
            line_to_fill.append(str(int(monthly_income) * int(work_time)))
        for k in range(0, len(line_to_fill)):
            book.write(count, k, line_to_fill[k])
        return count + 1


    if work_loc in OutProvince:
        outProvince_count = filling_values(outProvince_sheet, outProvince_count)
    elif work_loc in {"县内"}:
        inCounty_count = filling_values(inCounty_sheet, inCounty_count)
    else:
        inProvince_count = filling_values(inProvince_sheet, inProvince_count)

inCounty.save("务工县内.xls")
inProvince.save("务工省内.xls")
outProvince.save("务工省外.xls")
