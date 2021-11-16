OutProvince = {"北京", "大理", "杭州", "江苏", "陕西", "上海", "省外", "新疆"}

import xlwt
from xlrd import open_workbook

jiuye_book = open_workbook('就业信息表.xlsx.xlsx')
info_book = open_workbook("脱贫户信息统计表.xls.xlsx")
party_member_book = open_workbook("党员.xls")
jiu_ye = jiuye_book.sheet_by_index(0)
info = info_book.sheet_by_index(0)
party_member = party_member_book.sheet_by_index(1)


bookWrite_outProcince = xlwt.Workbook()
outProvince = bookWrite_outProcince.add_sheet("sheet0")
outProvince_count = 0

bookWrite_inProvince = xlwt.Workbook()
inProvince = bookWrite_inProvince.add_sheet("sheet0")
inProvince_count = 0

bookWrite_inCount = xlwt.Workbook()
inCounty = bookWrite_inCount.add_sheet("sheet0")
inCounty_count = 0

dic_jiuye = {}
dic_party_member = set()

for i in range(0, jiu_ye.nrows):
    values = jiu_ye.row_values(i)
    name = values[1]
    dic_jiuye[name] = i

for i in range(0, party_member.nrows):
    values = party_member.row_values(i)
    name = values[1]
    dic_party_member.add(name)

for i in range(0, info.nrows):
    values = info.row_values(i)
    if values[7] is not "":
        name = values[7]
    else:
        break
    if name in dic_jiuye:
        line_num = dic_jiuye[name]
    else:
        continue
    values_jiuye = jiu_ye.row_values(line_num)
    if values_jiuye[9] in {"务农", "上学", "上大学", "在家养病", "技术工程学校", "", "服役"}:
        continue
    sex = values_jiuye[2]
    birth_date = values_jiuye[3][6:12]
    edu = values_jiuye[4]
    out_year = values_jiuye[10]
    monthly_income = values_jiuye[11]
    out_time = values[16]
    income_per_p = values[21]
    labor_loc = values_jiuye[8]
    labor_type = values_jiuye[9]


    def sheet_handler(book_count, book_name) -> int:
        book_count += 1
        if name in dic_party_member:
            party = "党员"
        else:
            party = "群众"
        line_to_fill = [book_count, name, sex, birth_date, edu, party, out_year, labor_loc, labor_type,
                        monthly_income, out_time]
        if out_time is "" or monthly_income is "":
            total_income = income_per_p
        else:
            total_income = int(monthly_income) * int(out_time)
        line_to_fill.append(total_income)
        for k in range(0, len(line_to_fill)):
            book_name.write(book_count - 1, k, line_to_fill[k])
        return book_count

    if labor_loc in {"县内"}:
        inCounty_count = sheet_handler(inCounty_count, inCounty)
    elif labor_loc in OutProvince:
        outProvince_count = sheet_handler(outProvince_count, outProvince)
    else:
        inProvince_count = sheet_handler(inProvince_count, inProvince)

bookWrite_inProvince.save("脱贫省内.xls")
bookWrite_outProcince.save("脱贫省外.xls")
bookWrite_inCount.save("脱贫县内.xls")
