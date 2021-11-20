import xlrd
import xlwt

OutProvince = {"北京", "大理", "杭州", "江苏", "陕西", "上海", "省外", "新疆", "河南", "河北", "内蒙", "海口", "福建", "山东", "浙江", "广西"}
InProvince = {"太原", "岚县", "宁武", "忻州", "省内", "原平", "长治", "兴县", "离石", "古交", "娄烦", "朔州", "临汾", "柳林"}

work_book = xlrd.open_workbook("务工台账.xlsx")
work_sheet = work_book.sheet_by_index(3)
housing_info = xlrd.open_workbook("住户信息.xlsx.xlsx")
housing_sheet = housing_info.sheet_by_index(0)

# info = xlrd.open_workbook("脱贫户信息统计表.xls.xlsx")
# info_sheet = info.sheet_by_index(0)
#
# party_member_book = xlrd.open_workbook("党员.xls")
# party_member = party_member_book.sheet_by_index(1)

dic_info = {}
dic_party = set()

for i in range(0, work_sheet.nrows):
    values = work_sheet.row_values(i)
    name = values[2]
    if name is "":
        break
    else:
        dic_info[name] = i

# for i in range(0, party_member.nrows):
#     values = party_member.row_values(i)
#     name = values[1]
#     dic_party.add(name)

inProvince = xlwt.Workbook()
inProvince_sheet = inProvince.add_sheet("sheet0")

outProvince = xlwt.Workbook()
outProvince_sheet = outProvince.add_sheet("sheet0")

inCounty = xlwt.Workbook()
inCounty_sheet = inCounty.add_sheet("sheet0")

inProvince_count, inCounty_count, outProvince_count = 0, 0, 0

for i in range(0, housing_sheet.nrows):
    values = housing_sheet.row_values(i)
    name = values[2]
    if name not in dic_info:
        continue

    content = work_sheet.row_values(dic_info[name])
    s = str(content[5])[16]
    sex = "男" if int(s) % 2 == 1 else "女"
    birth = content[5][6: 14]
    work_loc = "县内"
    for key in OutProvince:
        if key in content[6]:
            work_loc = key
            break
    for key in InProvince:
        if key in content[6]:
            work_loc = key
            break
    work_type = content[6]
    monthly_income = content[7]

    def filling_values(book, count):
        line_to_fill = [count + 1, name, sex, birth, "", "", "", work_loc, work_type,
                        monthly_income, "", ""]
        # if work_time is "" or monthly_income is "":
        #     line_to_fill.append(income_per_person)
        # else:
        #     line_to_fill.append(str(int(monthly_income) * int(work_time)))
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
