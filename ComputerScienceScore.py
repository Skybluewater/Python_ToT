import xlrd
import xlwt

score_book = xlrd.open_workbook("职高计算机.xlsx")
score_sheet = score_book.sheet_by_index(0)

book_write = xlwt.Workbook()
write_sheet = book_write.add_sheet("sheet0")

max_count = 0

for i in range(0, score_sheet.nrows):
    content = score_sheet.row_values(i)
    if content[3] is "":
        continue
    max_count = max_count if content[3] - content[4] < max_count else content[3] - content[4]

for i in range(0, score_sheet.nrows):
    content = score_sheet.row_values(i)
    if content[2] is "" and content[3] is "":
        write_sheet.write(i, 0, "")
        continue
    elif content[2] is "":
        score = ((content[3] - content[4]) / max_count) * 20
    elif content[3] is "":
        score = content[2] * 0.8
        if score < 20:
            score = 20
    else:
        score = 20 if content[2] * 0.8 < 20 else content[2] * 0.8
        score += ((content[3] - content[4]) / max_count) * 20
    write_sheet.write(i, 0, round(score))

book_write.save("计算机成绩.xls")
