import xlrd
import xlwt
import datetime


def compare():
    readfile = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\1.xls')
    readfile1 = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\2.xls')
    sheet = readfile.sheet_by_name("sheet")
    sheet1 = readfile1.sheet_by_name("sheet")
    row = sheet.nrows
    row1 = sheet1.nrows
    user = []
    user1 = []

    for i in range(0, row):
        user.append([sheet.cell_value(i, 0), sheet.cell_value(i, 1), sheet.cell_value(i, 2)])

    for j in range(0, row1):
        user1.append([sheet1.cell_value(j, 0), sheet1.cell_value(j, 1), sheet1.cell_value(j, 2)])

    # print(user1)
    # print(len(user1))
    # print('----------------------------------------------------')
    # print(user)
    # print(len(user))
    usertrue = []
    userfalse = []
    for i in range(0, len(user)):  # range(0,10):
        for j in range(0, len(user1)):  # range(0,10):
            if user[i][1] == user1[j][1]:
                if user[i][0] != user1[j][0] or user[i][2] != user1[j][2]:
                    # print(user[i][1]+'--'+user[i][0]+'--'+user[i][2])
                    # print(user1[j][1] + '--' + user1[j][0] + '--' + user1[j][2])

                    usertrue.append([user[i][1], user[i][0], user[i][2]])
                    userfalse.append([user1[j][1], user1[j][0], user1[j][2]])
                    print(user[i])
                    print('----------------------------')

                    pass
                    # print(usernamelist[i])
                    # print(i)
    return [usertrue, userfalse]


def write_excel():
    total_list = compare()
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)
    num = 0
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 2  # 5 背景颜色为黄色
    # 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon,
    # 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray

    style = xlwt.XFStyle()
    style.pattern = pattern

    for item in total_list[0]:
        sheet.write(num, 0, item[0])
        sheet.write(num, 1, item[1])
        sheet.write(num, 2, item[2])
        sheet.write(num, 3, '')
        sheet.write(num, 4, total_list[1][num][0])
        sheet.write(num, 5, total_list[1][num][1])
        sheet.write(num, 6, total_list[1][num][2])
        if item[1] != total_list[1][num][1]:
            sheet.write(num, 5, total_list[1][num][1], style)
        if item[2] != total_list[1][num][2]:
            sheet.write(num, 6, total_list[1][num][2], style)

        num += 1
    book.save("4A人员账号手机不匹配" + ".xls")


if __name__ == '__main__':
    write_excel()
