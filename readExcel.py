import json
import os

import xlrd
import xlwt
import datetime


def read(path):
    readfile = xlrd.open_workbook(path)
    # sheetname = readfile.sheet_names()
    sheet = readfile.sheet_by_name("Sheet1")
    row = sheet.nrows
    col = sheet.ncols
    namelist = []
    # for i in range(col):
    #     print(sheet.col_values(i))
    # for i in range(row):
    #     print(sheet.row_values(i))
    if col != 1:
        print("初始化Excel失败，文件格式不正确")
        exit()
    else:
        for i in range(row):
            namelist.append(sheet.cell_value(i, 0))
    print(namelist)
    return namelist


def read_file(path, web_entity):
    readfile = xlrd.open_workbook(path)
    file_dict = {}
    total_sheet = readfile.sheets()
    for i in range(0, len(total_sheet)):
        sheet = total_sheet[i]
        sheet_name = readfile.sheet_names()[i]
        row = sheet.nrows
        col = sheet.ncols
        sheet_data = []
        for r in range(row):
            # print(sheet.row_values(r))
            row_data = []
            for c in range(col):
                row_data.append(sheet.cell_value(r, c))
                # print(sheet.col_values(c))
            sheet_data.append(row_data)
        file_dict[sheet_name] = sheet_data
    web_entity.excel_data = file_dict
    os.remove(path)
    return json.dumps(file_dict, ensure_ascii=False)


def write_excel(name_list):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)
    num = 0
    for name in name_list:
        sheet.write(num, 0, name)
        num += 1
    now = datetime.datetime.now()
    formatted_date = now.strftime("%Y%m%d%H%M%S")
    book.save(str(formatted_date) + "未查询到人员姓名" + ".xls")
    print("本次运行有未查询到的人员，生成文件：" + str(formatted_date) + "未查询到人员姓名" + ".xls")


def write_list(name_list):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)
    num = 0
    for item in name_list:
        for i in range(0, len(item)):
            sheet.write(num, i, item[i])
        num += 1
    now = datetime.datetime.now()
    formatted_date = now.strftime("%Y%m%d%H%M%S")
    book.save("人员名字部门查询"+str(formatted_date) + ".xls")


if __name__ == '__main__':
    name = ['123', '123']
    write_excel(name)
