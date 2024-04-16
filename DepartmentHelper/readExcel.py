import datetime
import os
import xlrd


# import xlwt


def readXlsFile(path):
    try:
        wb = xlrd.open_workbook(path)
        s_names = wb.sheet_names()
        wb_dict = {}
        for item in s_names:
            sheet = wb.sheet_by_name(item)
            wb_dict[item] = {}
            row_num = sheet.nrows  # 行数
            col_num = sheet.ncols  # 列数
            if col_num == 0 and row_num == 0:
                wb_dict[item] = []
            else:
                for i in range(row_num):
                    wb_dict[item][str(i)] = []
                    for j in range(col_num):
                        wb_dict[item][str(i)].append(sheet.cell_value(i, j))  # cell(行，列)
        # if len(wb_dict.keys()) != 1:
        #     print("文件格式有误...请检查")
        #     return False
        # print(wb_dict)
        res = []

        for sheet in wb_dict.keys():
            depDict = {"depName": [], "parentName": [], "changeDepName": [], "changeDepCode": [], "prefixCode": "",
                       "organization": str(sheet)}
            # {"depName": [], "parentName": [], "changeDepName": [], "changeDepCode": []}
            for key in wb_dict[sheet].keys():
                for index, item in enumerate(wb_dict[sheet][key]):
                    if index != 4 and (item == '' or type(item) != str):
                        print("文件格式有误...请检查")
                        return False
                if key != '0':
                    depDict["depName"].append(wb_dict[sheet][key][0])
                    depDict["parentName"].append(wb_dict[sheet][key][1])
                    depDict["changeDepName"].append(wb_dict[sheet][key][2])
                    depDict["changeDepCode"].append(wb_dict[sheet][key][3])
                if key == '1':
                    depDict["prefixCode"] = str(wb_dict[sheet][key][4])
            depDict["count"] = len(depDict["depName"])
            res.append(depDict)
        print("文件格式校验通过")
        # for i in res:
        #     print(i)
        return res
    except BaseException as e:
        print(e)
        return False


def readXlsFile4YG(path):
    try:
        wb = xlrd.open_workbook(path)
        s_names = wb.sheet_names()
        wb_dict = {}
        if len(s_names) != 1:
            print("只允许有一个工作簿")
            return False
        sheet = wb.sheet_by_name(s_names[0])
        wb_dict[s_names[0]] = {}
        row_num = sheet.nrows  # 行数
        col_num = sheet.ncols  # 列数
        if col_num == 0 and row_num == 0:
            wb_dict[s_names[0]] = []
        else:
            for i in range(col_num):
                wb_dict[s_names[0]][str(i)] = []
                for j in range(row_num):
                    wb_dict[s_names[0]][str(i)].append(sheet.cell_value(j, i))  # cell(行，列)
        # print(wb_dict)
        if len(wb_dict.keys()) != 1:
            print("文件格式有误...请检查")
            return False
        res = []
        for key in wb_dict["Sheet1"].keys():
            depDict = {"processName": "", "depName": [], "type": ""}
            print(wb_dict["Sheet1"][key])
            for index, item in enumerate(wb_dict["Sheet1"][key]):
                if item == '' and (index == 0 or index == 1):
                    print("文件格式有误...请检查")
                    return False
                if index == 0:
                    depDict["processName"] = item
                elif index == 1:
                    depDict["type"] = item
                else:
                    if item != '':
                        depDict["depName"].append(item)
            depDict["count"] = len(depDict["depName"])
            res.append(depDict)
        print("文件格式校验通过，请核对一下是否有误")
        print(res)
        print("共计流程：%d个" % len(res))
        return res
    except BaseException as e:
        print(e)
        return False


def delete_files(folder_path):
    files = os.listdir(folder_path)
    for file in files:
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path):
            os.remove(file_path)


# def write_excel(wb_dict):
#     # {[{},{}]}
#     if len(wb_dict.keys()) == 0:
#         return "error"
#     else:
#         book = xlwt.Workbook(encoding='utf-8')
#         for item in wb_dict.keys():
#             sheet_data = wb_dict[item]
#             sheet_name = item
#             sheet = book.add_sheet(sheet_name, cell_overwrite_ok=True)
#             header_style = xlwt.XFStyle()
#             font = xlwt.Font()
#             # 字体名称
#             font.name = '宋体'
#             # 字体大小（20是基准单位，18表示18px）
#             font.height = 15 * 15
#             # 是否使用粗体
#             font.bold = False
#             # 是否使用斜体
#             font.italic = False
#             # 字体颜色
#             font.colour_index = 0
#             header_style.font = font
#             for item1 in sheet_data.keys():
#                 row_data = sheet_data[item1]
#                 row = int(item1)
#                 for col in range(len(row_data)):
#                     sheet.write(row, col, row_data[col], header_style)  # 行，列，内容
#
#         now = datetime.datetime.now()
#         formatted_date = now.strftime("%Y-%m-%d %H:%M:%S")
#         file_name = formatted_date + "运行结果" + ".xls"
#         # file_path = "C:/Users/Administrator/Desktop/bot-demo/public/static/excelFile/"
#         book.save("./" + file_name)


if __name__ == '__main__':
    # readXlsFile(input())
    readXlsFile4YG(input())
