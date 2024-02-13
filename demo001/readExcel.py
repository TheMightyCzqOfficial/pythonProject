import os

import xlrd
import xlwt
import datetime
from configController import JsonCon

front_path = JsonCon().config["front_public_path"]


def read(path):
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
    print(wb_dict)
    return wb_dict


def write_excel(wb_dict):
    # {[{},{}]}
    if len(wb_dict.keys()) == 0:
        return "error"
    else:
        book = xlwt.Workbook(encoding='utf-8')
        for item in wb_dict.keys():
            sheet_data = wb_dict[item]
            sheet_name = item
            sheet = book.add_sheet(sheet_name, cell_overwrite_ok=True)
            header_style = xlwt.XFStyle()
            font = xlwt.Font()
            # 字体名称
            font.name = '宋体'
            # 字体大小（20是基准单位，18表示18px）
            font.height = 15 * 15
            # 是否使用粗体
            font.bold = False
            # 是否使用斜体
            font.italic = False
            # 字体颜色
            font.colour_index = 0
            header_style.font = font
            for item1 in sheet_data.keys():
                row_data = sheet_data[item1]
                row = int(item1)
                for col in range(len(row_data)):
                    sheet.write(row, col, row_data[col], header_style)  # 行，列，内容

        now = datetime.datetime.now()
        formatted_date = now.strftime("%Y%m%d%H%M%S")
        file_name = "File" + formatted_date + ".xls"
        # file_path = "C:/Users/Administrator/Desktop/bot-demo/public/static/excelFile/"
        book.save(front_path + file_name)


def delete_files(folder_path):
    files = os.listdir(folder_path)
    for file in files:
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path):
            os.remove(file_path)


def write4download(step_list):
    # {[{},{}]}
    if len(step_list) == 0:
        return "error"
    else:
        book = xlwt.Workbook(encoding='utf-8')
        sheet = book.add_sheet("Sheet1", cell_overwrite_ok=True)
        # 在此插入首行单元格格式修改代码
        # (默认情况下该单元格没有数据才能修改）
        # 如需调整背景颜色,需注释掉原始表头数据段：
        # 格式设置代码段开始
        header_style = xlwt.XFStyle()
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # # 0 - 黑色、1 - 白色、2 - 红色、3 - 绿色、4 - 蓝色、5 - 黄色、6 - 粉色、7 - 青色
        pattern.pattern_fore_colour = 5
        header_style.pattern = pattern
        font = xlwt.Font()
        # 字体名称
        font.name = '宋体'
        # 字体大小（20是基准单位，18表示18px）
        font.height = 15 * 15
        # 是否使用粗体
        font.bold = False
        # 是否使用斜体
        font.italic = False
        # 字体颜色
        font.colour_index = 0
        header_style.font = font
        for index, title in enumerate(step_list):
            sheet.col(index).width = 256 * 20
            sheet.write(0, index, title, header_style)
        now = datetime.datetime.now()
        formatted_date = now.strftime("%Y%m%d%H%M%S")
        file_name = "download" + formatted_date + ".xls"
        delete_files(front_path)
        book.save(front_path + file_name)
        return file_name


if __name__ == '__main__':
    # write_excel(read('excel/test.xls'))
    write4download(['step1', 'step1', 'step1', 'step1', 'step1'])
