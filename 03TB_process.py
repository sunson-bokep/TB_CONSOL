#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2021/09/04 15:48:39
@Author     :   bokep
@Version    :   0.0.1
@Contact    :   sunson89@gmail.com
'''

# 库导入
import win32com.client as VBA
from os import getcwd, listdir
import json


# 工作表合并函数
def excel_combination(excel_wb, target_wb, target_sht):
    print("<<<<<<<<<<<<<<<<<<<<<<<<<<<正在处理\"%s\"工作簿" % (excel_wb.Name))
    excel_sht_count = excel_wb.Worksheets.Count
    # print(excel_sht_count)
    for n in range(1, excel_sht_count + 1):
        excel_sht = excel_wb.Worksheets[n - 1]
        print("<<<<<<<<<<<<<<<<<<正在处理\"%s\"工作表" % (excel_sht.Name))
        limit_column_excel = excel_sht.Range("AB1048576").End(3).row
        print(limit_column_excel)
        # #对工作表进行筛选，方便后续合并
        cell_begin = excel_sht.Cells(1, "U")
        cell_end = excel_sht.Cells(limit_column_excel, "AB")
        filter_area = excel_sht.Range(cell_begin, cell_end)
        # ##需要判断是否存在筛选（不一定是第一次执行操作）。
        # print(excel_sht.AutoFilterMode)
        if excel_sht.AutoFilterMode is True:
            pass
        else:
            filter_area.AutoFilter()

        filter_criteria1 = "<>0"
        filter_area.AutoFilter(Field=8, Criteria1=filter_criteria1)
        filter_area.Copy()
        print(filter_area)

        # ##需要取到原有最大行数后一行，进行粘贴
        limit_column_target = target_sht.Range("A1048576").End(3).row + 1
        print(limit_column_target)
        cell_begin = target_sht.Cells(limit_column_target, "A")
        cell_begin.PasteSpecial(Paste=-4163)

        # target_wb.Save()

    excel_wb.Save()


# 实际程序
excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

# #基础信息
root_route = getcwd()
process_route = root_route + "\\02处理文件\\TB"
target_route = root_route + "\\02处理文件"

json_filename = "date_data.json"
with open(json_filename, "r") as f:
    dict_data = json.load(f)
month_mark = "Y" + dict_data["CY"] + "M" + dict_data["CM"]
target_fn = target_route + "\\CombinedTB" + month_mark + ".xlsx"

# #生成TB合并的新文件
target_wb = excelapp.Workbooks.Add()
target_wb.SaveAs(Filename=target_fn, FileFormat=51)
target_sht = target_wb.Worksheets[0]
target_sht.Name = "CombinedTB"
target_wb.Save()

for file in listdir(process_route):
    # print(file)
    try:
        # #目标工作簿，输入模板工作簿定义与具体执行。
        excel_fn = root_route + "\\02处理文件\\TB\\" + file
        excel_wb = excelapp.Workbooks.Open(excel_fn)
        length = file.index("#")
        formula_sn = file[:length]

        if formula_sn == "TB" or formula_sn == "ATB":
            excel_combination(excel_wb, target_sht, target_sht)
        else:
            pass

        excel_wb.Close()

    except Exception:
        raise

target_wb.Save()
target_wb.Close()

excelapp.Quit()
