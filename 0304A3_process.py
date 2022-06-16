#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2022/06/15 11:18:36
@Author     :   bokep
@Version    :   1.0.1
@Contact    :   sunson89@gmail.com
'''

# 库导入
import win32com.client as VBA
from os import getcwd
import json


# 导入模板公式功能函数
def ws_formula_input(excel_wb, formula_fn, suffix=""):
    '''将特定模板中设置的公式导入特定目标工作表中。'''
    # #考虑通用性，VBA程序假设在引用函数前已开启；
    # #考虑通用性，目标工作簿需在函数执行前先开启，函数执行后，保持开启状态；
    # #考虑输入方便性，目标工作表在函数执行中重新定义，函数执行后进行保存；
    # #输入模板工作表在函数中打开，函数执行后直接关闭。
    # #通过输入模板工作表定位需要输入公式的具体工作表
    # ##特制化：输入需处理的工作表有标准的后缀名，如Y21M12，输入表名与输入模板名字一致
    print("<<<<<<<<<<<<<<<<<<<<<<<<<<<正在处理\"%s\"工作簿" % (excel_wb.Name))
    formula_sn = excel_wb.Name.replace(suffix, "")

    # ##先打开输入模板工作表&目标工作表，并获取最高行数。
    formula_wb = excelapp.Workbooks.Open(formula_fn)
    print("<<<<<公式位于\"%s\"工作簿" % (formula_wb.Name))
    formula_sn = str(formula_sn)                        # #不定义str处理时会报错。
    n = formula_sn.index(".")
    formula_sn = formula_sn[:n]
    # print(formula_sn)
    formula_ws = formula_wb.Worksheets[formula_sn]

    limit_column_formula = formula_ws.Cells(1, "B").value
    limit_column_formula = int(limit_column_formula)

    for i in range(4, 4 + limit_column_formula):
        # #获取输入模板中的数据
        column_index = formula_ws.Cells(i, "A").value
        column_name = formula_ws.Cells(i, "B").value
        ws_name = formula_ws.Cells(i, "C").value
        formula_id = formula_ws.Cells(i, "D").value
        input_fn = formula_ws.Cells(i, "E").value
        input_ws = formula_ws.Cells(i, "F").value

        print("<<<<<<<<<正在处理\\%s工作表\\第%d列\\%s公式" % (
            ws_name, column_index, column_name))

        excel_ws = excel_wb.Worksheets[ws_name]
        limit_column_excel = excel_ws.Range("A1048576").End(3).row

        # #将文件名和工作表名导入公式中（如适用）
        try:
            input_fn = input_fn % (suffix)  # 只有工作簿可能有后缀，工作表不会有。
        except Exception:
            pass

        try:
            formula_id = formula_id.replace("{wb}", input_fn)
            formula_id = formula_id.replace("{ws}", input_ws)
        except Exception:                        # #如果公式里不含wb/ws则保留原有不替换。
            pass

        excel_ws.Cells(1, column_index).value = column_name
        cell_begin = excel_ws.Cells(2, column_index)
        cell_end = excel_ws.Cells(limit_column_excel, column_index)
        excel_ws.Range(cell_begin, cell_end).formulaR1C1 = formula_id

    excel_wb.Save()
    print("处理完成！")
    formula_wb.Close()

    excel_wb.Save()


# 设置列宽
def columns_width(target_sht, column_width, column_begin, column_end=""):
    '''设置列宽'''
    if column_end == "":
        column_end = column_begin
    else:
        pass

    column_begin = target_sht.Columns(column_begin)
    column_end = target_sht.Columns(column_end)
    set_range = target_sht.Range(column_begin, column_end)
    set_range.ColumnWidth = column_width


# 设置格式
def columns_format(target_sht, column_format, column_begin, column_end=""):
    '''设置格式'''
    if column_end == "":
        column_end = column_begin
    else:
        pass

    column_begin = target_sht.Columns(column_begin)
    column_end = target_sht.Columns(column_end)
    set_range = target_sht.Range(column_begin, column_end)
    set_range.NumberFormatLocal = column_format


# 基础设置
excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

root_route = getcwd()
template_fn = root_route + "\\00框架文件\\04Formula-2.xlsx"

json_filename = "date_data.json"
with open(json_filename, "r") as f:
    dict_data = json.load(f)
month_mark = "Y" + dict_data["CY"] + "M" + dict_data["CM"]

target_fn = root_route + "\\09完成文件\\CombinedTB#" + month_mark + ".xlsx"
target_wb = excelapp.Workbooks.Open(target_fn)

# 导入筛选公式
ws_formula_input(target_wb, template_fn, "#" + month_mark)

# 筛选不存在的唯一识别码复制到TB
filter_ws = target_wb.Worksheets["ADJ"]
max_row_number = filter_ws.Range("A1048576").End(3).row
cell_begin = filter_ws.Cells(1, "A")
cell_end = filter_ws.Cells(max_row_number, "M")
process_area = filter_ws.Range(cell_begin, cell_end)
filter_criteria1 = "<>0"
process_area.AutoFilter(Field=8, Criteria1=filter_criteria1)
process_area.AutoFilter(Field=11, Criteria1=filter_criteria1)
filter_criteria1 = "=#N/A"
process_area.AutoFilter(Field=9, Criteria1=filter_criteria1)
target_wb.Save()

tb_ws = target_wb.Worksheets["CombinedTB"]
tb_max_row_number = tb_ws.Range("A1048576").End(3).row

# #复制公司简称
cell_begin = filter_ws.Cells(1, "A")
cell_end = filter_ws.Cells(max_row_number - 1, "A")
copy_area = filter_ws.Range(cell_begin, cell_end)
paste_begin = tb_ws.Cells(tb_max_row_number + 1, "A")
copy_area.Copy()
paste_begin.PasteSpecial(Paste=-4163)
# #复制科目及唯一识别码
cell_begin = filter_ws.Cells(1, "E")
cell_end = filter_ws.Cells(max_row_number - 1, "F")
copy_area = filter_ws.Range(cell_begin, cell_end)
paste_begin = tb_ws.Cells(tb_max_row_number + 1, "E")
copy_area.Copy()
paste_begin.PasteSpecial(Paste=-4163)
# #去除筛选
process_area.AutoFilter(Field=9)
process_area.AutoFilter(Field=11)
process_area.AutoFilter(Field=8)

target_wb.Save()
print(">>>>>>>>>>已补充调整分录涉及的唯一识别码信息。")
# 导入公式
ws_formula_input(target_wb, template_fn, "#" + month_mark)

# ##整理格式
# ###设置列宽
column_begin = "I"
column_end = "J"
column_width = 20
columns_width(tb_ws, column_width, column_begin, column_end)
# ###设置格式
column_begin = "I"
column_end = "J"
column_format = "#,##0.00_);[红色](#,##0.00)"
columns_format(tb_ws, column_format, column_begin, column_end)

target_wb.Save()
print(">>>>>>>>>>格式设置完成。")
print(">>>>>>>>>>调整过入完成。")

target_wb.Close()

excelapp.Quit()
