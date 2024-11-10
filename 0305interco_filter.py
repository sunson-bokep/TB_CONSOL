#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2022/06/27 13:39:49
@Author     :   bokep
@Version    :   1.0.0
@Contact    :   sunson89@gmail.com
'''

# 库导入
import json
from os import getcwd
import win32com.client as VBA


# 导入程序
# #筛选单独新增页面
def filter_segment(input_ws, output_ws, column_index, filter_text):
    max_row_number = input_ws.Cells(1048576, column_index).End(3).row
    cell_begin = input_ws.Cells(1, 1)
    cell_end = input_ws.Cells(max_row_number, column_index)
    process_area = input_ws.Range(cell_begin, cell_end)

    if input_ws.AutoFilterMode is True:
        process_area.AutoFilter()
    else:
        pass
    process_area.AutoFilter()

    filter_criteria1 = "=*" + filter_text + "*"
    process_area.AutoFilter(Field=column_index, Criteria1=filter_criteria1)

    max_row_number_output = output_ws.Cells(1048576, "A").End(3).row

    cell_begin = input_ws.Cells(1, column_index)
    cell_end = input_ws.Cells(max_row_number, column_index)
    first_copy_area = input_ws.Range(cell_begin, cell_end)
    first_copy_area.Copy()
    output_ws.Cells(max_row_number_output + 1, "A").PasteSpecial(Paste=12)

    # 删除第一行
    output_ws.Cells(max_row_number_output + 1, "A").EntireRow.Delete()

    process_area.AutoFilter()


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
filter_fn = root_route + "\\00框架文件\\01Co&FX.xlsx"
template_fn = root_route + "\\00框架文件\\04Formula-3.xlsx"

json_filename = "date_data.json"
with open(json_filename, "r") as f:
    dict_data = json.load(f)
month_mark = "Y" + dict_data["CY"] + "M" + dict_data["CM"]

target_fn = root_route + "\\09完成文件\\CombinedTB#" + month_mark + ".xlsx"

# 提取全部筛选信息
filter_wb = excelapp.Workbooks.Open(filter_fn)
filter_ws = filter_wb.Worksheets["Basic"]
max_row_number = filter_ws.Range("D1048576").End(3).row
filter_lis = []

for _ in range(2, max_row_number + 1):
    _ = filter_ws.Cells(_, "D").value
    filter_lis.append(_)

# print(filter_lis)

# 筛选关联方为往来明细
target_wb = excelapp.Workbooks.Open(target_fn)
input_ws = target_wb.Worksheets["CombinedTB"]
output_ws = target_wb.Worksheets["RPT"]
column_index = 6
# #先清空第一列
output_ws.Columns("A").Clear()
for _ in filter_lis:
    print(_)
    filter_segment(input_ws, output_ws, column_index, _)
    target_wb.Save()
output_ws.Cells(1, "A").value = "唯一识别码"
target_wb.Save()

# 导入公式
ws_formula_input(target_wb, template_fn, "#" + month_mark)

# #设置格式
column_begin = "L"
column_format = "#,##0.00_);[红色](#,##0.00)"
columns_format(output_ws, column_format, column_begin)
target_wb.Save()

filter_wb.Close()

# #断开链接
connections_excel = target_wb.LinkSources(Type=1)
print(connections_excel)
if connections_excel is None:
    pass
else:
    for detail_connection in connections_excel:
        target_wb.BreakLink(Name=detail_connection, Type=1)
target_wb.Save()

# #特殊逻辑-补入轧差
max_row_number = input_ws.Cells(1048576, "F").End(3).row
if input_ws.Cells(max_row_number, "F").value == "00 合并\\其他应付款\\关联方往来轧差":
    pass
else:
    input_ws.Cells(max_row_number + 1, "E").value = "其他应付款"
    input_ws.Cells(max_row_number + 1, "F").value = "00 合并\\其他应付款\\关联方往来轧差"
    input_ws.Cells(max_row_number + 1, "N").formulaR1C1 = \
        "=-SUMIF(RPT!C1,CombinedTB!RC6,RPT!C12)"
    input_ws.Cells(max_row_number + 1, "P").formulaR1C1 = \
        "=SUM(RC[-4]:RC[-1])"

max_row_number = output_ws.Cells(1048576, "A").End(3).row
output_ws.Cells(max_row_number + 1, "A").value = "00 合并\\其他应付款\\关联方往来轧差"
output_ws.Cells(max_row_number + 1, "L").formulaR1C1 = "=-SUM(R2C:R[-1]C)"
print("<<<<<特殊处理完成。")

target_wb.Save()

# #最后在CombinedTB页加筛选
input_ws.Cells(1, "F").AutoFilter()
target_wb.Save()

target_wb.Close()

excelapp.Quit()
