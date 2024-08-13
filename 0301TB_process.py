#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2023/04/07 16:04:15
@Author     :   bokep
@Version    :   1.1.4
@Contact    :   sunson89@gmail.com
'''

# 库导入
import win32com.client as VBA
from os import getcwd, listdir
import json


# 工作表合并函数
def excel_combination(excel_wb, target_sht):
    print("<<<<<<<<<<<<<<<<<<<<<<<<<<<正在处理\"%s\"工作簿" % (excel_wb.Name))
    excel_sht_count = excel_wb.Worksheets.Count
    # print(excel_sht_count)
    for n in range(1, excel_sht_count + 1):
        excel_sht = excel_wb.Worksheets[n - 1]
        print("<<<<<<<<<<<<<<<<<<正在处理\"%s\"工作表" % (excel_sht.Name))
        limit_column_excel = excel_sht.Range("AB1048576").End(3).row
        # print(limit_column_excel)
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
        # ###部分抵消项需要剔除，报表科目匹配结果为0
        filter_area.AutoFilter(Field=4, Criteria1=filter_criteria1)
        filter_area.Copy()
        # print(filter_area)

        # ##需要取到原有最大行数后一行，进行粘贴
        limit_column_target = target_sht.Range("A1048576").End(3).row + 1
        # print(limit_column_target)
        cell_begin = target_sht.Cells(limit_column_target, "A")
        cell_begin.PasteSpecial(Paste=-4163)
        # target_wb.Save()

    excel_wb.Save()
    excel_wb.Close()


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


# 设置筛选
def columns_autofilter(target_sht, column_begin, column_end=""):
    '''设置筛选'''
    if column_end == "":
        column_end = column_begin
    else:
        pass

    if target_sht.AutoFilterMode is True:
        target_sht.Range("A1").AutoFilter()
    else:
        pass

    column_begin = target_sht.Columns(column_begin)
    column_end = target_sht.Columns(column_end)
    set_range = target_sht.Range(column_begin, column_end)
    set_range.AutoFilter()


# 设置排序
def columns_sort(target_sht, column_number, sortorder=1):
    '''设置排序，默认是按升序排列'''
    # Const xlAscending = 1
    # Const xlDescending = 2
    # Const xlSortOnValues = 0
    # Const xlSortNormal = 0
    # Const xlYes = 1
    # Const xlTopToBottom = 1
    # Const xlPinYin = 1

    # #先需要确认排序范围
    cell_begin = target_sht.Cells(1, column_number)
    cell_end = target_sht.Cells(1048576, column_number)
    limit_column = target_sht.Range(cell_begin, cell_end).End(3).row
    cell_end = target_sht.Cells(limit_column, column_number)
    key_range = target_sht.Range(cell_begin, cell_end)

    target_sht.AutoFilter.Sort.SortFields.Clear()
    target_sht.AutoFilter.Sort.SortFields.Add2(
        Key=key_range, SortOn=0, Order=sortorder, DataOption=0)

    target_sht.AutoFilter.Sort.Header = 1
    target_sht.AutoFilter.Sort.MatchCase = False
    target_sht.AutoFilter.Sort.Orientation = 1
    target_sht.AutoFilter.Sort.SortMethod = 1
    target_sht.AutoFilter.Sort.Apply()


# 实际程序
excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

# #基础信息
root_route = getcwd()
process_route = root_route + "\\02处理文件\\TB_PV"
target_route = root_route + "\\09完成文件"

json_filename = "date_data.json"
with open(json_filename, "r") as f:
    dict_data = json.load(f)
month_mark = "Y" + dict_data["CY"] + "M" + dict_data["CM"]
target_fn = target_route + "\\CombinedTB#" + month_mark + ".xlsx"

# #生成TB合并的新文件
target_wb = excelapp.Workbooks.Add()
target_wb.SaveAs(Filename=target_fn, FileFormat=51)
target_sht = target_wb.Worksheets[0]
target_sht.Name = "CombinedTB"
target_ws2 = target_wb.Worksheets.Add()
target_ws2.Name = "GRIR"
target_wb.Save()

for file in listdir(process_route):
    # print(file)
    try:
        # #目标工作簿，输入模板工作簿定义与具体执行。
        excel_fn = process_route + "\\" + file
        excel_wb = excelapp.Workbooks.Open(excel_fn)
        length = file.index("#")
        formula_sn = file[:length]

        if formula_sn == "TB" or formula_sn == "ATB":
            excel_combination(excel_wb, target_sht)
            target_wb.Save()
        elif formula_sn == "GRIR":
            excel_combination(excel_wb, target_ws2)
        else:
            excel_wb.Close()  # 如果不是则对文件进行直接关闭，否则后续冻结操作会报错

    except Exception:
        # raise
        pass

# #处理合并工作表
# ##删除空白首行
target_sht.Cells(1, "A").EntireRow.Delete()
limit_column_target = target_sht.Range("A1048576").End(3).row
# print(limit_column_target)
cell_begin = target_sht.Cells(1, "A")
cell_end = target_sht.Cells(limit_column_target, "H")
filter_area = target_sht.Range(cell_begin, cell_end)
if limit_column_target > 1:
    filter_area.AutoFilter()
    filter_criteria1 = "RMB借正贷负"
    filter_area.AutoFilter(Field=8, Criteria1=filter_criteria1)
    # ##删除重复抬头
    cell_begin = target_sht.Cells(2, "A")
    cell_end = target_sht.Cells(limit_column_target, "H")
    filter_area = target_sht.Range(cell_begin, cell_end)
    filter_area.EntireRow.Delete()
    filter_area.AutoFilter()
    target_wb.Save()
    # ##整理格式
    # ###设置列宽
    column_begin = "A"
    column_end = "H"
    column_width = 20
    columns_width(target_sht, column_width, column_begin, column_end)
    column_begin = "F"
    column_width = 90
    columns_width(target_sht, column_width, column_begin)
    target_wb.Save()
    # ###设置格式
    column_begin = "G"
    column_end = "H"
    column_format = "#,##0.00_);[红色](#,##0.00)"
    columns_format(target_sht, column_format, column_begin, column_end)
    target_wb.Save()
    # ###设置筛选
    column_begin = "A"
    column_end = "H"
    columns_autofilter(target_sht, column_begin, column_end)
    target_wb.Save()
    # ###设置冻结
    target_sht.Cells(2, "F").Select()
    excelapp.ActiveWindow.FreezePanes = True
    target_wb.Save()
    # ###设置排序/先按公司后按科目排序（程序倒序）
    columns_sort(target_sht, 2)
    columns_sort(target_sht, 1)
    target_wb.Save()
    # ###设置合并
    column_begin = target_sht.Columns("B")
    column_end = target_sht.Columns("C")
    target_sht.Range(column_begin, column_end).Group()
    target_wb.Save()

else:
    pass    # #如果无数据，行数统计为1，会自动跳过上述处理。
# input("pause")  #用于找错
# #处理GRIR工作表
# ##删除空白首行
target_ws2.Cells(1, "A").EntireRow.Delete()
limit_column_target = target_ws2.Range("A1048576").End(3).row
# print(limit_column_target)
cell_begin = target_ws2.Cells(1, "A")
target_ws2.Cells(1, "I").value = "筛选列"
cell_end = target_ws2.Cells(limit_column_target, "I")
filter_area = target_ws2.Range(cell_begin, cell_end)
if limit_column_target > 5:
    # 根据测试，如果有1条数据（则必另有合计行），则数值为6,
    # 根据上述规律，数值小于6时，说明数据集为空，可以跳过
    print(limit_column_target)
    filter_area.AutoFilter()
    filter_criteria1 = "RMB借正贷负"
    filter_area.AutoFilter(Field=8, Criteria1=filter_criteria1)
    # ##删除重复抬头
    # ###先统计筛选后剩余行数（如果只有一家有GRIR，理论上筛选后剩余1行，直接删除会出错）
    zero_line_check = target_ws2.Range("A1048576").End(3).row
    # print(zero_line_check)
    if zero_line_check == 1:
        pass
    else:
        cell_begin = target_ws2.Cells(2, "A")
        cell_end = target_ws2.Cells(limit_column_target, "I")
        filter_area = target_ws2.Range(cell_begin, cell_end)
        filter_area.EntireRow.Delete()

    filter_area.AutoFilter()
    target_wb.Save()
    # ##整理格式
    # ###删除多余列
    target_ws2.Columns("G").Delete()
    target_ws2.Columns("B:E").Delete()
    # ###设置列宽
    column_begin = "A"
    column_end = "C"
    column_width = 20
    columns_width(target_ws2, column_width, column_begin, column_end)
    column_begin = "B"
    column_width = 90
    columns_width(target_ws2, column_width, column_begin)
    target_wb.Save()
    # ###设置格式
    column_begin = "C"
    column_end = "C"
    column_format = "#,##0.00_);[红色](#,##0.00)"
    columns_format(target_ws2, column_format, column_begin, column_end)
    target_wb.Save()
    # ###设置筛选
    column_begin = "A"
    column_end = "C"
    columns_autofilter(target_ws2, column_begin, column_end)
    target_wb.Save()
    # ###设置冻结
    target_ws2.Activate()
    target_ws2.Cells(2, "A").Select()
    excelapp.ActiveWindow.FreezePanes = True
    target_wb.Save()
    # ###设置排序/先按公司后按科目排序（程序倒序）
    columns_sort(target_ws2, 2)
    columns_sort(target_ws2, 1)
    target_wb.Save()
else:
    pass    # #如果数值偏小，会自动跳过上述处理。

target_wb.Close()

excelapp.Quit()
