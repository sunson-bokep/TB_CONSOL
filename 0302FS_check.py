#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2021/09/15 17:35:24
@Author     :   bokep
@Version    :   1.0.1
@Contact    :   sunson89@gmail.com
'''

# 库导入
import json
from os import getcwd, listdir, path
import win32com.client as VBA


# 按值设置筛选
def filter_values(
        target_sht, filter_column_no, filter_value,
        column_begin, column_end=""):
    '''按值设置筛选'''
    if column_end == "":
        column_end = column_begin
    else:
        pass

    # print(target_sht.AutoFilterMode)
    if target_sht.AutoFilterMode is True:
        target_sht.Range("A1").AutoFilter()
    else:
        pass

    # column_begin = target_sht.Columns(column_begin)
    # column_end = target_sht.Columns(column_end)
    range_area = column_begin + ":" + column_end
    set_range = target_sht.Range(range_area)

    # print(filter_value)
    set_range.AutoFilter(
        Field=filter_column_no, Criteria1=filter_value, Operator=7)
    # Const xlFilterValues = 7
    # set_range.AutoFilter(Field=1, Criteria1=(
    #     "货币资金", "交易性金融资产", "开发支出", "累计折旧"), Operator=7)


# 设置条件筛选
def filter_conditions(
        target_sht, filter_column_no, filter_criteria,
        column_begin, column_end=""):
    '''设置条件筛选'''
    if column_end == "":
        column_end = column_begin
    else:
        pass

    # #特殊逻辑，为了避免上一步筛选失效，此处配置特殊逻辑
    if target_sht.AutoFilterMode is True:
        pass
    else:
        print("请确认按值筛选是否正确执行！")

    # print(filter_criteria)
    column_begin = target_sht.Columns(column_begin)
    column_end = target_sht.Columns(column_end)
    set_range = target_sht.Range(column_begin, column_end)
    criteria1 = filter_criteria[0]
    criteria2 = filter_criteria[1]
    set_range.AutoFilter(
        Field=filter_column_no,
        Criteria1=criteria1,
        Operator=1,
        Criteria2=criteria2)
    # Const xlAnd = 1


# 区域复制
def column_paste_value(
        source_sht, target_sht, copy_column,
        target_column, row_end):
    '''复制整列数据，从第二行开始复制'''
    # #设置复制区域
    cell_begin = source_sht.Cells(2, copy_column)
    cell_end = source_sht.Cells(row_end, copy_column)
    copy_range = source_sht.Range(cell_begin, cell_end)
    # print(copy_range, row_end)
    copy_range.Copy()

    # #复制到目标区域
    mark_no = target_column + "1048576"
    mark_row = target_sht.Range(mark_no).End(3).row
    cell_begin = target_sht.Cells(mark_row + 1, target_column)
    cell_begin.PasteSpecial(Paste=-4163)


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


# 导入模板公式功能函数
def excel_formula_input_specific_sht(specific_sht, formula_fn, formula_sn):
    '''将特定模板中设置的公式导入特定目标工作表中。'''
    # #考虑通用性，VBA程序假设在引用函数前已开启；
    # #考虑通用性，目标工作簿需在函数执行前先开启，函数执行后，保持开启状态；
    # #考虑输入方便性，目标工作表在函数执行中重新定义，函数执行后进行保存；
    # #输入模板工作表在函数中打开，函数执行后直接关闭。

    # ##先打开输入模板工作表&目标工作表，并获取最高行数。
    formula_wb = excelapp.Workbooks.Open(formula_fn)
    formula_sn = str(formula_sn)                        # #不定义str处理时会报错。
    formula_sht = formula_wb.Worksheets[formula_sn]

    limit_column_formula = formula_sht.Cells(1, "B").value
    limit_column_formula = int(limit_column_formula)

    # ##对每张工作表都进行处理
    excel_sht = specific_sht
    print("<<<<<<<<<<<<<<<<<<正在处理\"%s\"工作表公式" % (excel_sht.Name))
    limit_column_excel = excel_sht.Range("A1048576").End(3).row

    # ###具体导入公式操作
    for i in range(4, 4 + limit_column_formula):
        # #获取输入模板中的数据
        column_index = formula_sht.Cells(i, "A").value
        column_name = formula_sht.Cells(i, "B").value
        formula_id = formula_sht.Cells(i, "C").value
        input_fn = formula_sht.Cells(i, "D").value
        input_sht = formula_sht.Cells(i, "E").value

        # print("<<<<<<<<<正在处理第%d列%s公式" % (column_index, column_name))

        # #将文件名和工作表名导入公式中（如适用）
        try:
            formula_id = formula_id % (input_fn, input_sht)
        except Exception:                               # #如果公式里不含%s会走这条路径
            pass

        excel_sht.Cells(1, column_index).value = column_name
        cell_begin = excel_sht.Cells(2, column_index)
        cell_end = excel_sht.Cells(limit_column_excel, column_index)
        excel_sht.Range(cell_begin, cell_end).formulaR1C1 = formula_id

    print("处理完成！")

    formula_wb.Close()


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


# 实际程序
excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

# #基础信息
root_route = getcwd()
process_route = root_route + "\\01输入文件"
target_route = root_route + "\\09完成文件"
mapping_route = root_route + "\\00框架文件"

json_filename = "date_data.json"
with open(json_filename, "r") as f:
    dict_data = json.load(f)
month_mark = "Y" + dict_data["CY"] + "M" + dict_data["CM"]
target_fn = target_route + "\\CombinedTB#" + month_mark + ".xlsx"

mapping_fn = mapping_route + "\\03MR.xlsx"
mapping_wb = excelapp.Workbooks.Open(mapping_fn)
mapping_sht = mapping_wb.Worksheets["报表筛选"]

# #打开主文档
target_wb = excelapp.Workbooks.Open(target_fn)
try:            # #如果原来存在，则先删除
    target_wb.WorkSheets["FScheck"].Delete()
except Exception:
    pass
target_sht_final = target_wb.WorkSheets.Add()
target_sht_final.Name = "FScheck"
target_sht_final.Cells(1, "A").value = "公司简称"
target_sht_final.Cells(1, "B").value = "报表科目名称"
target_sht_final.Cells(1, "C").value = "报表导出金额"

target_wb.Save()

# #主体程序
for file in listdir(process_route):
    # #只有FS开头，月份信息准确，后缀为xlsx的文件属于单元需要处理的
    file_name, file_extension = path.splitext(file)
    checkpoint1 = (file_name[:2] == "FS")
    checkpoint2 = (file_name[-6:] == month_mark)
    checkpoint3 = (file_extension[1:] == "xlsx")
    # print(checkpoint1, checkpoint2, checkpoint3)
    if (checkpoint1 and checkpoint2) and checkpoint3:
        # ##满足条件的部分进行程序执行
        source_fn = process_route + "\\" + file
        source_wb = excelapp.Workbooks.Open(source_fn)

        company_name = file_name[3:-7]
        # print(company_name)

        sht_pool = [["BS", "A", 3, "A", "C"],
                    ["BS", "B", 3, "E", "G"],
                    ["PL", "C", 4, "A", "D"]]
        # ["报表","科目列","值列","筛选起始列","筛选终止列"]
        # sht_pool = [["BS", "A"]]    # 单循环测试用

        for sht in sht_pool:
            sht_name = sht[0]
            source_sht = source_wb.Worksheets[sht_name]

            # ###导入筛选数据中的值
            column_no = sht[1] + "1048576"
            max_row_no = mapping_sht.Range(column_no).End(3).row
            print(">>>>>正在处理", company_name, mapping_sht.Cells(1, sht[1]).text)
            i = 2
            filter_value = []

            while i <= max_row_no:
                add_value = mapping_sht.Cells(i, sht[1]).text
                filter_value.append(add_value)
                i += 1
            # print(filter_value)
            filter_value = tuple(filter_value)
            # print(filter_value)

            # ###设置筛选通用参数
            column_begin = sht[3]
            column_end = sht[4]        # 统括BS/PL的需求

            # ###按值进行筛选
            filter_values(
                source_sht, 1, filter_value,
                column_begin, column_end)

            # source_wb.Save()      # 测试用

            # ###筛选非空值
            filter_criteria = ["<>", "<>0"]      # 非空及非零筛选
            # print(sht[2])
            filter_conditions(
                source_sht, sht[2], filter_criteria,
                column_begin, column_end)

            # source_wb.Save()

            # ###进行复制
            target_sht = target_sht_final
            # ####资产/负债/利润分三种不同的情况
            if sht[1] == "A":
                # 资产的情况
                copy_column_pool = ["A", "C"]
            elif sht[1] == "B":
                # 负债的情况
                copy_column_pool = ["E", "G"]
            else:
                # 利润的情况
                copy_column_pool = ["A", "D"]

            row_end_mark = copy_column_pool[0]
            row_end_mark = row_end_mark + "1048576"
            # print(row_end_mark)
            row_end = source_sht.Range(row_end_mark).End(3).row
            # print(row_end)
            # ####分项目列和数字列进行处理
            for i in [0, 1]:
                copy_column = copy_column_pool[i]
                target_column = ["B", "C"][i]
                # print(sht[1], copy_column, target_column)

                column_paste_value(
                        source_sht, target_sht, copy_column,
                        target_column, row_end)

            # ###取消原有筛选
            source_sht.Range("A1").AutoFilter()
            target_wb.Save()

        # ##增加公司简称
        row_end1 = target_sht_final.Range("A1048576").End(3).row
        row_end2 = target_sht_final.Range("B1048576").End(3).row
        cell_begin = target_sht_final.Cells(row_end1 + 1, "A")
        cell_end = target_sht_final.Cells(row_end2, "A")
        fill_area = target_sht_final.Range(cell_begin, cell_end)
        fill_area.value = company_name
        target_wb.Save()

        source_wb.Close()
    else:
        pass

# #统一设置格式和公式
# ##设置列宽
columns_width(target_sht_final, 20, "A", "F")
target_wb.Save()

# ##设置格式
column_format = "#,##0.00_);[红色](#,##0.00)"
columns_format(target_sht_final, column_format, "C", "F")
target_wb.Save()

# ##导入公式
formula_fn = root_route + "\\00框架文件\\04Formula.xlsx"
formula_sn = "CombinedTB"
try:
    excel_formula_input_specific_sht(target_sht_final, formula_fn, formula_sn)
except Exception:
    print("请确认公式全部准确！")
target_wb.Save()
mapping_wb.Close()

# ##找出全部链接并切断
connections_excel = target_wb.LinkSources(Type=1)
# Const xlLinkTypeExcelLinks = 1
# print(connection_excel)
for detail_connection in connections_excel:
    target_wb.BreakLink(Name=detail_connection, Type=1)
    # Const xlExcelLinks = 1
target_wb.Save()

# ##设置筛选
columns_autofilter(target_sht_final, "A", "F")
target_wb.Save()

target_wb.Close()
excelapp.Quit()
