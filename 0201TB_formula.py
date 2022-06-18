#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2021/09/16 16:30:13
@Author     :   bokep
@Version    :   1.1.3
@Contact    :   sunson89@gmail.com
'''

# 库导入
import win32com.client as VBA
from os import getcwd, listdir


# 导入模板公式功能函数
def excel_formula_input(excel_wb, formula_fn, formula_sn):
    '''将特定模板中设置的公式导入特定目标工作表中。'''
    # #考虑通用性，VBA程序假设在引用函数前已开启；
    # #考虑通用性，目标工作簿需在函数执行前先开启，函数执行后，保持开启状态；
    # #考虑输入方便性，目标工作表在函数执行中重新定义，函数执行后进行保存；
    # #输入模板工作表在函数中打开，函数执行后直接关闭。

    print("<<<<<<<<<<<<<<<<<<<<<<<<<<<正在处理\"%s\"工作簿" % (excel_wb.Name))

    # ##先打开输入模板工作表&目标工作表，并获取最高行数。
    formula_wb = excelapp.Workbooks.Open(formula_fn)
    formula_sn = str(formula_sn)                        # #不定义str处理时会报错。
    formula_sht = formula_wb.Worksheets[formula_sn]

    limit_column_formula = formula_sht.Cells(1, "B").value
    limit_column_formula = int(limit_column_formula)

    excel_sht_count = excel_wb.Worksheets.Count

    # ##对每张工作表都进行处理
    for n in range(1, excel_sht_count + 1):
        excel_sht = excel_wb.Worksheets[n - 1]
        print("<<<<<<<<<<<<<<<<<<正在处理\"%s\"工作表" % (excel_sht.Name))
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

        excel_wb.Save()
        print("处理完成！")

    excel_wb.Save()
    formula_wb.Close()


# 实际程序
excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

root_route = getcwd()
process_route = root_route + "\\02处理文件\\TB"

# #打开需要使用到的工作表，并在使用结束后关闭。
input_fn1 = root_route + "\\00框架文件\\01Co&FX.xlsx"
input_fn2 = root_route + "\\00框架文件\\02SR.xlsx"
input_fn3 = root_route + "\\00框架文件\\03MR.xlsx"
input_fn1_wb = excelapp.Workbooks.Open(input_fn1)
input_fn2_wb = excelapp.Workbooks.Open(input_fn2)
input_fn3_wb = excelapp.Workbooks.Open(input_fn3)

formula_fn = root_route + "\\00框架文件\\04Formula.xlsx"

# #对02文件夹中的文件都进行处理
for file in listdir(process_route):
    # print(file)
    try:    # #文件夹下有非xlsx的文件，用try可以避免出错。
        # #目标工作簿，输入模板工作簿定义与具体执行。
        excel_fn = root_route + "\\02处理文件\\TB\\" + file
        excel_wb = excelapp.Workbooks.Open(excel_fn)
        length = file.index("#")
        formula_sn = file[:length]
        if formula_sn == "TB" or formula_sn == "ATB":

            try:
                excel_formula_input(excel_wb, formula_fn, formula_sn)
            except Exception:
                print("请确认公式全部准确！")

            # #为了计算非人民币公司的外币报表折算差额，单独写特例。
            determine1 = (formula_sn == "TB")
            # print(determine1)

            excel_sht = excel_wb.Worksheets[0]
            currency_mark = excel_sht.Cells(10, "AI").value
            # print(currency_mark)
            determine2 = (currency_mark != "RMB")
            # print(determine2)

            if determine1 and determine2:
                print("<<<<<<<<<对外币报表折算金额进行处理")
                limit_column_excel = excel_sht.Range("A1048576").End(3).row
                fa = "=R[-2]C"
                cell = excel_sht.Cells(limit_column_excel, "U")
                excel_sht.Range(cell, cell)\
                    .formulaR1C1 = fa

                fa = "=TEXT(9999,0)"
                cell = excel_sht.Cells(limit_column_excel, "V")
                excel_sht.Range(cell, cell)\
                    .formulaR1C1 = fa

                fa = "=RC[-5]&\"\\\"&RC[-4]&\"\\\"&RC[-3]"
                cell = excel_sht.Cells(limit_column_excel, "Z")
                excel_sht.Range(cell, cell)\
                    .formulaR1C1 = fa

                fa = "=ROUND(-SUM(R2C:R[-1]C),2)"
                cell = excel_sht.Cells(limit_column_excel, "AO")
                excel_sht.Range(cell, cell)\
                    .formulaR1C1 = fa

                # #可以直接取数
                # fa = "=-SUMIF(C[4],\"本年利润抵消明细\",C)"
                # cell = excel_sht.Cells(limit_column_excel, "AA")
                # excel_sht.Range(cell, cell)\
                #     .formulaR1C1 = fa

                excel_wb.Save()
                print("处理完成！")
            else:
                pass
        elif formula_sn == "GRIR":  # 对应付暂估表格进行处理
            try:
                excel_formula_input(excel_wb, formula_fn, formula_sn)
            except Exception:
                print("请确认公式全部准确！")

            # 修改最后一行的公式
            excel_sht_count = excel_wb.Worksheets.Count
            # ##对每张工作表都进行处理
            for n in range(1, excel_sht_count + 1):
                excel_sht = excel_wb.Worksheets[n - 1]
                limit_column_excel = excel_sht.Range("A1048576").End(3).row - 2
                # 倒数第三行是合计列
                # ###修改1
                excel_sht.Cells(limit_column_excel, "AG").value = \
                    "\\220202\\应付账款\\应付账款-暂估"
                # ###修改2
                excel_sht.Cells(limit_column_excel, "AB").formulaR1C1 = \
                    "=-SUM(R1C:R[-1]C)"
                # ###修改3
                excel_sht.Cells(limit_column_excel, "A").value = \
                    ""
            print("<<<<<<<<<<已修改所有页面最后行内容。")
            excel_wb.Save()
        else:
            pass

        excel_wb.Close()

    except Exception:
        pass

input_fn1_wb.Close()
input_fn2_wb.Close()
input_fn3_wb.Close()

excelapp.Quit()
