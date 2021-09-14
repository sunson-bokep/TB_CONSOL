#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2021/09/09 16:12:01
@Author     :   bokep
@Version    :   1.0.0
@Contact    :   sunson89@gmail.com
'''

# 库导入
from os import getcwd, listdir, path, remove
import win32com.client as VBA

# 实际程序
excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

root_route = getcwd()
input_route = root_route + "\\02处理文件\\TB"
output_route = root_route + "\\02处理文件\\TB_PV"

# #先把存放处理文件中的xlsx文件全部清空（防止有历史文件留存）
for file in listdir(output_route):
    if file[-4:] == "xlsx":
        output_fn = path.join(output_route, file)
        remove(output_fn)
    else:
        pass

i = 0

for file in listdir(input_route):
    determine1 = (file[-4:] == "xlsx")

    if determine1:
        # 将满足条件的文件存入TB_PV文件夹
        input_fn = path.join(input_route, file)
        wb = excelapp.Workbooks.Open(input_fn)
        output_fn = path.join(output_route, file)
        wb.SaveAs(Filename=output_fn, FileFormat=51)

        # 找出全部链接并切断
        connections_excel = wb.LinkSources(Type=1)
        # Const xlLinkTypeExcelLinks = 1
        # print(connection_excel)
        for detail_connection in connections_excel:
            wb.BreakLink(Name=detail_connection, Type=1)
            # Const xlExcelLinks = 1

        wb.Save()
        wb.Close()
        i += 1
    else:
        pass

if i == 0:
    pass
else:
    print("<<<<<<<<<当月全部xlsx文档已生成无公式链接版本。")

excelapp.Quit()
