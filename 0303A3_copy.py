#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2022/06/08 11:10:26
@Author     :   bokep
@Version    :   1.1.0
@Contact    :   sunson89@gmail.com
'''

# 库导入
import json
from os import getcwd
import win32com.client as VBA


# 复制工作表至首个
def ws_copy_to_first(source_sht, target_wb, ws_numbers=1):
    # #也删除同名的工作表，以防出错。
    if ws_numbers == 1:
        sht_name = source_sht.Name
        try:
            target_wb.Worksheets[sht_name].Delete()
            print("<<<<<原已有同名工作表，已进行删除！")
        except Exception:
            pass
        # source_sht.Copy(Before=target_wb.Worksheets[0])
        # target_wb.Save()
    else:
        for count in range(0, ws_numbers):
            sht_name = source_sht[count].Name
            try:
                target_wb.Worksheets[sht_name].Delete()
                print(f"<<<<<原已有同名工作表{sht_name}，已进行删除！")
            except Exception:
                pass

    source_sht.Copy(Before=target_wb.Worksheets[0])
    target_wb.Save()


# 替换链接
def link_replace(target_wb, source_fn, target_fn):
    target_wb.ChangeLink(Name=source_fn, NewName=target_fn, Type=1)
    target_wb.Save()


# 实际程序
excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

# #基础信息
root_route = getcwd()
target_route = root_route + "\\09完成文件"
mapping_route = root_route + "\\00框架文件"

json_filename = "date_data.json"
with open(json_filename, "r") as f:
    dict_data = json.load(f)
month_mark = "Y" + dict_data["CY"] + "M" + dict_data["CM"]

target_fn = target_route + "\\CombinedTB#" + month_mark + ".xlsx"
mapping_fn = mapping_route + "\\05A3.xlsx"

# #打开主文档
target_wb = excelapp.Workbooks.Open(target_fn)
# #激活被复制的工作表
mapping_wb = excelapp.Workbooks.Open(mapping_fn)
ws_numbers = mapping_wb.Worksheets.Count
arr = []
for count in range(0, ws_numbers):
    ws_name = mapping_wb.Worksheets[count].Name
    arr.append(ws_name)

# print(arr)
# arr, ws_numbers = [0], 1
try:
    mapping_ws = mapping_wb.Worksheets(arr)
except Exception:
    mapping_ws = mapping_wb.Worksheets[0]
# print(mapping_ws[0].Name)

# #复制工作表至新工作表
ws_copy_to_first(mapping_ws, target_wb, ws_numbers)

# #替换链接
source_fn = mapping_route + "\\待替换.xlsx"
link_replace(target_wb, source_fn, target_fn)

print(">>>>>已生成本期A3！")

target_wb.Close()
excelapp.Quit()
