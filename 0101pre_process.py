#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2021/09/09 13:51:43
@Author     :   bokep
@Version    :   1.1.0
@Contact    :   sunson89@gmail.com
'''

# 库导入
import json
from os import getcwd, listdir, path, remove, _exit
import win32com.client as VBA

# #需外部输入处理数据的年月，方便后续核验和数据处理
month_mark_input = input("请输入处理月份信息：（例：2021年6月输入值为2106）")
month_mark = "Y" + month_mark_input[:2] + "M" + month_mark_input[2:]
json_filename = "date_data.json"

dict = {}
dict["CY"] = month_mark_input[:2]
dict["CM"] = month_mark_input[2:]

with open(json_filename, "w") as f:
    # dict = json.dumps(dict, sort_keys=True, indent=4, separators=(',', ': '))
    json.dump(dict, f)


# #对输入文件夹中的xls文件进行重新保存，将格式保存为xlsx。
# ##先激活win32com.client，以供后续操作使用。
excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

# ##对目标文件夹中的xls文件进行具体操作。
root_route = getcwd()
input_route = root_route + "\\01输入文件"
output_route = root_route + "\\02处理文件\\TB"

# #先把存放处理文件的文件夹中的xlsx文件全部清空（防止有历史文件留存）
for file in listdir(output_route):
    if file[-4:] == "xlsx":
        output_fn = path.join(output_route, file)
        remove(output_fn)
    else:
        pass

if len(month_mark_input) == 4:
    input("请确认已更新框架文件中的当月汇率...")
else:
    print("请确认输入的月份信息是否准确！")
    excelapp.Quit()
    _exit(0)

i = 0

for file in listdir(input_route):
    # print(file[-3:])
    determine1 = (file[-3:] == "xls")
    determine2 = (file[-10:-4] == month_mark)
    # print(determine2)

    if determine1 and determine2:
        # 生成完整文件路径
        input_fn = path.join(input_route, file)
        # print(input_fn)
        # 对xls文档进行另存为xlsx处理
        xls_wb = excelapp.Workbooks.Open(input_fn)

        file = file[:-3] + "xlsx"
        print(file)
        output_fn = path.join(output_route, file)
        xls_wb.SaveAs(Filename=output_fn, FileFormat=51)
        # Const xlOpenXMLWorkbook = 51 (&H33)
        xls_wb.Close()
        i += 1          # #对处理文件数量进行统计
    else:
        pass

if i == 0:
    input(">>>>>请确认输入文件及输入日期是否匹配？")
else:
    print("<<<<<<<<<全部当月xls文档已另存为成xlsx文档。")

excelapp.Quit()
