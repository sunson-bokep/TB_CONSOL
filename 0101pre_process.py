#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2023/04/07 16:12:26
@Author     :   bokep
@Version    :   1.2.0
@Contact    :   sunson89@gmail.com
'''

# 库导入
import json
import datetime
from os import getcwd, listdir, path, remove, _exit
import win32com.client as VBA


# 导入程序
# #计算上一个月的标号
def prior_mark(current_input, lag_month):
    last_month_Y = int(current_input[:2])
    last_month_M = int(current_input[2:]) - int(lag_month)
    if last_month_M < 1:
        last_month_Y -= 1
        last_month_M += 12

    last_month_Y = str(last_month_Y).rjust(2, '0')
    last_month_M = str(last_month_M).rjust(2, '0')
    # #rjust可以用指定字符填充字符串至指定长度

    return(last_month_Y, last_month_M)


# 处理月份信息
current_date = datetime.datetime.today()
current_year = current_date.year
current_month = current_date.month
# current_year, current_month = 2022, 1  #跨期测试用

current_input = str(current_year)[2:] + str(current_month).rjust(2, '0')
# print(current_input)
last_month_Y, last_month_M = prior_mark(current_input, 1)
last_month_M = int(last_month_M)

# #需外部输入处理数据的年月，方便后续核验和数据处理
month_mark_input = input(f"请输入处理月份信息：\
（例：20{last_month_Y}年{last_month_M}月输入值为\
{last_month_Y}{str(last_month_M).rjust(2, '0')}）")

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
    input("请确认已更新期末审计调整...")
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
