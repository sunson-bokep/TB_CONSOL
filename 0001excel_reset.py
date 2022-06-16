#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2022/01/18 15:40:49
@Author     :   bokep
@Version    :   2.0.1
@Contact    :   sunson89@gmail.com
'''

# 库导入
import win32com.client as VBA
import os

# 终止一般情况下的正常excel程序
excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

excelapp.Quit()

# 做检查，如仍有excel程序流程，进行强制关闭
i = 100
while i > 0:
    f = "EXCEL.EXE" in os.popen('tasklist /FI "IMAGENAME eq excel.exe"').read()
    # print(f)
    if f:
        os.system('TASKKILL /F /IM excel.exe')
        # /F：指定强制终止进程，
        # /IM：指定要终止的进程的映像名称，通配符 '*'可用来指定所有任务或映像名称。
        print(f">>>>>已尝试强制后台关闭excel程序{101 - i}次。")
        i -= 1
    else:
        i = 0
        print(">>>>>后台已无excel程序。")
