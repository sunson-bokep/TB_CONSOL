#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update:   2021/08/27 12:11:31
@Author     :   bokep
@Version    :   1.1
@Contact    :   sunson89@gmail.com
'''

# 库导入
import win32com.client as VBA

"""在主程序执行前，关闭全部Excel文档，避免后续妨碍程序运行"""

excelapp = VBA.Dispatch("Excel.Application")

excelapp.DisplayAlerts = False
excelapp.Visible = False

excelapp.Quit()
