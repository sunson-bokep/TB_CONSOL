#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Last update  :  2026/02/03 18:03:49
@Author       :  bokep
@Collaborator :  Gemini (Google AI)
@Version      :  1.1.1
@Contact      :  sunson89@gmail.com
'''

# 库导入
import threading
import time
from os import getcwd, listdir, path, remove
import win32com.client as VBA


# --- 辅助函数：心跳监测线程 ---
def heart_beat(stop_event, start_time):
    """
    独立线程：每 5 秒报告一次程序运行总耗时，不受 Excel 阻塞影响。
    """
    while not stop_event.is_set():
        time.sleep(1)  # 每秒检查一次
        elapsed = int(time.time() - start_time)
        if elapsed > 0 and elapsed % 5 == 0:
            print(f"  > [后台提示] 程序已持续运行 {elapsed}秒，请耐心等待 Excel 响应...")


# --- 主程序逻辑 ---
def main():
    # 初始化 Excel 进程
    excelapp = VBA.Dispatch("Excel.Application")
    excelapp.DisplayAlerts = False
    excelapp.Visible = False

    root_route = getcwd()
    input_route = path.join(root_route, "02处理文件", "TB")
    output_route = path.join(root_route, "02处理文件", "TB_PV")

    # 清空历史文件
    for file in listdir(output_route):
        if file.endswith(".xlsx"):
            remove(path.join(output_route, file))

    # 获取待处理列表
    files_to_process = [f for f in listdir(input_route) if f.endswith(".xlsx")]
    total_files = len(files_to_process)

    start_time = time.time()
    print(f" 开始处理，共计 {total_files} 个文件...")

    # 启动后台心跳线程
    stop_heartbeat = threading.Event()
    monitor_thread = threading.Thread(
        target=heart_beat,
        args=(stop_heartbeat, start_time)
    )
    monitor_thread.daemon = True
    monitor_thread.start()

    try:
        for index, file in enumerate(files_to_process, 1):
            input_fn = path.join(input_route, file)
            output_fn = path.join(output_route, file)

            # 打印当前任务状态
            print(f"【正在处理】({index}/{total_files}): {file}")

            # 打开并处理（此处如果是大文件可能会阻塞很久）
            wb = excelapp.Workbooks.Open(input_fn)
            wb.SaveAs(Filename=output_fn, FileFormat=51)

            # 断开链接
            connections_excel = wb.LinkSources(Type=1)
            if connections_excel is not None:
                for detail_connection in connections_excel:
                    wb.BreakLink(Name=detail_connection, Type=1)

            wb.Save()
            wb.Close()

        # 任务完成提示
        duration = int(time.time() - start_time)
        end_msg = (f"\n任务全部完成！共转换 {total_files} 个文档。"
                   f"总耗时: {duration}秒")
        print(end_msg)

    except Exception as e:
        print(f"运行中出现错误: {e}")

    finally:
        # 停止心跳线程并关闭 Excel
        stop_heartbeat.set()
        excelapp.Quit()
        print("Excel 进程已安全退出。")


if __name__ == "__main__":
    main()
