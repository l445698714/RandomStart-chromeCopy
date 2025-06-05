#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
输入工具模块
包含随机数字输入和文本文件输入功能
"""

import random
import time
import os
import win32gui
import win32con
import win32api
import ctypes
from typing import List, Union
import logging


def input_random_number(window_handles: List[int], min_val: Union[int, float], max_val: Union[int, float], 
                       is_float: bool = False, decimal_places: int = 2, 
                       overwrite: bool = True, delayed: bool = False) -> bool:
    """
    向指定窗口输入随机数字
    
    Args:
        window_handles: 窗口句柄列表
        min_val: 最小值
        max_val: 最大值
        is_float: 是否生成浮点数
        decimal_places: 小数位数
        overwrite: 是否覆盖原有内容
        delayed: 是否模拟人工输入（逐字输入并添加延迟）
        
    Returns:
        bool: 操作是否成功
    """
    try:
        if not window_handles:
            return False
            
        if min_val > max_val:
            min_val, max_val = max_val, min_val
            
        success_count = 0
        
        for hwnd in window_handles:
            try:
                # 检查窗口是否有效
                if not win32gui.IsWindow(hwnd):
                    continue
                    
                # 生成随机数
                if is_float:
                    value = random.uniform(min_val, max_val)
                    text = f"{value:.{decimal_places}f}"
                else:
                    value = random.randint(int(min_val), int(max_val))
                    text = str(value)
                
                # 激活窗口
                try:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.1)
                except:
                    pass
                
                # 如果需要覆盖原有内容，先全选
                if overwrite:
                    # 发送 Ctrl+A 全选
                    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                    win32api.keybd_event(ord('A'), 0, 0, 0)
                    win32api.keybd_event(ord('A'), 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
                    time.sleep(0.05)
                
                # 输入文本
                if delayed:
                    # 模拟人工输入，逐字输入
                    for char in text:
                        _send_char(char)
                        time.sleep(random.uniform(0.05, 0.15))
                else:
                    # 快速输入
                    _send_text(text)
                
                success_count += 1
                time.sleep(0.1)  # 窗口间延迟
                
            except Exception as e:
                logging.error(f"向窗口 {hwnd} 输入随机数失败: {e}")
                continue
        
        return success_count > 0
        
    except Exception as e:
        logging.error(f"随机数字输入失败: {e}")
        return False


def input_text_from_file(window_handles: List[int], file_path: str, 
                        input_method: str = "sequential", overwrite: bool = True, 
                        delayed: bool = False) -> bool:
    """
    从文件读取文本并输入到指定窗口
    
    Args:
        window_handles: 窗口句柄列表
        file_path: 文本文件路径
        input_method: 输入方式 ("sequential": 顺序, "random": 随机)
        overwrite: 是否覆盖原有内容
        delayed: 是否模拟人工输入（逐字输入并添加延迟）
        
    Returns:
        bool: 操作是否成功
    """
    try:
        if not window_handles:
            return False
            
        if not os.path.exists(file_path):
            logging.error(f"文件不存在: {file_path}")
            return False
        
        # 读取文件内容
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f.readlines() if line.strip()]
        except UnicodeDecodeError:
            try:
                with open(file_path, 'r', encoding='gbk') as f:
                    lines = [line.strip() for line in f.readlines() if line.strip()]
            except Exception as e:
                logging.error(f"读取文件失败: {e}")
                return False
        except Exception as e:
            logging.error(f"读取文件失败: {e}")
            return False
            
        if not lines:
            logging.error("文件内容为空")
            return False
        
        success_count = 0
        
        for i, hwnd in enumerate(window_handles):
            try:
                # 检查窗口是否有效
                if not win32gui.IsWindow(hwnd):
                    continue
                
                # 选择要输入的文本
                if input_method == "random":
                    text = random.choice(lines)
                else:  # sequential
                    text = lines[i % len(lines)]
                
                # 激活窗口
                try:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.1)
                except:
                    pass
                
                # 如果需要覆盖原有内容，先全选
                if overwrite:
                    # 发送 Ctrl+A 全选
                    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                    win32api.keybd_event(ord('A'), 0, 0, 0)
                    win32api.keybd_event(ord('A'), 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
                    time.sleep(0.05)
                
                # 输入文本
                if delayed:
                    # 模拟人工输入，逐字输入
                    for char in text:
                        _send_char(char)
                        time.sleep(random.uniform(0.05, 0.15))
                else:
                    # 快速输入
                    _send_text(text)
                
                success_count += 1
                time.sleep(0.1)  # 窗口间延迟
                
            except Exception as e:
                logging.error(f"向窗口 {hwnd} 输入文本失败: {e}")
                continue
        
        return success_count > 0
        
    except Exception as e:
        logging.error(f"文本文件输入失败: {e}")
        return False


def _send_text(text: str):
    """
    快速发送文本到当前活动窗口
    """
    try:
        # 使用剪贴板方法发送文本（适用于Unicode字符）
        import pyperclip
        
        # 保存当前剪贴板内容
        original_clipboard = ""
        try:
            original_clipboard = pyperclip.paste()
        except:
            pass
        
        # 设置剪贴板内容
        pyperclip.copy(text)
        time.sleep(0.05)
        
        # 发送 Ctrl+V 粘贴
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        time.sleep(0.1)
        
        # 恢复原始剪贴板内容
        try:
            pyperclip.copy(original_clipboard)
        except:
            pass
            
    except ImportError:
        # 如果没有pyperclip，使用逐字符输入
        for char in text:
            _send_char(char)
            time.sleep(0.01)
    except Exception as e:
        logging.error(f"发送文本失败: {e}")


def _send_char(char: str):
    """
    发送单个字符到当前活动窗口
    """
    try:
        # 对于ASCII字符，使用键盘事件
        if ord(char) < 128:
            if char.isalpha():
                vk_code = ord(char.upper())
                if char.islower():
                    # 小写字母
                    win32api.keybd_event(vk_code, 0, 0, 0)
                    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                else:
                    # 大写字母
                    win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
                    win32api.keybd_event(vk_code, 0, 0, 0)
                    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
            elif char.isdigit():
                vk_code = ord(char)
                win32api.keybd_event(vk_code, 0, 0, 0)
                win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            else:
                # 特殊字符
                _send_special_char(char)
        else:
            # Unicode字符，使用Unicode输入
            _create_unicode_input(char)
            
    except Exception as e:
        logging.error(f"发送字符 '{char}' 失败: {e}")


def _send_special_char(char: str):
    """
    发送特殊字符
    """
    special_chars = {
        ' ': win32con.VK_SPACE,
        '.': win32con.VK_OEM_PERIOD,
        ',': win32con.VK_OEM_COMMA,
        '-': win32con.VK_OEM_MINUS,
        '=': win32con.VK_OEM_PLUS,
        ';': win32con.VK_OEM_1,
        '/': win32con.VK_OEM_2,
        '`': win32con.VK_OEM_3,
        '[': win32con.VK_OEM_4,
        '\\': win32con.VK_OEM_5,
        ']': win32con.VK_OEM_6,
        "'": win32con.VK_OEM_7,
    }
    
    shift_chars = {
        '!': '1', '@': '2', '#': '3', '$': '4', '%': '5',
        '^': '6', '&': '7', '*': '8', '(': '9', ')': '0',
        '_': '-', '+': '=', '{': '[', '}': ']', '|': '\\',
        ':': ';', '"': "'", '<': ',', '>': '.', '?': '/'
    }
    
    try:
        if char in special_chars:
            vk_code = special_chars[char]
            win32api.keybd_event(vk_code, 0, 0, 0)
            win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif char in shift_chars:
            base_char = shift_chars[char]
            vk_code = ord(base_char.upper())
            win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
            win32api.keybd_event(vk_code, 0, 0, 0)
            win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
        else:
            # 对于其他字符，使用Unicode输入
            _create_unicode_input(char)
    except Exception as e:
        logging.error(f"发送特殊字符 '{char}' 失败: {e}")


def _create_unicode_input(char: str):
    """
    使用Unicode方式输入字符
    """
    try:
        # 使用Windows API的SendInput发送Unicode字符
        from ctypes.wintypes import WORD, DWORD
        from ctypes import Structure, Union, c_ulong, sizeof, byref
        
        class KEYBDINPUT(Structure):
            _fields_ = [
                ("wVk", WORD),
                ("wScan", WORD),
                ("dwFlags", DWORD),
                ("time", DWORD),
                ("dwExtraInfo", c_ulong),
            ]
        
        class HARDWAREINPUT(Structure):
            _fields_ = [
                ("uMsg", DWORD),
                ("wParamL", WORD),
                ("wParamH", WORD),
            ]
        
        class MOUSEINPUT(Structure):
            _fields_ = [
                ("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", DWORD),
                ("dwFlags", DWORD),
                ("time", DWORD),
                ("dwExtraInfo", c_ulong),
            ]
        
        class _INPUT_UNION(Union):
            _fields_ = [
                ("ki", KEYBDINPUT),
                ("mi", MOUSEINPUT),
                ("hi", HARDWAREINPUT),
            ]
        
        class _INPUT(Structure):
            _anonymous_ = ("u",)
            _fields_ = [
                ("type", DWORD),
                ("u", _INPUT_UNION),
            ]
        
        # 发送Unicode字符
        for c in char:
            inputs = []
            
            # 按下键
            ki = KEYBDINPUT()
            ki.wVk = 0
            ki.wScan = ord(c)
            ki.dwFlags = win32con.KEYEVENTF_UNICODE
            ki.time = 0
            ki.dwExtraInfo = 0
            
            input_down = _INPUT()
            input_down.type = win32con.INPUT_KEYBOARD
            input_down.u.ki = ki
            inputs.append(input_down)
            
            # 释放键
            ki = KEYBDINPUT()
            ki.wVk = 0
            ki.wScan = ord(c)
            ki.dwFlags = win32con.KEYEVENTF_UNICODE | win32con.KEYEVENTF_KEYUP
            ki.time = 0
            ki.dwExtraInfo = 0
            
            input_up = _INPUT()
            input_up.type = win32con.INPUT_KEYBOARD
            input_up.u.ki = ki
            inputs.append(input_up)
            
            # 发送输入
            input_array = (_INPUT * len(inputs))(*inputs)
            ctypes.windll.user32.SendInput(len(inputs), byref(input_array), sizeof(_INPUT))
            
    except Exception as e:
        # 如果Unicode输入失败，尝试直接发送WM_CHAR消息
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                for c in char:
                    win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(c), 0)
        except Exception as e2:
            logging.error(f"Unicode输入失败: {e}, WM_CHAR也失败: {e2}")


# 为了向后兼容，保留这些结构定义
try:
    from ctypes.wintypes import WORD, DWORD
    from ctypes import Structure, Union, c_ulong
    
    class KEYBDINPUT(Structure):
        _fields_ = [
            ("wVk", WORD),
            ("wScan", WORD),
            ("dwFlags", DWORD),
            ("time", DWORD),
            ("dwExtraInfo", c_ulong),
        ]
    
    class HARDWAREINPUT(Structure):
        _fields_ = [
            ("uMsg", DWORD),
            ("wParamL", WORD),
            ("wParamH", WORD),
        ]
    
    class MOUSEINPUT(Structure):
        _fields_ = [
            ("dx", ctypes.c_long),
            ("dy", ctypes.c_long),
            ("mouseData", DWORD),
            ("dwFlags", DWORD),
            ("time", DWORD),
            ("dwExtraInfo", c_ulong),
        ]
    
    class _INPUT_UNION(Union):
        _fields_ = [
            ("ki", KEYBDINPUT),
            ("mi", MOUSEINPUT),
            ("hi", HARDWAREINPUT),
        ]
    
    class _INPUT(Structure):
        _anonymous_ = ("u",)
        _fields_ = [
            ("type", DWORD),
            ("u", _INPUT_UNION),
        ]
        
except ImportError:
    # 如果导入失败，定义空的结构体类
    class KEYBDINPUT:
        pass
    class HARDWAREINPUT:
        pass
    class MOUSEINPUT:
        pass
    class _INPUT_UNION:
        pass
    class _INPUT:
        pass