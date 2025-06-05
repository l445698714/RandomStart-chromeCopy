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
        # 使用剪贴板方式发送文本（更可靠）
        import win32clipboard
        
        # 将文本放入剪贴板
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
        win32clipboard.CloseClipboard()
        
        # 发送 Ctrl+V 粘贴
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        
    except Exception as e:
        # 备用方案：逐字符发送
        for char in text:
            _send_char(char)


def _send_char(char: str):
    """
    发送单个字符到当前活动窗口
    """
    try:
        # 对于特殊字符，使用Unicode输入
        if ord(char) > 127:
            # Unicode字符
            win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)  # Alt down
            for digit in str(ord(char)):
                vk_code = ord(digit) - ord('0') + win32con.VK_NUMPAD0
                win32api.keybd_event(vk_code, 0, 0, 0)
                win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)  # Alt up
        else:
            # ASCII字符
            if char.isalpha():
                # 字母
                vk_code = ord(char.upper())
                if char.islower():
                    win32api.keybd_event(vk_code, 0, 0, 0)
                    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                else:
                    win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
                    win32api.keybd_event(vk_code, 0, 0, 0)
                    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
            elif char.isdigit():
                # 数字
                vk_code = ord(char)
                win32api.keybd_event(vk_code, 0, 0, 0)
                win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            else:
                # 特殊字符
                _send_special_char(char)
                
    except Exception as e:
        logging.error(f"发送字符失败: {char} - {e}")


def _send_special_char(char: str):
    """
    发送特殊字符
    """
    special_chars = {
        ' ': win32con.VK_SPACE,
        '.': win32con.VK_OEM_PERIOD,
        ',': win32con.VK_OEM_COMMA,
        ';': win32con.VK_OEM_1,
        '/': win32con.VK_OEM_2,
        '`': win32con.VK_OEM_3,
        '[': win32con.VK_OEM_4,
        '\\': win32con.VK_OEM_5,
        ']': win32con.VK_OEM_6,
        "'": win32con.VK_OEM_7,
        '-': win32con.VK_OEM_MINUS,
        '=': win32con.VK_OEM_PLUS,
        '\n': win32con.VK_RETURN,
        '\t': win32con.VK_TAB,
    }
    
    if char in special_chars:
        vk_code = special_chars[char]
        win32api.keybd_event(vk_code, 0, 0, 0)
        win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
    else:
        # 对于其他特殊字符，尝试使用SendInput
        try:
            ctypes.windll.user32.SendInput(1, ctypes.byref(_create_unicode_input(char)), ctypes.sizeof(_INPUT))
        except:
            pass


def _create_unicode_input(char: str):
    """
    创建Unicode输入结构
    """
    try:
        from ctypes import Structure, Union, c_ulong, c_ushort, c_short, POINTER
        
        class KEYBDINPUT(Structure):
            _fields_ = [
                ("wVk", c_ushort),
                ("wScan", c_ushort),
                ("dwFlags", c_ulong),
                ("time", c_ulong),
                ("dwExtraInfo", POINTER(c_ulong))
            ]
        
        class HARDWAREINPUT(Structure):
            _fields_ = [
                ("uMsg", c_ulong),
                ("wParamL", c_short),
                ("wParamH", c_ushort)
            ]
        
        class MOUSEINPUT(Structure):
            _fields_ = [
                ("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", c_ulong),
                ("dwFlags", c_ulong),
                ("time", c_ulong),
                ("dwExtraInfo", POINTER(c_ulong))
            ]
        
        class _INPUT_UNION(Union):
            _fields_ = [
                ("mi", MOUSEINPUT),
                ("ki", KEYBDINPUT),
                ("hi", HARDWAREINPUT)
            ]
        
        class _INPUT(Structure):
            _anonymous_ = ("u",)
            _fields_ = [
                ("type", c_ulong),
                ("u", _INPUT_UNION)
            ]
        
        # 创建输入结构
        inp = _INPUT()
        inp.type = 1  # INPUT_KEYBOARD
        inp.ki.wVk = 0
        inp.ki.wScan = ord(char)
        inp.ki.dwFlags = 4  # KEYEVENTF_UNICODE
        inp.ki.time = 0
        inp.ki.dwExtraInfo = None
        
        return inp
        
    except Exception:
        return None


# 全局常量
_INPUT = None
try:
    from ctypes import Structure, Union, c_ulong, c_ushort, c_short, POINTER
    
    class KEYBDINPUT(Structure):
        _fields_ = [
            ("wVk", c_ushort),
            ("wScan", c_ushort),
            ("dwFlags", c_ulong),
            ("time", c_ulong),
            ("dwExtraInfo", POINTER(c_ulong))
        ]
    
    class HARDWAREINPUT(Structure):
        _fields_ = [
            ("uMsg", c_ulong),
            ("wParamL", c_short),
            ("wParamH", c_ushort)
        ]
    
    class MOUSEINPUT(Structure):
        _fields_ = [
            ("dx", ctypes.c_long),
            ("dy", ctypes.c_long),
            ("mouseData", c_ulong),
            ("dwFlags", c_ulong),
            ("time", c_ulong),
            ("dwExtraInfo", POINTER(c_ulong))
        ]
    
    class _INPUT_UNION(Union):
        _fields_ = [
            ("mi", MOUSEINPUT),
            ("ki", KEYBDINPUT),
            ("hi", HARDWAREINPUT)
        ]
    
    class _INPUT(Structure):
        _anonymous_ = ("u",)
        _fields_ = [
            ("type", c_ulong),
            ("u", _INPUT_UNION)
        ]

except Exception:
    pass 