import os
import win32gui
import win32con
import win32api
import win32process
import win32com.client
import json
import time
import traceback
import sys
import ctypes
import re
import math
from PIL import Image, ImageDraw, ImageFont
import pythoncom
from typing import List, Dict, Optional, Tuple, Union, Any
import random
import logging
from ctypes import wintypes
import tkinter as tk
import platform
import hashlib
import shutil
import threading
from io import BytesIO

from config import ICON_DIR, SETTINGS_FILE, DEFAULT_SETTINGS

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)

current_dir = os.path.dirname(os.path.abspath(__file__))

ICON_DIR = os.path.join(current_dir, "icons")
os.makedirs(ICON_DIR, exist_ok=True)

def ensure_chrome_png_exists():
    chrome_png_path = os.path.join(ICON_DIR, "chrome.png")
    if not os.path.exists(chrome_png_path):
        try:
            from PIL import Image
            img = Image.new('RGBA', (256, 256), (0, 0, 0, 0))
            img.save(chrome_png_path)
        except Exception as e:
            pass

ensure_chrome_png_exists()

def log_error(message: str, exception: Optional[Exception] = None):
    error_str = f"错误: {message}"
    if exception:
        error_str += f" - {str(exception)}\n{traceback.format_exc()}"

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        log_error("检查管理员权限失败", e)
        return False

def run_as_admin():
    try:
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
    except Exception as e:
        log_error("以管理员权限运行失败", e)

def load_settings() -> dict:
    from config import DEFAULT_SETTINGS, SETTINGS_FILE
    
    settings = DEFAULT_SETTINGS.copy()
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                file_settings = json.load(f)
                settings.update(file_settings)
        else:
            save_settings(settings)
    except Exception as e:
        log_error("加载设置失败", e)
    
    return settings

def save_settings(settings: dict):
    from config import SETTINGS_FILE
    
    try:
        settings_dir = os.path.dirname(SETTINGS_FILE)
        if not os.path.exists(settings_dir) and settings_dir:
            os.makedirs(settings_dir, exist_ok=True)
            
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)
            
        print("配置保存成功")
    except Exception as e:
        print(f"保存设置失败: {str(e)}")
        log_error("保存设置失败", e)

def parse_window_numbers(numbers_str: str) -> List[int]:
    if not numbers_str.strip():
        return list(range(1, 49))

    result = []
    parts = numbers_str.split(",")
    for part in parts:
        part = part.strip()
        if "-" in part:
            start, end = map(int, part.split("-"))
            result.extend(range(start, end + 1))
        else:
            try:
                result.append(int(part))
            except ValueError:
                log_error(f"无效的窗口编号: {part}")
    
    return sorted(list(set(result)))

def generate_color_icon(window_number: int, size=256) -> Optional[str]:
    try:
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        icon_dir = os.path.join(base_dir, "icons")
        if not os.path.exists(icon_dir):
            try:
                os.makedirs(icon_dir, exist_ok=True)
            except Exception as e:
                pass
                
        icon_path = os.path.join(icon_dir, f"{window_number}.ico")
        
        if os.path.exists(icon_path):
            return icon_path
            
        r = 30
        g = 30
        b = 30
        
        img = Image.new("RGBA", (size, size), color=(0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        bg_image_path = os.path.join(icon_dir, "chrome.png")
        if os.path.exists(bg_image_path):
            bg_image = Image.open(bg_image_path).resize((size, size))
            img.paste(bg_image, (0, 0))
            
        scale_factor = size / 48
        
        ellipse_width = size * 0.85
        ellipse_height = size * 0.5
        ellipse_left = (size - ellipse_width) / 2
        ellipse_top = (size - ellipse_height) / 2 + (12 * scale_factor)
        ellipse_right = ellipse_left + ellipse_width
        ellipse_bottom = ellipse_top + ellipse_height
        
        draw.ellipse(
            (ellipse_left, ellipse_top, ellipse_right, ellipse_bottom),
            fill=(r, g, b, 255),
        )
        
        font_size = int(24 * scale_factor)
        
        font = None
        try:
            font_path_regular = os.path.join(os.environ["WINDIR"], "Fonts", "Arial.ttf")
            if os.path.exists(font_path_regular):
                 font = ImageFont.truetype(font_path_regular, font_size)
            else:
                font = ImageFont.load_default()
        except Exception as font_error:
            font = ImageFont.load_default()
            
        text = str(window_number)
        
        try:
            if hasattr(font, "getbbox"):
                bbox = font.getbbox(text)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
            elif hasattr(draw, "textsize"):
                text_width, text_height = draw.textsize(text, font=font)
            else:
                text_width = font_size * len(text) * 0.6
                text_height = font_size
            x = (size - text_width) / 2
            y = (size - text_height) / 2 + (8 * scale_factor)
            
            text_color = (255, 255, 255, 255)
            draw.text((x, y), text, fill=text_color, font=font)
        except Exception as text_error:
            x = size // 4
            y = size // 4
            draw.text((x, y), text, fill=(255, 255, 255, 255))
        
        try:
            # 保存为多个分辨率的ICO文件
            sizes = [16, 24, 32, 48, 64, 128, 256]
            images = []
            for s in sizes:
                if s <= size:
                    resized = img.resize((s, s), Image.LANCZOS)
                    images.append(resized)
            
            img.save(icon_path, format="ICO", sizes=[(img.width, img.height) for img in images])
            
            print(f"图标保存成功: {icon_path}")
            return icon_path
        except Exception as save_error:
            print(f"保存ICO失败: {save_error}")
            png_path = icon_path.replace('.ico', '.png')
            img.save(png_path, format="PNG")
            return png_path
            
    except Exception as e:
        print(f"生成彩色图标失败: {str(e)}")
        log_error(f"生成彩色图标失败 (窗口{window_number})", e)
        return None

def title_similarity(title1: str, title2: str) -> float:
    """
    计算两个标题的相似度
    """
    try:
        # 简单的基于词汇的相似度算法
        words1 = set(title1.lower().split())
        words2 = set(title2.lower().split())
        
        if not words1 and not words2:
            return 1.0  # 都是空的，认为相同
        
        if not words1 or not words2:
            return 0.0  # 其中一个为空，不相同
        
        intersection = words1.intersection(words2)
        union = words1.union(words2)
        
        return len(intersection) / len(union)
    except Exception as e:
        log_error("计算标题相似度失败", e)
        return 0.0

def get_chrome_popups(chrome_hwnd: int) -> List[int]:
    def enum_windows_callback(hwnd, _):
        if hwnd != chrome_hwnd and win32gui.IsWindowVisible(hwnd):
            try:
                parent_hwnd = win32gui.GetParent(hwnd)
                owner_hwnd = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
                
                if parent_hwnd == chrome_hwnd or owner_hwnd == chrome_hwnd:
                    popup_hwnds.append(hwnd)
            except Exception:
                pass
        return True
    
    popup_hwnds = []
    try:
        win32gui.EnumWindows(enum_windows_callback, None)
    except Exception as e:
        log_error("枚举弹窗窗口失败", e)
    
    return popup_hwnds

def normalize_path(path: str) -> str:
    """
    标准化路径，处理双反斜杠等问题
    """
    try:
        normalized = os.path.normpath(path)
        # 移除可能的双引号
        normalized = normalized.strip('"')
        return normalized
    except Exception as e:
        log_error("路径标准化失败", e)
        return path

def center_window(window, parent=None):
    """
    将窗口居中显示
    """
    try:
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        
        if parent:
            parent_x = parent.winfo_x()
            parent_y = parent.winfo_y()
            parent_width = parent.winfo_width()
            parent_height = parent.winfo_height()
            
            x = parent_x + (parent_width - width) // 2
            y = parent_y + (parent_height - height) // 2
        else:
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
        
        window.geometry(f"+{x}+{y}")
    except Exception as e:
        log_error("窗口居中失败", e)

def find_chrome_path() -> str:
    """查找Chrome浏览器路径"""
    possible_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe")
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    return ""

def update_screen_list() -> Tuple[List[Dict], List[str]]:
    """更新屏幕列表"""
    screens = []
    screen_names = []
    
    try:
        def callback(hmonitor, hdc, lprect, lparam):
            monitor_info = win32api.GetMonitorInfo(hmonitor)
            monitor_rect = monitor_info['Monitor']
            work_rect = monitor_info['Work']
            
            device_name = monitor_info.get('Device', f'Screen {len(screens) + 1}')
            
            width = monitor_rect[2] - monitor_rect[0]
            height = monitor_rect[3] - monitor_rect[1]
            
            is_primary = monitor_info.get('Flags', 0) & 1 == 1
            
            screen_info = {
                'name': device_name,
                'width': width,
                'height': height,
                'left': monitor_rect[0],
                'top': monitor_rect[1],
                'right': monitor_rect[2],
                'bottom': monitor_rect[3],
                'work_left': work_rect[0],
                'work_top': work_rect[1],
                'work_right': work_rect[2],
                'work_bottom': work_rect[3],
                'is_primary': is_primary,
                'handle': hmonitor
            }
            
            screens.append(screen_info)
            
            primary_text = " (主显示器)" if is_primary else ""
            screen_name = f"{device_name} - {width}x{height}{primary_text}"
            screen_names.append(screen_name)
            
            return True
        
        win32api.EnumDisplayMonitors(None, None, callback, 0)
        
        # 按主显示器优先，然后按左上角位置排序
        screens.sort(key=lambda s: (not s['is_primary'], s['left'], s['top']))
        screen_names.sort()
        
    except Exception as e:
        log_error("更新屏幕列表失败", e)
        # 添加默认屏幕
        screens = [{
            'name': 'Screen 1',
            'width': 1920,
            'height': 1080,
            'left': 0,
            'top': 0,
            'right': 1920,
            'bottom': 1080,
            'work_left': 0,
            'work_top': 0,
            'work_right': 1920,
            'work_bottom': 1040,
            'is_primary': True,
            'handle': None
        }]
        screen_names = ["Screen 1 - 1920x1080 (主显示器)"]
    
    return screens, screen_names

def set_chrome_icon(hwnd: int, icon_path: str, retries=3, delay=0.3) -> bool:
    """
    为Chrome窗口设置自定义图标
    """
    if not os.path.exists(icon_path):
        print(f"图标文件不存在: {icon_path}")
        return False
    
    for attempt in range(retries):
        try:
            # 加载图标
            hicon = win32gui.LoadImage(
                None, icon_path, win32con.IMAGE_ICON,
                0, 0, win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            )
            
            if hicon == 0:
                print(f"加载图标失败: {icon_path}")
                continue
            
            # 设置窗口图标
            result1 = win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, hicon)
            result2 = win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, hicon)
            
            # 强制重绘窗口
            win32gui.InvalidateRect(hwnd, None, True)
            win32gui.UpdateWindow(hwnd)
            
            print(f"图标设置成功: {icon_path} -> 窗口 {hwnd}")
            return True
            
        except Exception as e:
            print(f"设置图标失败 (尝试 {attempt + 1}/{retries}): {str(e)}")
            if attempt < retries - 1:
                time.sleep(delay)
    
    return False

def show_notification(title: str, message: str):
    """
    显示系统通知
    """
    try:
        import win10toast
        toaster = win10toast.ToastNotifier()
        toaster.show_toast(title, message, duration=3)
    except ImportError:
        try:
            import tkinter.messagebox as msgbox
            msgbox.showinfo(title, message)
        except Exception:
            print(f"通知: {title} - {message}")
    except Exception as e:
        print(f"显示通知失败: {str(e)}")

def is_ultrawide_screen(screen_width, screen_height):
    """判断是否为超宽屏"""
    return screen_width / screen_height >= 2.0

def arrange_windows_by_profile_id(profiles_data: List[Dict], screen_index: int = 0):
    """根据配置ID排列窗口"""
    try:
        screens, _ = update_screen_list()
        if not screens or screen_index >= len(screens):
            print("无效的屏幕索引")
            return False
        
        target_screen = screens[screen_index]
        screen_width = target_screen['work_right'] - target_screen['work_left']
        screen_height = target_screen['work_bottom'] - target_screen['work_top']
        screen_left = target_screen['work_left']
        screen_top = target_screen['work_top']
        
        # 根据窗口数量和屏幕尺寸计算布局
        window_count = len(profiles_data)
        if window_count == 0:
            return True
        
        # 计算网格布局
        if is_ultrawide_screen(screen_width, screen_height):
            # 超宽屏使用更多列
            cols = min(window_count, 4)
        else:
            cols = min(window_count, 3)
        
        rows = math.ceil(window_count / cols)
        
        window_width = screen_width // cols
        window_height = screen_height // rows
        
        for i, profile in enumerate(profiles_data):
            if 'hwnd' not in profile:
                continue
                
            hwnd = profile['hwnd']
            row = i // cols
            col = i % cols
            
            x = screen_left + col * window_width
            y = screen_top + row * window_height
            
            try:
                win32gui.SetWindowPos(
                    hwnd, win32con.HWND_TOP,
                    x, y, window_width, window_height,
                    win32con.SWP_SHOWWINDOW
                )
                
                print(f"窗口 {profile.get('number', '?')} 移动到 ({x}, {y}) 大小 {window_width}x{window_height}")
                
            except Exception as e:
                print(f"移动窗口失败: {str(e)}")
        
        return True
        
    except Exception as e:
        log_error("排列窗口失败", e)
        return False