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
             draw.text((x, y), text, fill=(255, 255, 255, 255), font=font)
        
        os.makedirs(os.path.dirname(icon_path), exist_ok=True)
        
        try:
            img.save(icon_path, format="ICO", sizes=[(48, 48), (256, 256)])
        except Exception as save_error:
            png_path = os.path.join(icon_dir, f"{window_number}.png")
            img.save(png_path, format="PNG")
            icon_path = png_path
            
        return icon_path
    except Exception as e:
        log_error(f"生成图标失败: 窗口 {window_number}", e)
        return None

def title_similarity(title1: str, title2: str) -> float:
    if not title1 or not title2:
        return 0.0
    
    clean1 = re.sub(r'[\[\]\(\)\{\}★]', '', title1).strip()
    clean2 = re.sub(r'[\[\]\(\)\{\}★]', '', title2).strip()
    
    if not clean1 or not clean2:
        return 0.0
    
    clean1 = clean1.lower()
    clean2 = clean2.lower()
    
    if clean1 == clean2:
        return 1.0
    
    if clean1 in clean2 or clean2 in clean1:
        return 0.8
    
    set1 = set(clean1)
    set2 = set(clean2)
    
    intersection = len(set1.intersection(set2))
    union = len(set1.union(set2))
    
    if union == 0:
        return 0.0
    
    return intersection / union

def get_chrome_popups(chrome_hwnd: int) -> List[int]:
    popup_list = []
    
    def enum_windows_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            try:
                class_name = win32gui.GetClassName(hwnd)
                if "Chrome_WidgetWin" in class_name:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    _, parent_pid = win32process.GetWindowThreadProcessId(chrome_hwnd)
                    if pid == parent_pid and hwnd != chrome_hwnd:
                        popup_list.append(hwnd)
            except Exception:
                pass
        return True
    
    try:
        win32gui.EnumWindows(enum_windows_callback, None)
    except Exception as e:
        log_error("枚举弹出窗口失败", e)
    
    return popup_list

def normalize_path(path: str) -> str:
    if not path:
        return ""
    
    normalized = os.path.normpath(path)
    
    if os.path.isdir(normalized) and not normalized.endswith(os.path.sep):
        normalized += os.path.sep
    
    return normalized

def center_window(window, parent=None):
    window.update_idletasks()
    
    window_width = window.winfo_width()
    window_height = window.winfo_height()
    
    if parent:
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        x = parent_x + (parent_width - window_width) // 2
        y = parent_y + (parent_height - window_height) // 2
    else:
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
    
    window.geometry(f"+{x}+{y}")

def find_chrome_path() -> str:
    possible_paths = [
        os.path.join(os.environ.get('PROGRAMFILES', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
        os.path.join(os.environ.get('PROGRAMFILES(X86)', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    return ""

def update_screen_list() -> Tuple[List[Dict], List[str]]:
    try:
        screens = []
        screen_names = []

        def callback(hmonitor, hdc, lprect, lparam):
            try:
                monitor_info = win32api.GetMonitorInfo(hmonitor)
                screen_name = f"屏幕 {len(screens) + 1}"
                if monitor_info["Flags"] & 1:
                    screen_name += " (主)"
                
                monitor_rect = monitor_info["Monitor"]
                width = monitor_rect[2] - monitor_rect[0]
                height = monitor_rect[3] - monitor_rect[1]
                
                screen_name += f" - {width}x{height}"
                
                screens.append(
                    {
                        "name": screen_name,
                        "rect": monitor_info["Monitor"],
                        "work_rect": monitor_info["Work"],
                        "monitor": hmonitor,
                    }
                )
                screen_names.append(screen_name)
            except Exception as e:
                pass
            return True

        MONITORENUMPROC = ctypes.WINFUNCTYPE(
            ctypes.c_bool,
            ctypes.c_ulong,
            ctypes.c_ulong,
            ctypes.POINTER(wintypes.RECT),
            ctypes.c_longlong,
        )

        callback_function = MONITORENUMPROC(callback)

        if (
            ctypes.windll.user32.EnumDisplayMonitors(0, 0, callback_function, 0)
            == 0
        ):
            try:
                virtual_width = win32api.GetSystemMetrics(
                    win32con.SM_CXVIRTUALSCREEN
                )
                virtual_height = win32api.GetSystemMetrics(
                    win32con.SM_CYVIRTUALSCREEN
                )
                virtual_left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
                virtual_top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)

                primary_monitor = win32api.MonitorFromPoint(
                    (0, 0), win32con.MONITOR_DEFAULTTOPRIMARY
                )
                primary_info = win32api.GetMonitorInfo(primary_monitor)
                
                monitor_rect = primary_info["Monitor"]
                width = monitor_rect[2] - monitor_rect[0]
                height = monitor_rect[3] - monitor_rect[1]
                screen_name = f"屏幕 1 (主) - {width}x{height}"

                screens.append(
                    {
                        "name": screen_name,
                        "rect": primary_info["Monitor"],
                        "work_rect": primary_info["Work"],
                        "monitor": primary_monitor,
                    }
                )
                screen_names.append(screen_name)

                try:
                    second_monitor = win32api.MonitorFromPoint(
                        (
                            virtual_left + virtual_width - 1,
                            virtual_top + virtual_height // 2,
                        ),
                        win32con.MONITOR_DEFAULTTONULL,
                    )
                    if second_monitor and second_monitor != primary_monitor:
                        second_info = win32api.GetMonitorInfo(second_monitor)
                        
                        monitor_rect = second_info["Monitor"]
                        width = monitor_rect[2] - monitor_rect[0]
                        height = monitor_rect[3] - monitor_rect[1]
                        screen_name = f"屏幕 2 - {width}x{height}"
                        
                        screens.append(
                            {
                                "name": screen_name,
                                "rect": second_info["Monitor"],
                                "work_rect": second_info["Work"],
                                "monitor": second_monitor,
                            }
                        )
                        screen_names.append(screen_name)
                except:
                    pass

            except Exception as e:
                pass

        if not screens:
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            
            screen_name = f"屏幕 1 (主) - {screen_width}x{screen_height}"
            
            screens.append(
                {
                    "name": screen_name,
                    "rect": (0, 0, screen_width, screen_height),
                    "work_rect": (0, 0, screen_width, screen_height),
                    "monitor": None,
                }
            )
            screen_names.append(screen_name)

        screens.sort(key=lambda x: x["rect"][0])

        screen_names = [screen["name"] for screen in screens]

        return screens, screen_names

    except Exception as e:
        return [{"name": "主屏幕 - 1920x1080", "rect": (0, 0, 1920, 1080), "work_rect": (0, 0, 1920, 1080), "monitor": None}], ["主屏幕 - 1920x1080"]

def set_chrome_icon(hwnd: int, icon_path: str, retries=3, delay=0.3) -> bool:
    for attempt in range(retries):
        try:
            if not win32gui.IsWindow(hwnd):
                return False
                
            big_icon = win32gui.LoadImage(
                0, icon_path, win32con.IMAGE_ICON, 32, 32, win32con.LR_LOADFROMFILE
            )
            small_icon = win32gui.LoadImage(
                0, icon_path, win32con.IMAGE_ICON, 16, 16, win32con.LR_LOADFROMFILE
            )
            
            if not win32gui.IsWindow(hwnd):
                return False
            
            win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, big_icon)
            time.sleep(delay/2)
            win32gui.SendMessage(
                hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, small_icon
            )
            
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
            
            time.sleep(delay)
            return True
            
        except Exception as e:
            log_error(f"设置图标失败 (尝试 {attempt+1}/{retries})", e)
            time.sleep(delay)
    
    return False

_active_notifications = set()

def show_notification(title: str, message: str):
    try:
        from plyer import notification
        
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "app.ico")
        if not os.path.exists(icon_path):
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
        
        if os.path.exists(icon_path):
            notification.notify(
                title=title,
                message=message,
                app_name="Chrome多窗口管理器", 
                app_icon=icon_path,
                timeout=5
            )
        else:
            notification.notify(
                title=title,
                message=message,
                app_name="Chrome多窗口管理器",
                timeout=5
            )
        return
            
    except ImportError:
        try:
            import win32gui
            import win32con
            win32gui.MessageBox(0, message, title, win32con.MB_ICONINFORMATION)
        except Exception as err:
            pass
    except Exception as e:
        try:
            import win32gui
            import win32con
            win32gui.MessageBox(0, message, title, win32con.MB_ICONINFORMATION)
        except Exception as err:
            pass 

def is_ultrawide_screen(screen_width, screen_height):
    return (screen_width / screen_height) > 2.1 

def arrange_windows_by_profile_id(profiles_data: List[Dict], screen_index: int = 0):
    """
    按照分身编号顺序排列窗口，从左到右，从上到下。

    Args:
        profiles_data: 一个字典列表，每个字典包含至少 'id' (分身编号, 假设是整数或可排序的字符串)
                       和 'hwnd' (窗口句柄)。
                       例如: [{'id': 1, 'hwnd': 12345}, {'id': 2, 'hwnd': 67890}]
        screen_index: 要在哪个屏幕上排列窗口，0 代表主屏幕。
    """
    if not profiles_data:
        log_error("没有提供分身数据进行排列。", None)
        return

    # 1. 按分身编号排序
    try:
        # 确保 'id' 可以被正确排序，如果是数字字符串，先转为整数
        sorted_profiles = sorted(profiles_data, key=lambda p: int(str(p.get('id', 0))))
    except ValueError:
        log_error("分身ID包含无法转换为整数的值，按原始顺序排列。", None)
        sorted_profiles = sorted(profiles_data, key=lambda p: str(p.get('id', '')))
    except Exception as e:
        log_error(f"排序分身数据时出错: {e}", e)
        return # 或者使用未排序的数据继续，但可能不符合预期

    active_windows = []
    for profile in sorted_profiles:
        hwnd = profile.get('hwnd')
        if hwnd and win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
            active_windows.append(profile)
        else:
            log_error(f"分身 {profile.get('id')} 的窗口句柄无效或不可见，跳过排列。", None)
    
    if not active_windows:
        log_error("没有有效的活动窗口进行排列。", None)
        # show_notification("排列窗口", "没有找到有效的活动窗口进行排列。") # 可选
        return

    num_windows = len(active_windows)

    # 2. 获取目标屏幕的工作区域
    try:
        # 使用已有的 update_screen_list 或类似函数获取屏幕信息
        # 这里简化为直接使用 win32api 获取指定屏幕信息
        # 注意：更健壮的做法是复用 update_screen_list 的逻辑
        monitors = win32api.EnumDisplayMonitors()
        if not monitors or screen_index >= len(monitors):
            log_error(f"屏幕索引 {screen_index} 无效。", None)
            # 默认使用主屏幕
            monitor_handle = win32api.MonitorFromPoint((0,0), win32con.MONITOR_DEFAULTTOPRIMARY)
        else:
            monitor_handle = monitors[screen_index][0] # monitors[i][0] is the HMONITOR

        monitor_info = win32api.GetMonitorInfo(monitor_handle)
        work_area = monitor_info['Work'] # (left, top, right, bottom)
        screen_width = work_area[2] - work_area[0]
        screen_height = work_area[3] - work_area[1]
        screen_left = work_area[0]
        screen_top = work_area[1]

    except Exception as e:
        log_error(f"获取屏幕信息失败: {e}", e)
        # 使用默认的主屏幕尺寸作为后备
        screen_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN) # 或者 SM_CXSCREEN
        screen_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN) # 或者 SM_CYSCREEN
        screen_left = 0
        screen_top = 0
        # 尝试获取主屏幕工作区作为更精确的后备
        try:
            monitor_handle = win32api.MonitorFromPoint((0,0), win32con.MONITOR_DEFAULTTOPRIMARY)
            monitor_info = win32api.GetMonitorInfo(monitor_handle)
            work_area = monitor_info['Work']
            screen_width = work_area[2] - work_area[0]
            screen_height = work_area[3] - work_area[1]
            screen_left = work_area[0]
            screen_top = work_area[1]
        except:
            pass # 保持使用虚拟屏幕尺寸

    if screen_width <= 0 or screen_height <= 0:
        log_error("获取到的屏幕尺寸无效。", None)
        return

    # 3. 计算布局：行列数和窗口尺寸
    # 目标是尽可能平均分配，可以尝试让窗口接近正方形或某个宽高比
    # 简化版：尝试计算每行多少个窗口，使得总行数尽可能少
    if num_windows == 0:
        return

    cols = math.ceil(math.sqrt(num_windows * (screen_width / screen_height))) # 尝试根据屏幕宽高比调整列数估计
    if cols == 0: cols = 1
    rows = math.ceil(num_windows / cols)
    if rows == 0: rows = 1
    
    # 再次调整，确保列数不会过多导致窗口过窄
    # 如果每列的窗口宽度小于某个阈值（例如300像素），则减少列数
    min_sensible_width = 400 # 假设Chrome窗口有个合理的最小宽度
    while cols > 1 and (screen_width / cols) < min_sensible_width:
        cols -= 1
        rows = math.ceil(num_windows / cols)

    win_width = int(screen_width / cols)
    win_height = int(screen_height / rows)

    if win_width <=0 or win_height <=0:
        log_error(f"计算得到的窗口尺寸无效: W={win_width}, H={win_height}", None)
        return

    # 4. 排列窗口
    for i, profile in enumerate(active_windows):
        hwnd = profile['hwnd']
        
        row = i // cols
        col = i % cols
        
        x = screen_left + col * win_width
        y = screen_top + row * win_height
        
        try:
            # 确保窗口不是最小化状态，如果是，先恢复
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.05) # 给窗口一点时间响应

            # 将窗口置顶，以便 MoveWindow 生效（有时被其他窗口遮挡会影响）
            # win32gui.SetForegroundWindow(hwnd) # 这个要小心使用，可能会抢焦点
            # win32gui.BringWindowToTop(hwnd) # 更温和的方式

            # 设置窗口位置和大小
            # SWP_NOZORDER: 保持Z序 SWP_NOACTIVATE: 不激活窗口
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, win_width, win_height, win32con.SWP_NOACTIVATE)
            # 或者使用 MoveWindow，它会激活窗口
            # win32gui.MoveWindow(hwnd, x, y, win_width, win_height, True)
            log_error(f"排列分身 {profile.get('id')}: HWND={hwnd} 到 X={x}, Y={y}, W={win_width}, H={win_height}", None) # 使用log_error临时调试
        except Exception as e:
            log_error(f"排列窗口 {hwnd} (分身 {profile.get('id')}) 失败: {e}", e)

    # show_notification("排列完成", f"{num_windows} 个窗口已按编号排列。") # 可选 