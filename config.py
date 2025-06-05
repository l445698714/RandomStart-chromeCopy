import os
import sys
import ctypes
import win32con
from ctypes import wintypes
import json
from typing import Dict, Any

WINDOW_WIDTH = 700
WINDOW_HEIGHT = 360
STANDARD_WHEEL_DELTA = 120
WHEEL_THRESHOLD = 0.05
BASE_DEBUG_PORT = 9222
DWMWA_BORDER_COLOR = 34
DWM_MAGIC_COLOR = 0x00FF0000
DEFAULT_ICON_SIZE = 256

if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ICON_DIR = os.path.join(BASE_DIR, "icons")
os.makedirs(ICON_DIR, exist_ok=True)

PROCESS_CREATION_MITIGATION_POLICY_BLOCK_NON_MICROSOFT_BINARIES_ALWAYS_ON = 0x100000000000

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("pt", POINT),
        ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG)),
    ]

STYLES = {
    "default_font": ("Microsoft YaHei UI", 9),
    "small_entry": {"padding": (4, 0)},
    "link_label": {
        "foreground": "#0d6efd",
        "cursor": "hand2",
        "font": ("Microsoft YaHei UI", 9, "underline")
    },
    "master_tag": {
        "background": "#0d6efd",
        "foreground": "white"
    }
}

SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

DEFAULT_SETTINGS: Dict[str, Any] = {
    "shortcut_path": "",
    "cache_dir": "",
    "screen_selection": "",
    "admin_confirmed": False,
    "custom_urls": [],
    "auto_modify_shortcut_icon": True,
    "show_chrome_tip": True,
    "sync_shortcut": None,
    "window_position": None,
    "screen_arrange_config": []
}