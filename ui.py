import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import webbrowser
import threading
import time
import sv_ttk
import win32gui
import win32com.client
import ctypes
import random
import keyboard
import math
import win32con
import win32api
import traceback
import pystray
from PIL import Image
import sys
import json
import queue
import ctypes
import shutil
import random
import asyncio
import logging
import platform
import threading
import tkinter as tk
import webbrowser
import win32com.client
import subprocess

from core import ChromeManager
from utils import (
    center_window,
    parse_window_numbers,
    log_error,
    load_settings,
    save_settings,
    update_screen_list,
    show_notification,
    generate_color_icon
)
from input_tools import input_random_number, input_text_from_file
from config import STYLES

class ChromeManagerUI:

    def __init__(self):
        self.start_time = time.time()
        self.root = tk.Tk()
        self.root.title("NoBiggie社区Chrome多窗口管理器 V3.0")
        self.root.withdraw()
        try:
            self.app_icon_path = os.path.join(os.path.dirname(__file__), "icons", "app.ico")
            if os.path.exists(self.app_icon_path):
                self.root.iconbitmap(self.app_icon_path)
        except Exception as e:
            self.app_icon_path = None
            log_error("设置图标失败", e)
        
        self.window_width = 700
        self.window_height = 380
        self.root.geometry(f"{self.window_width}x{self.window_height}")
        self.root.resizable(False, False)
        
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        
        sv_ttk.set_theme("light")
        print(f"[{time.time() - self.start_time:.3f}s] 主题加载完成")
        
        self.settings = load_settings()
        
        last_position = self.load_window_position()
        if last_position:
            try:
                self.root.geometry(f"{self.window_width}x{self.window_height}{last_position}")
            except Exception as e:
                log_error("应用窗口位置失败", e)
        
        self.manager = ChromeManager(self)
        self.manager.ui_update_callback = self.update_window_list
        
        self.random_min_value = tk.StringVar(value="1000")
        self.random_max_value = tk.StringVar(value="2000")
        self.random_overwrite = tk.BooleanVar(value=True)
        self.random_delayed = tk.BooleanVar(value=False)
        
        self.window_list = None
        self.select_all_var = tk.StringVar(value="全部选择")
        self.screens = []
        self.screens_names = []
        self.shortcut_path = self.settings.get("shortcut_path", "")
        self.cache_dir = self.settings.get("cache_dir", "")
        self.screen_selection = self.settings.get("screen_selection", "")
        self.custom_urls = self.settings.get("custom_urls", [])
        self.selected_url = tk.StringVar()
        
        # 新增状态变量，用于依次启动和随机启动
        self.sequential_launch_active = False
        self.sequential_launch_range_str = None
        self.sequential_launch_profiles = [] # 存储解析后的数字列表
        self.sequential_launch_current_index = 0
        self.random_launch_count_var = tk.StringVar(value="1") # 随机启动的数量
        self.last_launched_sequentially = None # 记录上一个依次启动的分身号

        self.create_styles()
        self.create_widgets()
        self.update_treeview_style()
        self.create_context_menus()
        
        self.show_chrome_tip = True
        
        if "show_chrome_tip" in self.settings:
            self.show_chrome_tip = self.settings["show_chrome_tip"]
        
        self.close_behavior = self.settings.get("close_behavior", None)
        
        self.create_systray_icon()
        
        self.root.after(100, self.delayed_initialization)
        print(f"[{time.time() - self.start_time:.3f}s] __init__ 完成, 已安排延迟初始化")
    
    def create_systray_icon(self):
        try:
            menu = (
                pystray.MenuItem('打开管理器面板', self.show_window, default=True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem('关闭', self.quit_window)
            )
            
            image = Image.open(self.app_icon_path) if self.app_icon_path else Image.new('RGB', (64, 64), color='blue')
            self.icon = pystray.Icon("chrome_manager", image, "Chrome多窗口管理器", menu)
            
            threading.Thread(target=self.icon.run, daemon=True).start()
            
            self.root.after(500, self.show_window)
            
        except Exception as e:
            log_error("创建系统托盘图标失败", e)
            self.root.after(100, self.root.deiconify)
    
    def hide_window(self):
        try:
            if self.close_behavior is None:
                result = messagebox.askyesnocancel(
                    "关闭选项", 
                    "您希望如何处理程序？\n\n点击'是'：最小化到系统托盘\n点击'否'：直接退出程序\n点击'取消'：取消此次操作",
                    icon="question"
                )
                
                if result is None:
                    return
                elif result:
                    self.close_behavior = "minimize"
                else:
                    self.close_behavior = "exit"
                
                self.settings["close_behavior"] = self.close_behavior
                save_settings(self.settings)
            
            if self.close_behavior == "minimize":
                self.root.withdraw()
                if hasattr(self, 'icon'):
                    self.icon.notify("程序已最小化到系统托盘，点击图标可以重新打开", "Chrome多窗口管理器")
            else:
                self.quit_window()
                
        except Exception as e:
            log_error("处理窗口关闭事件失败", e)
    
    def show_window(self, *args):
        try:
            self.root.deiconify()
            self.root.focus_force()
        except Exception as e:
            log_error("显示窗口失败", e)
    
    def quit_window(self, *args):
        try:
            if hasattr(self, 'icon'):
                self.icon.stop()
            
            self.on_closing()
        except Exception as e:
            log_error("退出程序失败", e)
            self.root.destroy()
    
    def create_styles(self):
        style = ttk.Style()
        
        default_font = STYLES["default_font"]
        
        style.configure("Small.TEntry", padding=STYLES["small_entry"]["padding"], font=default_font)
        
        style.configure("TButton", font=default_font)
        style.configure("TLabel", font=default_font)
        style.configure("TEntry", font=default_font)
        style.configure("Treeview", font=default_font)
        style.configure("Treeview.Heading", font=default_font)
        style.configure("TLabelframe.Label", font=default_font)
        style.configure("TNotebook.Tab", font=default_font)
        
        # 链接样式
        style.configure(
            "Link.TLabel",
            foreground=STYLES["link_label"]["foreground"],
            cursor=STYLES["link_label"]["cursor"],
            font=STYLES["link_label"]["font"],
        )
    
    def update_treeview_style(self):
        if self.window_list:
            self.window_list.tag_configure(
                "master",
                background=STYLES["master_tag"]["background"],
                foreground=STYLES["master_tag"]["foreground"]
            )
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.X, padx=10, pady=5)
        
        upper_frame = ttk.Frame(main_frame)
        upper_frame.pack(fill=tk.X)
        
        arrange_frame = ttk.LabelFrame(upper_frame, text="自定义排列")
        arrange_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(3, 0))
        
        manage_frame = ttk.LabelFrame(upper_frame, text="窗口管理")
        manage_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        button_rows = ttk.Frame(manage_frame)
        button_rows.pack(fill=tk.X)
        
        first_row = ttk.Frame(button_rows)
        first_row.pack(fill=tk.X)
        
        ttk.Button(
            first_row,
            text="导入窗口",
            command=self.import_windows,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=2)
        
        select_all_label = ttk.Label(
            first_row, textvariable=self.select_all_var, style="Link.TLabel"
        )
        select_all_label.pack(side=tk.LEFT, padx=5)
        select_all_label.bind("<Button-1>", self.toggle_select_all)
        
        ttk.Button(first_row, text="自动排列", command=self.auto_arrange_windows).pack(
            side=tk.LEFT, padx=2
        )
        
        ttk.Button(
            first_row, text="关闭选中", command=self.close_selected_windows
        ).pack(side=tk.LEFT, padx=2)
        
        self.sync_button = ttk.Button(
            first_row,
            text="▶ 开始同步",
            command=self.toggle_sync,
            style="Accent.TButton",
        )
        self.sync_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            first_row, text="🔗 设置", command=self.show_settings_dialog, width=8
        ).pack(side=tk.LEFT, padx=2)
        
        list_frame = ttk.Frame(manage_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        
        self.window_list = ttk.Treeview(
            list_frame,
            columns=("select", "number", "title", "master", "hwnd"),
            show="headings",
            height=4,
            style="Accent.Treeview",
        )
        self.window_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.window_list.heading("select", text="选择")
        self.window_list.heading("number", text="窗口序号")
        self.window_list.heading("title", text="页面标题")
        self.window_list.heading("master", text="主控")
        self.window_list.heading("hwnd", text="")
        
        self.window_list.column("select", width=50, anchor="center")
        self.window_list.column("number", width=60, anchor="center")
        self.window_list.column("title", width=260)
        self.window_list.column("master", width=50, anchor="center")
        self.window_list.column("hwnd", width=0, stretch=False)
        
        self.window_list.tag_configure("master", background="lightblue")
        
        self.window_list.bind("<Button-1>", self.on_click)
        
        scrollbar = ttk.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=self.window_list.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.window_list.configure(yscrollcommand=scrollbar.set)
        
        params_frame = ttk.Frame(arrange_frame)
        params_frame.pack(fill=tk.X, padx=5, pady=2)
        
        left_frame = ttk.Frame(params_frame)
        left_frame.pack(side=tk.LEFT, padx=(0, 5))
        right_frame = ttk.Frame(params_frame)
        right_frame.pack(side=tk.LEFT)
        
        ttk.Label(left_frame, text="起始X坐标").pack(anchor=tk.W)
        self.start_x = ttk.Entry(left_frame, width=8, style="Small.TEntry")
        self.start_x.pack(fill=tk.X, pady=(0, 2))
        self.start_x.insert(0, "0")
        self.setup_right_click_menu(self.start_x)
        
        ttk.Label(left_frame, text="窗口宽度").pack(anchor=tk.W)
        self.window_width_entry = ttk.Entry(left_frame, width=8, style="Small.TEntry")
        self.window_width_entry.pack(fill=tk.X, pady=(0, 2))
        self.window_width_entry.insert(0, "500")
        self.setup_right_click_menu(self.window_width_entry)
        
        ttk.Label(left_frame, text="水平间距").pack(anchor=tk.W)
        self.h_spacing = ttk.Entry(left_frame, width=8, style="Small.TEntry")
        self.h_spacing.pack(fill=tk.X, pady=(0, 2))
        self.h_spacing.insert(0, "0")
        self.setup_right_click_menu(self.h_spacing)
        
        ttk.Label(right_frame, text="起始Y坐标").pack(anchor=tk.W)
        self.start_y = ttk.Entry(right_frame, width=8, style="Small.TEntry")
        self.start_y.pack(fill=tk.X, pady=(0, 2))
        self.start_y.insert(0, "0")
        self.setup_right_click_menu(self.start_y)
        
        ttk.Label(right_frame, text="窗口高度").pack(anchor=tk.W)
        self.window_height_entry = ttk.Entry(right_frame, width=8, style="Small.TEntry")
        self.window_height_entry.pack(fill=tk.X, pady=(0, 2))
        self.window_height_entry.insert(0, "400")
        self.setup_right_click_menu(self.window_height_entry)
        
        ttk.Label(right_frame, text="垂直间距").pack(anchor=tk.W)
        self.v_spacing = ttk.Entry(right_frame, width=8, style="Small.TEntry")
        self.v_spacing.pack(fill=tk.X, pady=(0, 2))
        self.v_spacing.insert(0, "0")
        self.setup_right_click_menu(self.v_spacing)
        
        for widget in left_frame.winfo_children() + right_frame.winfo_children():
            if isinstance(widget, ttk.Entry):
                widget.pack_configure(pady=(0, 2))
        
        bottom_frame = ttk.Frame(arrange_frame)
        bottom_frame.pack(fill=tk.X, padx=5, pady=2)
        
        row_frame = ttk.Frame(bottom_frame)
        row_frame.pack(side=tk.LEFT)
        ttk.Label(row_frame, text="每行窗口数").pack(anchor=tk.W)
        self.windows_per_row = ttk.Entry(row_frame, width=8, style="Small.TEntry")
        self.windows_per_row.pack(pady=(2, 0))
        self.windows_per_row.insert(0, "5")
        self.setup_right_click_menu(self.windows_per_row)
        
        ttk.Button(
            bottom_frame,
            text="自定义排列",
            command=self.custom_arrange_windows,
            style="Accent.TButton",
        ).pack(side=tk.RIGHT, pady=(15, 0))
        
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, padx=10, pady=(5, 0))
        
        self.tab_control = ttk.Notebook(bottom_frame)
        self.tab_control.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.create_open_window_tab()
        self.create_url_tab()
        self.create_tab_manage_tab()
        self.create_random_number_tab()
        self.create_env_create_tab()
        
        self.create_footer()
        
        self.load_arrange_params()
    
    def create_open_window_tab(self):
        open_window_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(open_window_tab, text="打开窗口")

        # 第一行：包含窗口编号、指定启动、随机启动、依次启动的所有控件
        controls_frame = ttk.Frame(open_window_tab)
        controls_frame.pack(fill=tk.X, padx=10, pady=(10,10)) 

        # 窗口编号部分
        ttk.Label(controls_frame, text="窗口编号:").pack(side=tk.LEFT)
        self.numbers_entry = ttk.Entry(controls_frame, width=7) # 宽度从 8 调整为 7
        self.numbers_entry.pack(side=tk.LEFT, padx=(5, 3)) 
        self.setup_right_click_menu(self.numbers_entry)

        if "last_window_numbers" in self.settings:
            self.numbers_entry.insert(0, self.settings["last_window_numbers"])

        self.numbers_entry.bind("<Return>", lambda e: self.open_windows())

        ttk.Button(
            controls_frame,
            text="指定启动",
            command=self.open_windows,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=(3, 3)) 

        # 随机启动部分
        ttk.Label(controls_frame, text="随机数量:").pack(side=tk.LEFT, padx=(3, 0)) 
        self.random_launch_count_entry = ttk.Entry(controls_frame, textvariable=self.random_launch_count_var, width=2) 
        self.random_launch_count_entry.pack(side=tk.LEFT, padx=(3,3)) 
        self.setup_right_click_menu(self.random_launch_count_entry)

        ttk.Button(
            controls_frame,
            text="随机启动",
            command=self.launch_random_windows,
            style="Accent.TButton"
        ).pack(side=tk.LEFT, padx=(3, 3)) 

        # 依次启动部分
        self.sequential_status_label = ttk.Label(controls_frame, text="依次启动状态: 未激活")
        self.sequential_status_label.pack(side=tk.LEFT, padx=(3,3)) 

        self.launch_sequentially_button = ttk.Button(
            controls_frame,
            text="依次启动", 
            command=self.launch_sequentially,
            style="Accent.TButton"
        )
        self.launch_sequentially_button.pack(side=tk.LEFT, padx=3) 
        
        ttk.Button(
            controls_frame,
            text="重置序列",
            command=self.reset_sequential_launch,
        ).pack(side=tk.LEFT, padx=3) 
        
        # 原来的 launch_options_frame 及其内部的 random_frame 和 sequential_frame 已被移除，
        # 其内容整合到了上方的 controls_frame 中。
    
    def create_url_tab(self):
        url_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(url_tab, text="批量打开网页")
        
        url_frame = ttk.Frame(url_tab)
        url_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(url_frame, text="网址:").pack(side=tk.LEFT)
        self.url_entry = ttk.Entry(url_frame, width=20)
        self.url_entry.pack(side=tk.LEFT, padx=5)
        self.url_entry.insert(0, "www.google.com")
        
        self.url_entry.bind("<Return>", lambda e: self.batch_open_urls())
        
        ttk.Button(
            url_frame,
            text="批量打开",
            command=self.batch_open_urls,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=5)
        
        try:
            self.selected_url = tk.StringVar()
            
            self.url_combobox = ttk.Combobox(
                url_frame, 
                textvariable=self.selected_url,
                width=20,
                height=12,
                state="readonly"
            )
            
            self.url_combobox.bind("<MouseWheel>", self._handle_url_combobox_scroll)  
            self.url_combobox.bind("<<ComboboxDropdown>>", self._setup_combobox_scrollbar)       
            self.url_combobox.bind("<<ComboboxSelected>>", self.on_url_selected)  
            self.url_combobox["values"] = []
            self.url_combobox.pack(side=tk.LEFT, padx=5)   
            self.update_url_combobox()
        except Exception as e:
            print(f"创建URL下拉菜单失败: {str(e)}")
        
        ttk.Button(
            url_frame,
            text="自定义网址",
            command=self.show_url_manager_dialog,
            width=10,
        ).pack(side=tk.LEFT, padx=5)
    
    def _setup_combobox_scrollbar(self, event):
        try:
            combobox = event.widget
            listbox_path = combobox.tk.call("ttk::combobox::PopdownWindow", combobox)
            
            if listbox_path:
                children = combobox.tk.splitlist(combobox.tk.call("winfo", "children", listbox_path))
                
                for child in children:
                    child_class = combobox.tk.call("winfo", "class", child)
                    if child_class == "Scrollbar":
                        combobox.tk.call("pack", "configure", child, "-side", "right", "-fill", "y")
                        combobox.tk.call("pack", "configure", child, "-padx", 0, "-pady", 0)
                        print(f"找到并配置了滚动条: {child}")
                    elif child_class == "Listbox":
                        combobox.tk.call(child, "configure", "-activestyle", "none")
                        combobox.tk.call(child, "configure", "-selectbackground", "#4a6984")
                        combobox.tk.call(child, "configure", "-selectforeground", "white")
                        print(f"配置了Listbox: {child}")
        except Exception as e:
            print(f"设置Combobox滚动条失败: {str(e)}")
    
    def _handle_url_combobox_scroll(self, event):
        try:
            if hasattr(self, 'url_combobox'):
                combobox = self.url_combobox
                if not combobox.winfo_ismapped() and combobox.current() >= 0:
                    current_index = combobox.current()
                    
                    direction = -1 if event.delta > 0 else 1
                    new_index = current_index + direction
                    
                    values = combobox["values"]
                    if values and 0 <= new_index < len(values):
                        combobox.current(new_index)
                        combobox.event_generate("<<ComboboxSelected>>")
                        
                    return "break"
        except Exception as e:
            print(f"处理URL下拉菜单滚轮事件失败: {str(e)}")
            
        return None
    
    def create_tab_manage_tab(self):
        tab_manage_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(tab_manage_tab, text="标签页管理")
        
        tab_manage_frame = ttk.Frame(tab_manage_tab)
        tab_manage_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            tab_manage_frame,
            text="仅保留当前标签页",
            command=self.keep_only_current_tab,
            width=20,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            tab_manage_frame,
            text="仅保留新标签页",
            command=self.keep_only_new_tab,
            width=20,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=5)
    
    def create_random_number_tab(self):
        random_number_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(random_number_tab, text="批量文本输入")
        
        buttons_frame = ttk.Frame(random_number_tab)
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            buttons_frame,
            text="随机数字输入",
            command=self.show_random_number_dialog,
            width=20,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(
            buttons_frame,
            text="指定文本输入",
            command=self.show_text_input_dialog,
            width=20,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=10)
    
    def create_env_create_tab(self):
        env_create_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(env_create_tab, text="批量创建环境")
        
        input_row = ttk.Frame(env_create_tab)
        input_row.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(input_row, text="创建编号:").pack(side=tk.LEFT)
        self.env_numbers = ttk.Entry(input_row, width=20)
        self.env_numbers.pack(side=tk.LEFT, padx=5)
        self.setup_right_click_menu(self.env_numbers)
        
        ttk.Button(
            input_row,
            text="开始创建",
            command=self.create_environments,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(input_row, text="示例: 1-5,7,9-12").pack(side=tk.LEFT, padx=5)
    
    def create_footer(self):
        footer_frame = ttk.Frame(self.root)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        
        donate_frame = ttk.Frame(footer_frame)
        donate_frame.pack(side=tk.LEFT)
        
        donate_label = ttk.Label(
            donate_frame,
            text="铸造一个看上去没什么用的NFT 0.1SOL（其实就是打赏啦 😁）",
            cursor="hand2",
            foreground="black",
        )
        donate_label.pack(side=tk.LEFT)
        donate_label.bind(
            "<Button-1>",
            lambda e: webbrowser.open("https://truffle.wtf/project/Devilflasher"),
        )
        
        author_frame = ttk.Frame(footer_frame)
        author_frame.pack(side=tk.RIGHT)
        
        ttk.Label(author_frame, text="Compiled by Devilflasher").pack(side=tk.LEFT)
        
        ttk.Label(author_frame, text="  ").pack(side=tk.LEFT)
        
        twitter_label = ttk.Label(
            author_frame, text="Twitter", cursor="hand2", font=("Arial", 9)
        )
        twitter_label.pack(side=tk.LEFT)
        twitter_label.bind(
            "<Button-1>", lambda e: webbrowser.open("https://x.com/DevilflasherX")
        )
        
        ttk.Label(author_frame, text="  ").pack(side=tk.LEFT)
        
        telegram_label = ttk.Label(
            author_frame, text="Telegram", cursor="hand2", font=("Arial", 9)
        )
        telegram_label.pack(side=tk.LEFT)
        telegram_label.bind(
            "<Button-1>", lambda e: webbrowser.open("https://t.me/devilflasher0")
        )
    
    def create_context_menus(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="剪切", command=self.cut_text)
        self.context_menu.add_command(label="复制", command=self.copy_text)
        self.context_menu.add_command(label="粘贴", command=self.paste_text)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="全选", command=self.select_all_text)
        
        self.window_list_menu = tk.Menu(self.root, tearoff=0)
        self.window_list_menu.add_command(
            label="关闭此窗口", command=self.close_selected_windows
        )
        self.window_list.bind("<Button-3>", self.show_window_list_menu)
        
        self.current_text_widget = None
    
    def show_context_menu(self, event):
        self.current_text_widget = event.widget
        self.context_menu.post(event.x_root, event.y_root)
    
    def cut_text(self):
        if self.current_text_widget:
            self.current_text_widget.event_generate("<<Cut>>")
    
    def copy_text(self):
        if self.current_text_widget:
            self.current_text_widget.event_generate("<<Copy>>")
    
    def paste_text(self):
        if self.current_text_widget:
            self.current_text_widget.event_generate("<<Paste>>")
    
    def select_all_text(self):
        if self.current_text_widget:
            self.current_text_widget.select_range(0, tk.END)
    
    def setup_right_click_menu(self, widget):
        widget.bind("<Button-3>", self.show_context_menu)
    
    def show_window_list_menu(self, event):
        item = self.window_list.identify_row(event.y)
        if item:
            self.window_list.selection_set(item)
            self.window_list_menu.post(event.x_root, event.y_root)
    
    def close_selected_windows(self):
        selected = []
        
        for item in self.window_list.get_children():
            if self.window_list.set(item, "select") == "√":
                values = self.window_list.item(item)["values"]
                hwnd = int(values[4])
                selected.append((item, hwnd))
        
        if not selected:
            messagebox.showinfo("提示", "请选择要关闭的窗口！")
            return
        
        try:
            self.manager.close_windows([hwnd for _, hwnd in selected])
            
            for item, _ in selected:
                self.window_list.delete(item)
            
            self.root.update()
            
            self.update_select_all_status()
            
            def check_after_close():
                if not self.window_list.get_children():
                    print("所有窗口已关闭，重置同步状态")
                    if hasattr(self.manager, "is_syncing") and self.manager.is_syncing:
                        try:
                            self.manager.stop_sync()
                            self.manager.is_syncing = False
                        except Exception as e:
                            print(f"停止同步失败: {str(e)}")
                    self.sync_button.configure(text="▶ 开始同步", style="Accent.TButton")
                    self.select_all_var.set("全部选择")
            
            self.root.after(100, check_after_close)
            
            if self.show_chrome_tip:
                self.show_chrome_settings_tip()
        
        except Exception as e:
            print(f"关闭窗口失败: {str(e)}")
    
    def delayed_initialization(self):
        try:
            self.root.deiconify()
            
            try:
                self.screens, self.screen_names = update_screen_list()
                
                if not self.screen_selection and self.screen_names:
                    self.screen_selection = self.screen_names[0]
                    self.settings["screen_selection"] = self.screen_selection
                    save_settings(self.settings)
            except Exception as e:
                log_error("更新屏幕列表失败", e)
            
            print(f"[{time.time() - self.start_time:.3f}s] 延迟初始化完成")
        except Exception as e:
            log_error("延迟初始化失败", e)
    
    def load_window_position(self):
        try:
            if "window_position" in self.settings:
                return self.settings["window_position"]
        except Exception as e:
            log_error("加载窗口位置失败", e)
        return ""
    
    def save_window_position(self):
        try:
            x = self.root.winfo_x()
            y = self.root.winfo_y()
            self.settings["window_position"] = f"+{x}+{y}"
            save_settings(self.settings)
        except Exception as e:
            log_error("保存窗口位置失败", e)
    
    def load_arrange_params(self):
        for param, entry in [
            ("start_x", self.start_x),
            ("start_y", self.start_y),
            ("window_width", self.window_width_entry),
            ("window_height", self.window_height_entry),
            ("h_spacing", self.h_spacing),
            ("v_spacing", self.v_spacing),
            ("windows_per_row", self.windows_per_row)
        ]:
            if param in self.settings:
                entry.delete(0, tk.END)
                entry.insert(0, self.settings[param])
    
    def get_arrange_params(self):
        return {
            "start_x": self.start_x.get(),
            "start_y": self.start_y.get(),
            "window_width": self.window_width_entry.get(),
            "window_height": self.window_height_entry.get(),
            "h_spacing": self.h_spacing.get(),
            "v_spacing": self.v_spacing.get(),
            "windows_per_row": self.windows_per_row.get()
        }
    
    def run(self):
        self.root.mainloop()
    
    def on_closing(self):
        try:
            self.save_window_position()
            
            self.save_custom_urls()
            
            if hasattr(self.manager, "is_syncing") and self.manager.is_syncing:
                try:
                    print("程序正在关闭，停止同步...")
                    self.manager.stop_sync()
                    self.manager.is_syncing = False
                except Exception as e:
                    print(f"关闭程序时停止同步失败: {str(e)}")
            
            save_settings(self.settings)
            
            self.root.destroy()
        
        except Exception as e:
            log_error("关闭程序时出错", e)
            self.root.destroy()
    
    def on_click(self, event):
        try:
            region = self.window_list.identify_region(event.x, event.y)
            if region == "cell":
                column = self.window_list.identify_column(event.x)
                item = self.window_list.identify_row(event.y)
                
                if column == "#1":  # 选择列
                    current = self.window_list.set(item, "select")
                    self.window_list.set(item, "select", "" if current == "√" else "√")
                    self.update_select_all_status()
                elif column == "#4":  # 主控列
                    self.set_master_window(item)
        except Exception as e:
            print(f"处理点击事件失败: {str(e)}")
    
    def toggle_select_all(self, event=None):
        try:
            items = self.window_list.get_children()
            if not items:
                return
            
            current_text = self.select_all_var.get()
            
            if current_text == "全部选择":
                for item in items:
                    self.window_list.set(item, "select", "√")
            else:
                for item in items:
                    self.window_list.set(item, "select", "")
            
            self.update_select_all_status()
        
        except Exception as e:
            print(f"切换全选状态失败: {str(e)}")
    
    def update_select_all_status(self):
        try:
            items = self.window_list.get_children()
            if not items:
                self.select_all_var.set("全部选择")
                return
            
            selected_count = sum(
                1 for item in items if self.window_list.set(item, "select") == "√"
            )
            
            if selected_count == len(items):
                self.select_all_var.set("取消全选")
            else:
                self.select_all_var.set("全部选择")
        
        except Exception as e:
            print(f"更新全选状态失败: {str(e)}")
    
    def set_master_window(self, item):
        try:
            values = self.window_list.item(item)["values"]
            hwnd = int(values[4])
            log_error(f"UI: Attempting to set master window. Item ID in Treeview: {item}, Target HWND: {hwnd}") # 添加日志

            if self.manager.is_syncing:
                log_error("UI: Master window change - stopping sync.") # 添加日志
                self.manager.stop_sync()
                self.sync_button.configure(text="▶ 开始同步", style="Accent.TButton")
                self.manager.is_syncing = False
            
            # Reset previous master window (if any) in UI and core
            for i in self.window_list.get_children():
                if i != item: # Don't reset the one we are about to set
                    # Get HWND of the other window
                    other_values = self.window_list.item(i)["values"]
                    if other_values and len(other_values) >= 5:
                        other_hwnd = int(other_values[4])
                        # Tell core to reset its style
                        log_error(f"UI: Resetting master style for previously mastered HWND: {other_hwnd}") # 添加日志
                        self.manager.reset_master_window(other_hwnd)
                        
                        # Update title in UI if it was changed by master status
                        # This part might be redundant if refresh_window_titles is called later,
                        # but good for immediate UI feedback if core's reset_master_window doesn't update title directly.
                        current_title_in_ui = self.window_list.set(i, "title")
                        actual_title_from_os = win32gui.GetWindowText(other_hwnd)
                        if current_title_in_ui != actual_title_from_os:
                             self.window_list.set(i, "title", actual_title_from_os)

                    # Clear master mark in UI
                    self.window_list.set(i, "master", "")
                    self.window_list.item(i, tags=())
            
            # Set new master window in core and UI
            log_error(f"UI: Calling self.manager.set_master_window for HWND: {hwnd}") # 添加日志
            success_core = self.manager.set_master_window(hwnd) # This calls core.py
            log_error(f"UI: self.manager.set_master_window returned: {success_core} for HWND: {hwnd}") # 添加日志

            if success_core:
                self.manager.master_window = hwnd # Update manager's master_window reference
                self.window_list.set(item, "master", "√")
                self.window_list.item(item, tags=("master",))
                log_error(f"UI: Successfully set master in UI for HWND: {hwnd}") # 添加日志
            else:
                log_error(f"UI: Core manager failed to set master window for HWND: {hwnd}. UI will not mark as master.") # 添加日志
                # Optionally, show a message to the user
                # messagebox.showerror("错误", f"设置主控窗口 (HWND: {hwnd}) 核心操作失败。")

            # Refresh all titles in the list to reflect any changes made by set_master_window or reset_master_window
            log_error("UI: Calling refresh_window_titles after attempting to set/reset master windows.") # 添加日志
            self.refresh_window_titles()
        
        except Exception as e:
            log_error(f"UI: Exception in set_master_window: {str(e)}", exc_info=True) # exc_info=True for full traceback
    
    def refresh_window_titles(self):
        try:
            for item in self.window_list.get_children():
                values = self.window_list.item(item)["values"]
                if values and len(values) >= 5:
                    hwnd = int(values[4])
                    title = win32gui.GetWindowText(hwnd)
                    self.window_list.set(item, "title", title)
        except Exception as e:
            log_error("刷新窗口标题失败", e)
    
    def import_windows(self):
        print(f"DEBUG: ChromeManagerUI.import_windows CALLED at {time.time()}")
        try:
            if hasattr(self.manager, "is_syncing"):
                if self.manager.is_syncing:
                    try:
                        self.manager.stop_sync()
                    except Exception as e:
                        print(f"停止同步失败: {str(e)}")
                self.manager.is_syncing = False
                self.sync_button.configure(text="▶ 开始同步", style="Accent.TButton")
                print("导入窗口前已重置同步状态")
            
            self.select_all_var.set("全部选择")
            
            if not self.manager.shortcut_path:
                messagebox.showinfo("提示", "请先在设置中设置快捷方式目录！")
                self.show_settings_dialog()
                return
            
            for item in self.window_list.get_children():
                self.window_list.delete(item)
            
            import_dialog = tk.Toplevel(self.root)
            import_dialog.title("导入窗口")
            import_dialog.geometry("300x100")
            import_dialog.withdraw()
            import_dialog.transient(self.root)
            import_dialog.resizable(False, False)
            import_dialog.grab_set()
            self.set_dialog_icon(import_dialog)
            center_window(import_dialog, self.root)
            import_dialog.deiconify()
            
            progress_label = ttk.Label(import_dialog, text="正在搜索Chrome窗口...")
            progress_label.pack(pady=10)
            
            progress = ttk.Progressbar(import_dialog, orient=tk.HORIZONTAL, length=250, mode="indeterminate")
            progress.pack(pady=5)
            progress.start()
            
            cancel_button = ttk.Button(import_dialog, text="取消", command=lambda: import_dialog.destroy())
            cancel_button.pack(pady=5)
            
            import_dialog.update()
            
            def import_thread():
                try:
                    self.manager.import_windows()

                    self.root.after(100, lambda: close_search_dialog(import_dialog))

                except Exception as e:
                    log_error("导入窗口线程发生错误", e)
                    if import_dialog and import_dialog.winfo_exists():
                        self.root.after(0, lambda: import_dialog.destroy())
                    self.root.after(0, lambda: messagebox.showerror("错误", f"导入时发生内部错误: {str(e)}"))
            
            def close_search_dialog(dialog_to_close):
                if dialog_to_close and dialog_to_close.winfo_exists():
                    dialog_to_close.destroy()
            
            threading.Thread(target=import_thread).start()
        
        except Exception as e:
            log_error("导入窗口失败", e)
            messagebox.showerror("错误", str(e))
    
    def update_window_list(self, windows, dialog=None, show_icon_progress=False):
        print(f"DEBUG: update_window_list called with {len(windows)} windows: {windows}")
        try:
            for item in self.window_list.get_children():
                self.window_list.delete(item)
            
            if not windows:
                self.select_all_var.set("全部选择")
                print(f"DEBUG: update_window_list finished (no windows). Treeview items: {self.window_list.get_children()}")
                return

            for window in windows:
                item = self.window_list.insert("", "end", values=[
                    "", window["number"], window["title"], "", window["hwnd"]
                ])
            
            if self.window_list.get_children():
                first_item = self.window_list.get_children()[0]
                self.window_list.set(first_item, "select", "√")
                self.update_select_all_status()

            print(f"DEBUG: update_window_list finished populating. Treeview items: {self.window_list.get_children()}")

            icon_progress_dialog = tk.Toplevel(self.root)
            icon_progress_dialog.title("图标处理")
            icon_progress_dialog.geometry("300x100")
            icon_progress_dialog.withdraw()
            icon_progress_dialog.transient(self.root)
            icon_progress_dialog.resizable(False, False)
            icon_progress_dialog.grab_set()
            self.set_dialog_icon(icon_progress_dialog)
            center_window(icon_progress_dialog, self.root)
            
            icon_progress_label = ttk.Label(icon_progress_dialog, text="正在准备图标替换...")
            icon_progress_label.pack(pady=10)
            icon_progress_bar = ttk.Progressbar(icon_progress_dialog, orient=tk.HORIZONTAL, length=250, mode="determinate", value=0)
            icon_progress_bar.pack(pady=5)
            
            icon_progress_dialog.deiconify()
            icon_progress_dialog.update()

            def start_icon_replacement_flow():
                def update_icon_dialog_progress(percent, text=None):
                    if icon_progress_dialog and icon_progress_dialog.winfo_exists():
                        icon_progress_bar.config(value=percent)
                        if text:
                            icon_progress_label.config(text=text)
                        icon_progress_dialog.update_idletasks()

                def actual_icon_replacement_task():
                    try:
                        pid_to_number_map = getattr(self.manager, 'pid_to_number', {})
                        if not pid_to_number_map:
                            temp_map = {}
                            for win_num, win_data in getattr(self.manager, 'windows', {}).items():
                                pid = win_data.get('pid')
                                if pid:
                                    temp_map[pid] = win_num
                            pid_to_number_map = temp_map

                        update_icon_dialog_progress(10, "正在生成图标...")
                        self.manager.apply_icons_to_chrome_windows(pid_to_number_map)

                        if hasattr(self.manager, "auto_modify_shortcut_icon") and self.manager.auto_modify_shortcut_icon:
                            self.root.after(1000, lambda: update_icon_dialog_progress(50, "正在更新快捷方式图标..."))
                        else:
                            self.root.after(1000, lambda: update_icon_dialog_progress(50, "已跳过快捷方式图标更新..."))
                        
                        self.root.after(2000, lambda: update_icon_dialog_progress(75, "正在替换任务栏图标..."))
                        self.root.after(3000, lambda: update_icon_dialog_progress(100, "图标替换完成!"))

                        self.root.after(3500, lambda: close_icon_dialog(icon_progress_dialog))
                    except Exception as e_icon:
                        log_error("图标替换任务失败", e_icon)
                        if icon_progress_dialog and icon_progress_dialog.winfo_exists():
                           icon_progress_label.config(text=f"图标处理失败: {e_icon}")
                        self.root.after(3000, lambda: close_icon_dialog(icon_progress_dialog))

                threading.Thread(target=actual_icon_replacement_task, daemon=True).start()

            def close_icon_dialog(dialog_to_close):
                if dialog_to_close and dialog_to_close.winfo_exists():
                    dialog_to_close.destroy()

            self.root.after(300, start_icon_replacement_flow)

        except Exception as e:
            log_error("更新窗口列表或启动图标流程失败", e)
        
    def open_windows(self):
        if hasattr(self.manager, "update_activity_timestamp"):
            self.manager.update_activity_timestamp()
            
        numbers_str = self.numbers_entry.get()
        
        if not numbers_str:
            messagebox.showwarning("警告", "请输入窗口编号！")
            return
        
        try:
            log_error(f"UI: 指定启动调用 manager.open_windows with numbers: {numbers_str}")
            success = self.manager.open_windows(numbers_str) # core.py 中的 open_windows
            if success:
                self.settings["last_window_numbers"] = numbers_str
                save_settings(self.settings)
                # 导入窗口以更新列表是个好主意，但 open_windows 本身不返回已启动的列表
                # 或许 open_windows 应该触发一个列表刷新
                # self.root.after(1500, self.import_windows) # 延迟一点导入，给窗口启动时间
                show_notification("操作已发送", f"已发送启动请求: {numbers_str[:50]}")

        except Exception as e:
            log_error("指定启动 - 打开窗口失败", e)
            messagebox.showerror("错误", str(e))

    def launch_random_windows(self):
        if hasattr(self.manager, "update_activity_timestamp"):
            self.manager.update_activity_timestamp()

        range_str = self.numbers_entry.get() # 复用窗口编号输入框作为范围
        count_str = self.random_launch_count_var.get()

        if not range_str:
            messagebox.showwarning("警告", "请输入窗口编号范围 (如 1-100)！")
            return
        if not count_str:
            messagebox.showwarning("警告", "请输入随机启动的数量！")
            return
        
        try:
            count = int(count_str)
            if count <= 0:
                messagebox.showwarning("警告", "随机启动数量必须大于0！")
                return
        except ValueError:
            messagebox.showerror("错误", "随机启动数量必须是有效的数字！")
            return

        if not self.manager.shortcut_path or not os.path.exists(self.manager.shortcut_path):
            messagebox.showerror("错误", "请先在设置中配置有效的快捷方式目录！")
            self.show_settings_dialog()
            return

        log_error(f"UI: 随机启动请求。范围: {range_str}, 数量: {count}")
        
        # 禁用按钮避免重复点击
        # (或者在 core.py 中实现更复杂的异步处理和状态反馈)
        # 这里简单地直接调用，依赖core的实现
        
        try:
            launched_numbers = self.manager.launch_random_profiles(range_str, count)
            if launched_numbers:
                msg = f"成功随机启动 {len(launched_numbers)} 个窗口: {', '.join(map(str, launched_numbers))}"
                log_error(f"UI: {msg}")
                show_notification("随机启动成功", msg)
                # self.root.after(1500, self.import_windows) # 更新列表
            elif count > 0 : # 如果请求启动的数量大于0但没有成功启动的
                msg = "未能随机启动任何新窗口 (可能范围内都已启动或无有效快捷方式)。"
                log_error(f"UI: {msg}")
                messagebox.showinfo("提示", msg)
            # 如果 count 是0，上面已经处理了

        except ValueError as ve: # 来自 parse_window_numbers 的错误
            log_error(f"UI: 随机启动范围解析错误: {ve}")
            messagebox.showerror("范围错误", str(ve))
        except Exception as e:
            log_error(f"UI: 随机启动失败 - {str(e)}", exc_info=True)
            messagebox.showerror("错误", f"随机启动失败: {str(e)}")

    def launch_sequentially(self):
        if hasattr(self.manager, "update_activity_timestamp"):
            self.manager.update_activity_timestamp()

        current_range_str = self.numbers_entry.get().strip()
        if not current_range_str:
            messagebox.showwarning("警告", "请输入窗口编号范围 (如 1-10) 以进行依次启动！")
            return

        if not self.manager.shortcut_path or not os.path.exists(self.manager.shortcut_path):
            messagebox.showerror("错误", "请先在设置中配置有效的快捷方式目录！")
            self.show_settings_dialog()
            return

        # 初始化或重新初始化序列
        if not self.sequential_launch_active or self.sequential_launch_range_str != current_range_str:
            log_error(f"UI: 初始化/重新初始化依次启动序列。范围: {current_range_str}")
            try:
                # 让 core 来解析范围并获取有效、未启动的配置文件列表
                # 我们需要一个新的 core 方法来获取这个列表
                profiles_in_range = self.manager.get_valid_profiles_for_sequential_launch(current_range_str)
                
                if not profiles_in_range:
                    msg = f"在范围 '{current_range_str}' 内没有找到有效或未启动的快捷方式。"
                    self.sequential_status_label.config(text=f"序列: {current_range_str} (无可用)")
                    log_error(f"UI: {msg}")
                    messagebox.showinfo("提示", msg)
                    self.reset_sequential_launch() # 清理状态
                    return

                self.sequential_launch_profiles = profiles_in_range
                self.sequential_launch_current_index = 0
                self.sequential_launch_active = True
                self.sequential_launch_range_str = current_range_str
                self.last_launched_sequentially = None
                
                next_to_launch = self.sequential_launch_profiles[0]
                status_msg = f"序列激活: {current_range_str}. 下一个: {next_to_launch}. 共 {len(self.sequential_launch_profiles)} 个待启动."
                self.sequential_status_label.config(text=status_msg)
                log_error(f"UI: {status_msg}")

            except ValueError as ve: # 来自解析数字范围的错误
                log_error(f"UI: 依次启动范围解析错误: {ve}")
                messagebox.showerror("范围错误", str(ve))
                self.reset_sequential_launch()
                return
            except Exception as e:
                log_error(f"UI: 依次启动序列初始化失败: {str(e)}", exc_info=True)
                messagebox.showerror("错误", f"序列初始化失败: {str(e)}")
                self.reset_sequential_launch()
                return
        
        # 如果序列已激活，执行启动下一个
        if not self.sequential_launch_active:
             messagebox.showinfo("提示", "依次启动序列未激活或已完成。请使用有效范围重新开始或重置序列。")
             return

        # 跳过已在 self.manager.windows 中的（以防万一在序列进行中手动打开了）
        # Core 的 get_valid_profiles_for_sequential_launch 应该已经处理了这个
        # 但作为双重保险，或如果 core 方法没那么智能
        while self.sequential_launch_current_index < len(self.sequential_launch_profiles):
            profile_to_check = self.sequential_launch_profiles[self.sequential_launch_current_index]
            # 需要一个方法从 core manager 获取某个profile是否已启动
            if self.manager.is_profile_running(profile_to_check): # 假设 core.py 有此方法
                log_error(f"UI: 依次启动 - 编号 {profile_to_check} 已在运行，跳过。")
                self.sequential_launch_current_index += 1
            else:
                break 
        
        if self.sequential_launch_current_index >= len(self.sequential_launch_profiles):
            msg = f"依次启动序列 '{self.sequential_launch_range_str}' 已全部完成或剩余均已启动。"
            self.sequential_status_label.config(text=msg)
            log_error(f"UI: {msg}")
            show_notification("序列完成", msg)
            self.reset_sequential_launch(notify=False)
            return

        profile_to_launch = self.sequential_launch_profiles[self.sequential_launch_current_index]
        self.last_launched_sequentially = profile_to_launch # 记录，用于回调

        status_msg = f"序列: {self.sequential_launch_range_str}. 尝试启动: {profile_to_launch} ({self.sequential_launch_current_index + 1}/{len(self.sequential_launch_profiles)})"
        self.sequential_status_label.config(text=status_msg)
        log_error(f"UI: {status_msg}")

        try:
            # 直接让 core 启动这一个，core 的 open_windows 接受字符串编号
            success = self.manager.open_windows(str(profile_to_launch))
            self.on_sequential_item_launched(success, profile_to_launch)
        except Exception as e:
            log_error(f"UI: 依次启动 - 调用 manager.open_windows 失败 for {profile_to_launch}: {str(e)}", exc_info=True)
            self.on_sequential_item_launched(False, f"启动 {profile_to_launch} 失败: {str(e)}")

    def on_sequential_item_launched(self, success: bool, launched_info: any):
        """依次启动中，单个项目启动尝试完成后的回调。
        launched_info: 如果成功，是启动的编号(int)；如果失败，是错误信息(str)。
        """
        if not self.sequential_launch_active:
            log_error("UI: on_sequential_item_launched called but sequence not active. Ignoring.")
            return

        if success:
            launched_number = int(launched_info) # launched_info is profile_to_launch (number)
            log_error(f"UI: 依次启动 - 编号 {launched_number} 成功启动。")
            self.sequential_launch_current_index += 1
            
            if self.sequential_launch_current_index >= len(self.sequential_launch_profiles):
                msg = f"依次启动序列 '{self.sequential_launch_range_str}' 已成功完成所有启动。"
                self.sequential_status_label.config(text=msg)
                log_error(f"UI: {msg}")
                show_notification("序列完成", msg)
                self.reset_sequential_launch(notify=False)
            else:
                next_to_launch = self.sequential_launch_profiles[self.sequential_launch_current_index]
                remaining_count = len(self.sequential_launch_profiles) - self.sequential_launch_current_index
                status_msg = f"序列: {self.sequential_launch_range_str}. {launched_number}成功. 下一个: {next_to_launch}. 还剩 {remaining_count}."
                self.sequential_status_label.config(text=status_msg)
                log_error(f"UI: {status_msg}")
        else:
            # launched_info is error_msg string
            error_detail = str(launched_info)
            failed_number = self.last_launched_sequentially # 获取之前尝试启动的编号
            msg = f"序列: {self.sequential_launch_range_str}. 启动 {failed_number} 失败. 原因: {error_detail[:100]}"
            self.sequential_status_label.config(text=msg)
            log_error(f"UI: {msg}")
            # 修改错误提示框中的按钮文本
            messagebox.showerror("依次启动失败", f"启动编号 {failed_number} 失败。\n{error_detail}\n\n您可以再次点击'依次启动'按钮重试当前失败的，或修改范围后重置序列。")
            # 索引不增加，以便下次点击按钮时重试同一个
        
        # self.root.after(1000, self.import_windows) # 更新列表，可选

    def reset_sequential_launch(self, notify=True):
        self.sequential_launch_active = False
        self.sequential_launch_range_str = None
        self.sequential_launch_profiles = []
        self.sequential_launch_current_index = 0
        self.last_launched_sequentially = None
        self.sequential_status_label.config(text="依次启动状态: 未激活")
        if notify:
            log_error("UI: 依次启动序列已重置。")
            show_notification("序列已重置", "依次启动序列已重置。")

    def close_all_windows(self):
        if hasattr(self.manager, "update_activity_timestamp"):
            self.manager.update_activity_timestamp()
            
        window_handles = []
        
        for item in self.window_list.get_children():
            values = self.window_list.item(item)["values"]
            hwnd = int(values[4])
            window_handles.append(hwnd)
        
        if window_handles:
            self.manager.close_windows(window_handles)
    
    def toggle_sync(self, force_enable=None):
        if hasattr(self.manager, "update_activity_timestamp"):
            self.manager.update_activity_timestamp()
            
        try:
            if not self.window_list.get_children():
                print("没有可同步的窗口")
                messagebox.showinfo("提示", "请先导入窗口！")
                if hasattr(self.manager, "is_syncing") and self.manager.is_syncing:
                    print("重置同步状态")
                    try:
                        self.manager.stop_sync()
                    except Exception:
                        pass
                    self.manager.is_syncing = False
                self.sync_button.configure(text="▶ 开始同步", style="Accent.TButton")
                self.select_all_var.set("全部选择")
                return
            
            selected = []
            for item in self.window_list.get_children():
                if self.window_list.set(item, "select") == "√":
                    selected.append(item)

            if not selected:
                messagebox.showinfo("提示", "请选择要同步的窗口！")
                return

            master_items = [
                item
                for item in self.window_list.get_children()
                if self.window_list.set(item, "master") == "√"
            ]
            
            if not master_items:
                self.set_master_window(selected[0])
            
            is_syncing = self.manager.is_syncing if hasattr(self.manager, "is_syncing") else False
            
            if force_enable is not None:
                should_sync = force_enable
            else:
                should_sync = not is_syncing
                
            if should_sync and not is_syncing:
                try:
                    success = self.manager.start_sync(selected)
                    if success:
                        self.sync_button.configure(text="■ 停止同步", style="Accent.TButton")
                        print("同步已开启")
                        self.root.after(10, lambda: show_notification("同步已开启", "Chrome多窗口同步功能已启动"))
                    else:
                        messagebox.showerror("错误", "启动同步失败")
                except Exception as e:
                    messagebox.showerror("错误", f"启动同步失败: {str(e)}")
                    print(f"启动同步失败: {str(e)}")
                    traceback.print_exc()
            elif not should_sync and is_syncing:
                try:
                    success = self.manager.stop_sync()
                    if success:
                        self.sync_button.configure(text="▶ 开始同步", style="Accent.TButton")
                        print("同步已停止")
                    else:
                        messagebox.showerror("错误", "停止同步失败")
                except Exception as e:
                    messagebox.showerror("错误", f"停止同步失败: {str(e)}")
                    print(f"停止同步失败: {str(e)}")
                    traceback.print_exc()
        except Exception as e:
            print(f"切换同步状态失败: {str(e)}")
            messagebox.showerror("错误", f"切换同步状态失败: {str(e)}")
            traceback.print_exc()
    
    def auto_arrange_windows(self):
        if hasattr(self.manager, "update_activity_timestamp"):
            self.manager.update_activity_timestamp()
            
        try:
            print("开始自动排列窗口...")
            was_syncing = False
            if hasattr(self.manager, "is_syncing") and self.manager.is_syncing:
                was_syncing = self.manager.is_syncing
                self.manager.stop_sync()

            selected = []
            for item in self.window_list.get_children():
                if self.window_list.set(item, "select") == "√":
                    values = self.window_list.item(item)["values"]
                    if values and len(values) >= 5:
                        number = int(values[1])
                        hwnd = int(values[4])
                        selected.append((number, hwnd, item))

            if not selected:
                messagebox.showinfo("提示", "请先选择要排列的窗口！")
                return

            print(f"选中了 {len(selected)} 个窗口")

            selected.sort(key=lambda x: x[0])
            print("窗口排序结果:")
            for num, hwnd, _ in selected:
                print(f"编号: {num}, 句柄: {hwnd}")

            self.manager.screen_selection = self.screen_selection
            print(f"UI当前选择的屏幕: {self.screen_selection}")
            
            success = self.manager.auto_arrange_windows(selected)
            
            if not success:
                messagebox.showerror("错误", "自动排列窗口失败！")
                return

            hwnd_list = [hwnd for _, hwnd, _ in selected]
            self.manager.set_window_priority(hwnd_list)

            master_hwnd = None
            for item in self.window_list.get_children():
                if self.window_list.set(item, "master") == "√":
                    values = self.window_list.item(item)["values"]
                    if values and len(values) >= 5:
                        master_hwnd = int(values[4])
                        break

            if master_hwnd:
                self.manager.activate_window(master_hwnd)

            if was_syncing:
                self.toggle_sync(force_enable=True)

            print("窗口排列完成")

        except Exception as e:
            print(f"自动排列失败: {str(e)}")
            messagebox.showerror("错误", f"自动排列失败: {str(e)}")
            traceback.print_exc()
    
    def custom_arrange_windows(self):
        """自定义排列窗口"""
        if hasattr(self.manager, "update_activity_timestamp"):
            self.manager.update_activity_timestamp()
            
        try:
            print("开始自定义排列窗口...")
            was_syncing = False
            if hasattr(self.manager, "is_syncing") and self.manager.is_syncing:
                was_syncing = self.manager.is_syncing
                self.manager.stop_sync()

            windows = []
            for item in self.window_list.get_children():
                if self.window_list.set(item, "select") == "√":
                    hwnd = int(self.window_list.set(item, "hwnd"))
                    windows.append((item, hwnd))

            if not windows:
                messagebox.showinfo("提示", "请选择要排列的窗口！")
                return

            print(f"选中了 {len(windows)} 个窗口")

            try:
                arrange_params = self.get_arrange_params()
                start_x = int(arrange_params.get("start_x", 0))
                start_y = int(arrange_params.get("start_y", 0))
                width = int(arrange_params.get("window_width", 500))
                height = int(arrange_params.get("window_height", 400))
                h_spacing = int(arrange_params.get("h_spacing", 0))
                v_spacing = int(arrange_params.get("v_spacing", 0))
                windows_per_row = int(arrange_params.get("windows_per_row", 5))
                    
                print(f"排列参数: 起始位置=({start_x}, {start_y}), 大小={width}x{height}, 间距=({h_spacing}, {v_spacing}), 每行窗口数={windows_per_row}")
                
                has_multi_screen_config = False
                screen_configs = self.settings.get("screen_arrange_config", [])
                if screen_configs:
                    has_multi_screen_config = any(config["enabled"] for config in screen_configs)
                
                print(f"是否启用多屏幕配置: {has_multi_screen_config}")
                
                self.manager.screen_selection = self.screen_selection
                print(f"UI当前选择的屏幕: {self.screen_selection}")
                ordered_windows = []
                for item, hwnd in windows:
                    ordered_windows.append((item, hwnd))
                
                success = False
                
                if has_multi_screen_config:
                    active_screens = self.manager.get_active_screens()
                    
                    if active_screens:
                        print(f"使用多屏幕配置，找到 {len(active_screens)} 个活跃屏幕")
                        success = self.manager.custom_arrange_on_multiple_screens(
                            ordered_windows, 
                            active_screens, 
                            start_x, 
                            start_y, 
                            width, 
                            height, 
                            h_spacing, 
                            v_spacing, 
                            windows_per_row
                        )
                    else:
                        print("未找到活跃屏幕，回退到单屏幕排列方式")
                        screen_index = 0
                        for i, name in enumerate(self.screen_names):
                            if name == self.screen_selection:
                                screen_index = i
                                break
                        
                        print(f"使用屏幕 {screen_index}: {self.screen_selection}")
                        success = self.manager.custom_arrange_on_single_screen(
                            ordered_windows, 
                            screen_index, 
                            start_x, 
                            start_y, 
                            width, 
                            height, 
                            h_spacing, 
                            v_spacing, 
                            windows_per_row
                        )
                else:
                    print("未使用多屏幕配置，使用单屏幕排列")
                    screen_index = 0
                    for i, name in enumerate(self.screen_names):
                        if name == self.screen_selection:
                            screen_index = i
                            break
                            
                    print(f"使用屏幕 {screen_index}: {self.screen_selection}")
                    success = self.manager.custom_arrange_on_single_screen(
                        ordered_windows, 
                        screen_index, 
                        start_x, 
                        start_y, 
                        width, 
                        height, 
                        h_spacing, 
                        v_spacing, 
                        windows_per_row
                    )
                
                if success:
                    hwnd_list = [hwnd for _, hwnd in ordered_windows]
                    self.manager.set_window_priority(hwnd_list)
                    
                    master_hwnd = None
                    for item in self.window_list.get_children():
                        if self.window_list.set(item, "master") == "√":
                            master_hwnd = int(self.window_list.set(item, "hwnd"))
                            break
                            
                    if master_hwnd:
                        self.manager.activate_window(master_hwnd)
                        
                    if was_syncing:
                        self.toggle_sync(force_enable=True)
                        
                    print("自定义排列窗口成功")
                else:
                    print("自定义排列窗口失败")
                    messagebox.showerror("错误", "排列窗口失败！请检查参数设置。")
                
                save_settings(self.settings)
                
            except ValueError as e:
                print(f"参数错误: {str(e)}")
                messagebox.showerror("错误", "请输入有效的数字参数！")
                return

        except Exception as e:
            messagebox.showerror("错误", f"排列窗口过程中发生错误: {str(e)}")
            print(f"排列窗口过程中发生错误: {str(e)}")
            traceback.print_exc()
    
    def get_active_screens(self):
        """获取活跃的屏幕列表，按优先级排序"""
        screen_configs = self.settings.get("screen_arrange_config", [])
        
        if not screen_configs:
            screen_configs = []
            for i in range(len(self.screens)):
                screen_configs.append({
                    "screen_id": i,
                    "enabled": True,
                    "priority": i
                })
        
        active_screens = []
        
        sorted_configs = sorted(screen_configs, key=lambda x: x["priority"])
        
        for config in sorted_configs:
            if config["enabled"] and config["screen_id"] < len(self.screens):
                active_screens.append(self.screens[config["screen_id"]])
        
        return active_screens

    def custom_arrange_on_single_screen(self, windows, start_x, start_y, width, height, h_spacing, v_spacing, windows_per_row):
        """在单个屏幕上进行自定义排列"""
        screen_index = 0
        for i, name in enumerate(self.screen_names):
            if name == self.screen_selection:
                screen_index = i
                break
        
        return self.manager.custom_arrange_on_single_screen(
            windows, 
            screen_index, 
            start_x, 
            start_y, 
            width, 
            height, 
            h_spacing, 
            v_spacing, 
            windows_per_row
        )

    def custom_arrange_on_multiple_screens(self, windows, active_screens, start_x, start_y, width, height, h_spacing, v_spacing, windows_per_row):
        return self.manager.custom_arrange_on_multiple_screens(
            windows, 
            active_screens, 
            start_x, 
            start_y, 
            width, 
            height, 
            h_spacing, 
            v_spacing, 
            windows_per_row
        )

    def auto_arrange_multi_screens(self, selected_windows):
        return self.manager.auto_arrange_multi_screens(selected_windows)
    
    def batch_open_urls(self):
        if hasattr(self.manager, "update_activity_timestamp"):
            self.manager.update_activity_timestamp()
            
        url = self.url_entry.get()
        if not url:
            messagebox.showwarning("警告", "请输入网址！")
            return
        
        numbers = []
        for item in self.window_list.get_children():
            if self.window_list.set(item, "select") == "√":
                values = self.window_list.item(item)["values"]
                number = int(values[1])
                numbers.append(number)
        
        if not numbers:
            messagebox.showinfo("提示", "请先选择要打开网址的窗口！")
            return
        
        numbers_str = ",".join(str(n) for n in numbers)
        
        try:
            self.settings["last_used_url"] = url
            save_settings(self.settings)
            
            self.manager.batch_open_urls(url, numbers_str)
        except Exception as e:
            print(f"批量打开URL失败: {url} - {str(e)}")
            messagebox.showerror("错误", str(e))
    
    def update_url_combobox(self):
        try:
            if hasattr(self, "url_combobox"):
                if not isinstance(self.custom_urls, dict):
                    self.custom_urls = {}
                
                display_items = []
                self.url_mapping = {}
                
                if not self.custom_urls:
                    display_items.append("暂未录入信息")
                    self.url_mapping["暂未录入信息"] = ""
                else:
                    for url, title in self.custom_urls.items():
                        display_text = title if title and title != url else url
                        display_items.append(display_text)
                        self.url_mapping[display_text] = url
                
                self.url_combobox["values"] = display_items
                
                last_used_url = self.settings.get("last_used_url", "")
                last_used_display = None
                
                if last_used_url:
                    for display_text, url in self.url_mapping.items():
                        if url == last_used_url:
                            last_used_display = display_text
                            break
                
                if display_items:
                    if last_used_display and last_used_display in display_items:
                        self.url_combobox.set(last_used_display)
                        self.selected_url.set(last_used_display)
                        url_to_use = self.url_mapping.get(last_used_display, "")
                        if url_to_use and hasattr(self, "url_entry"):
                            self.url_entry.delete(0, tk.END)
                            self.url_entry.insert(0, url_to_use)
                    else:
                        self.url_combobox.current(0)
                        self.selected_url.set(display_items[0])
                        url_to_use = self.url_mapping.get(display_items[0], "")
                        if url_to_use and hasattr(self, "url_entry"):
                            self.url_entry.delete(0, tk.END)
                            self.url_entry.insert(0, url_to_use)
        except Exception as e:
            print(f"更新URL下拉菜单失败: {str(e)}")
    
    def on_url_selected(self, event):
        try:
            selected_text = self.selected_url.get()
            if selected_text and hasattr(self, "url_mapping"):
                if selected_text == "点击右侧按钮录入自定义信息":
                    self.root.after(100, self.show_url_manager_dialog)
                    return
                
                actual_url = self.url_mapping.get(selected_text, selected_text)
                if actual_url:
                    self.url_entry.delete(0, tk.END)
                    self.url_entry.insert(0, actual_url)
        except Exception as e:
            print(f"处理URL选择失败: {str(e)}")
    
    def save_custom_urls(self):
        try:
            if not isinstance(self.custom_urls, dict):
                if isinstance(self.custom_urls, list):
                    self.custom_urls = {url: url for url in self.custom_urls}
                else:
                    self.custom_urls = {}
            
            self.settings["custom_urls"] = self.custom_urls
            save_settings(self.settings)
        except Exception as e:
            print(f"保存自定义网址失败: {str(e)}")
    
    def show_url_manager_dialog(self):
        try:
            dialog = tk.Toplevel(self.root)
            dialog.title("自定义网址管理")
            dialog.geometry("500x400")
            dialog.resizable(False, False)

            dialog.transient(self.root)
            dialog.grab_set()

            self.set_dialog_icon(dialog)

            center_window(dialog, self.root)
            
            if not hasattr(self, "custom_urls") or not isinstance(self.custom_urls, dict):
                if isinstance(self.custom_urls, list):
                    self.custom_urls = {url: url for url in self.custom_urls}
                else:
                    self.custom_urls = {}
            
            temp_urls = self.custom_urls.copy()
            
            main_frame = ttk.Frame(dialog, padding=10)
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            input_frame = ttk.LabelFrame(main_frame, text="添加网址", padding=10)
            input_frame.pack(fill=tk.X, pady=(0, 5))
            
            title_frame = ttk.Frame(input_frame)
            title_frame.pack(fill=tk.X, pady=(0, 5))
            ttk.Label(title_frame, text="标题:").pack(side=tk.LEFT)
            title_var = tk.StringVar()
            title_entry = ttk.Entry(title_frame, textvariable=title_var, width=45)
            title_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            url_frame = ttk.Frame(input_frame)
            url_frame.pack(fill=tk.X)
            ttk.Label(url_frame, text="网址:").pack(side=tk.LEFT)
            url_var = tk.StringVar()
            url_entry = ttk.Entry(url_frame, textvariable=url_var, width=45)
            url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            url_entry.focus_set()
            
            ttk.Button(url_frame, text="添加", command=lambda: add_url(), style="Accent.TButton").pack(side=tk.RIGHT)
            
            self.setup_right_click_menu(title_entry)
            self.setup_right_click_menu(url_entry)
            
            list_frame = ttk.LabelFrame(main_frame, text="已添加的网址", padding=10)
            list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
            
            list_container = ttk.Frame(list_frame, height=150)
            list_container.pack(fill=tk.BOTH, expand=False)
            list_container.pack_propagate(False)
            
            scrollbar = ttk.Scrollbar(list_container)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            columns = ("title", "url")
            url_tree = ttk.Treeview(
                list_container,
                columns=columns,
                show="headings",
                selectmode="browse",
                yscrollcommand=scrollbar.set
            )
            
            url_tree.heading("title", text="标题")
            url_tree.heading("url", text="网址")
            
            url_tree.column("title", width=150)
            url_tree.column("url", width=300)
            
            url_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=url_tree.yview)
            
            for url, title in temp_urls.items():
                display_title = title if title != url else ""
                url_tree.insert("", tk.END, values=(display_title, url))
            
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill=tk.X, pady=5)
            
            left_buttons = ttk.Frame(button_frame)
            left_buttons.pack(side=tk.LEFT)
            
            ttk.Button(left_buttons, text="删除选中", command=lambda: delete_url()).pack(side=tk.LEFT, padx=5)
            ttk.Button(left_buttons, text="编辑选中", command=lambda: edit_selected()).pack(side=tk.LEFT, padx=5)
            
            right_buttons = ttk.Frame(button_frame)
            right_buttons.pack(side=tk.RIGHT)
            
            ttk.Button(right_buttons, text="保存", command=lambda: save_and_close(), style="Accent.TButton").pack(side=tk.LEFT, padx=5)
            ttk.Button(right_buttons, text="取消", command=lambda: cancel()).pack(side=tk.LEFT, padx=5)
            
            def add_url():
                url = url_var.get().strip()
                title = title_var.get().strip()
                
                if not url:
                    messagebox.showinfo("提示", "请输入网址！")
                    return
                
                if not url.startswith(("http://", "https://", "www.")):
                    url = "https://" + url
                
                if not title:
                    title = url
                
                if url in temp_urls:
                    messagebox.showinfo("提示", "该网址已存在！")
                    return
                
                temp_urls[url] = title
                
                display_title = title if title != url else ""
                url_tree.insert("", tk.END, values=(display_title, url))
                
                url_var.set("")
                title_var.set("")
                
                title_entry.focus_set()
            
            def delete_url():
                selected = url_tree.selection()
                if not selected:
                    messagebox.showinfo("提示", "请先选择要删除的网址！")
                    return
                
                try:
                    item = selected[0]
                    values = url_tree.item(item)["values"]
                    url = values[1]
                    
                    url_tree.delete(item)
                    if url in temp_urls:
                        del temp_urls[url]
                except Exception as e:
                    print(f"删除网址失败: {str(e)}")
                    messagebox.showerror("错误", f"删除网址失败: {str(e)}")
            
            def edit_selected():
                selected = url_tree.selection()
                if not selected:
                    messagebox.showinfo("提示", "请先选择要编辑的网址！")
                    return
                
                try:
                    item = selected[0]
                    values = url_tree.item(item)["values"]
                    old_title = values[0]
                    old_url = values[1]
                    
                    title_var.set(old_title)
                    url_var.set(old_url)
                    
                    url_tree.delete(item)
                    if old_url in temp_urls:
                        del temp_urls[old_url]
                    
                    title_entry.focus_set()
                except Exception as e:
                    print(f"编辑网址失败: {str(e)}")
                    messagebox.showerror("错误", f"编辑网址失败: {str(e)}")
            
            def save_and_close():  
                self.custom_urls = temp_urls.copy()
                self.save_custom_urls()
                self.update_url_combobox()
                dialog.destroy()
            
            def cancel():
                dialog.destroy()
            
            def on_enter(event):
                add_url()
            
            title_entry.bind("<Return>", lambda e: url_entry.focus_set())
            url_entry.bind("<Return>", on_enter)
            
            url_tree.bind("<Double-1>", lambda e: edit_selected())
            
            dialog.protocol("WM_DELETE_WINDOW", cancel)
            
            dialog.wait_window()
        except Exception as e:
            print(f"显示网址管理对话框失败: {str(e)}")
            messagebox.showerror("错误", str(e))
    
    def keep_only_current_tab(self):
        selected = self.get_selected_windows()
        
        if not selected:
            messagebox.showinfo("提示", "请先选择窗口！")
            return
        
        self.root.config(cursor="wait")
        
        try:
            success = self.manager.keep_only_current_tab(selected)
            if not success:
                messagebox.showerror("错误", "操作失败，请检查Chrome是否已启动")
        except Exception as e:
            print(f"仅保留当前标签页失败: {str(e)}")
            messagebox.showerror("错误", str(e))
        finally:
            self.root.after(1000, lambda: self.root.config(cursor=""))
    
    def keep_only_new_tab(self):
        selected = self.get_selected_windows()
        
        if not selected:
            messagebox.showinfo("提示", "请先选择窗口！")
            return
        
        self.root.config(cursor="wait")
        
        try:
            success = self.manager.keep_only_new_tab(selected)
            if not success:
                messagebox.showerror("错误", "操作失败，请检查Chrome是否已启动")
        except Exception as e:
            print(f"仅保留新标签页失败: {str(e)}")
            messagebox.showerror("错误", str(e))
        finally:
            self.root.after(1000, lambda: self.root.config(cursor=""))
    
    def show_random_number_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("随机数字输入")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        self.set_dialog_icon(dialog)

        center_window(dialog, self.root)

        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        range_frame = ttk.LabelFrame(main_frame, text="数字范围", padding=10)
        range_frame.pack(fill=tk.X, pady=(0, 10))

        range_inner_frame = ttk.Frame(range_frame)
        range_inner_frame.pack(fill=tk.X)
        ttk.Label(range_inner_frame, text="最小值:").pack(side=tk.LEFT)
        min_entry = ttk.Entry(
            range_inner_frame, width=10, textvariable=self.random_min_value
        )
        min_entry.pack(side=tk.LEFT, padx=(5, 15))
        self.setup_right_click_menu(min_entry)

        ttk.Label(range_inner_frame, text="最大值:").pack(side=tk.LEFT)
        max_entry = ttk.Entry(
            range_inner_frame, width=10, textvariable=self.random_max_value
        )
        max_entry.pack(side=tk.LEFT, padx=5)
        self.setup_right_click_menu(max_entry)

        options_frame = ttk.LabelFrame(main_frame, text="输入选项", padding=10)
        options_frame.pack(fill=tk.X, pady=(0, 15))

        options_inner_frame = ttk.Frame(options_frame)
        options_inner_frame.pack(fill=tk.X)

        overwrite_row = ttk.Frame(options_inner_frame)
        overwrite_row.pack(fill=tk.X, anchor=tk.W, pady=5)
        
        overwrite_check = ttk.Checkbutton(
            overwrite_row,
            variable=self.random_overwrite
        )
        overwrite_check.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(
            overwrite_row,
            text="覆盖原有内容"
        ).pack(side=tk.LEFT)

        delayed_row = ttk.Frame(options_inner_frame)
        delayed_row.pack(fill=tk.X, anchor=tk.W)
        
        delayed_check = ttk.Checkbutton(
            delayed_row,
            variable=self.random_delayed
        )
        delayed_check.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(
            delayed_row,
            text="模拟人工输入（逐字输入并添加延迟）"
        ).pack(side=tk.LEFT)

        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X)

        ttk.Button(buttons_frame, text="取消", command=dialog.destroy, width=10).pack(
            side=tk.RIGHT, padx=5
        )

        ttk.Button(
            buttons_frame,
            text="开始输入",
            command=lambda: self.run_random_input(dialog),
            style="Accent.TButton",
            width=10,
        ).pack(side=tk.RIGHT, padx=5)

    def run_random_input(self, dialog):
        dialog.destroy()
        self.input_random_number()

    def input_random_number(self):
        try:
            selected = self.get_selected_windows()
            
            if not selected:
                messagebox.showwarning("警告", "请先选择要操作的窗口！")
                return

            min_str = self.random_min_value.get().strip()
            max_str = self.random_max_value.get().strip()

            if not min_str or not max_str:
                messagebox.showwarning("警告", "请输入有效的范围值！")
                return

            is_float = "." in min_str or "." in max_str
            decimal_places = 2

            try:
                if is_float:
                    min_val = float(min_str)
                    max_val = float(max_str)
                    decimal_places = max(
                        len(min_str.split(".")[-1]) if "." in min_str else 0,
                        len(max_str.split(".")[-1]) if "." in max_str else 0,
                    )
                    decimal_places = min(decimal_places, 10)
                else:
                    min_val = int(min_str)
                    max_val = int(max_str)
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字范围！")
                return

            overwrite = self.random_overwrite.get()
            delayed = self.random_delayed.get()

            window_handles = []
            for item in selected:
                hwnd = int(self.get_window_item_value(item, "hwnd"))
                window_handles.append(hwnd)
            
            success = input_random_number(
                window_handles=window_handles,
                min_val=min_val,
                max_val=max_val,
                is_float=is_float,
                decimal_places=decimal_places,
                overwrite=overwrite,
                delayed=delayed
            )
            
            if not success:
                messagebox.showerror("错误", "输入随机数字失败")

        except Exception as e:
            messagebox.showerror("错误", f"输入随机数时出错: {str(e)}")
    
    def show_text_input_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("指定文本输入")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        self.set_dialog_icon(dialog)

        center_window(dialog, self.root)

        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="文本文件", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        file_path_var = tk.StringVar()
        file_path_entry = ttk.Entry(file_frame, textvariable=file_path_var, width=40)
        file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.setup_right_click_menu(file_path_entry)

        def browse_file():
            filepath = filedialog.askopenfilename(
                title="选择文本文件",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            )
            if filepath:
                file_path_var.set(filepath)
                try:
                    with open(filepath, "r", encoding="utf-8") as f:
                        lines = f.read().splitlines()
                        preview_text = "\n".join(lines[:10])
                        if len(lines) > 10:
                            preview_text += "\n..."
                        preview.delete(1.0, tk.END)
                        preview.insert(tk.END, preview_text)
                except UnicodeDecodeError:
                    try:
                        with open(filepath, "r", encoding="gbk") as f:
                            lines = f.read().splitlines()
                            preview_text = "\n".join(lines[:10])
                            if len(lines) > 10:
                                preview_text += "\n..."
                            preview.delete(1.0, tk.END)
                            preview.insert(tk.END, preview_text)
                    except Exception as e:
                        print(f"读取文件失败: {str(e)}")
                        messagebox.showerror("错误", f"读取文件失败: {str(e)}")
                except Exception as e:
                    print(f"读取文件失败: {str(e)}")
                    messagebox.showerror("错误", f"读取文件失败: {str(e)}")

        ttk.Button(file_frame, text="浏览...", command=browse_file).pack(side=tk.RIGHT)

        preview_frame = ttk.LabelFrame(main_frame, text="文件内容预览", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        preview = tk.Text(preview_frame, height=6, width=50, wrap=tk.WORD)
        preview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        preview_scrollbar = ttk.Scrollbar(
            preview_frame, orient=tk.VERTICAL, command=preview.yview
        )
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        preview.configure(yscrollcommand=preview_scrollbar.set)

        input_method_frame = ttk.Frame(main_frame)
        input_method_frame.pack(fill=tk.X, pady=(0, 10))

        input_method = tk.StringVar(value="sequential")

        sequential_row = ttk.Frame(input_method_frame)
        sequential_row.pack(side=tk.LEFT, padx=(0, 15))
        
        sequential_radio = ttk.Radiobutton(
            sequential_row,
            variable=input_method,
            value="sequential",
        )
        sequential_radio.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(
            sequential_row,
            text="顺序输入"
        ).pack(side=tk.LEFT)

        random_row = ttk.Frame(input_method_frame)
        random_row.pack(side=tk.LEFT)
        
        random_radio = ttk.Radiobutton(
            random_row,
            variable=input_method,
            value="random"
        )
        random_radio.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(
            random_row,
            text="随机输入"
        ).pack(side=tk.LEFT)

        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=(0, 10))

        overwrite_var = tk.BooleanVar(value=True)

        overwrite_row = ttk.Frame(options_frame)
        overwrite_row.pack(side=tk.LEFT)
        
        overwrite_check = ttk.Checkbutton(
            overwrite_row,
            variable=overwrite_var
        )
        overwrite_check.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(
            overwrite_row,
            text="覆盖原有内容"
        ).pack(side=tk.LEFT)

        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X)

        ttk.Button(buttons_frame, text="取消", command=dialog.destroy, width=10).pack(
            side=tk.RIGHT, padx=5
        )

        ttk.Button(
            buttons_frame,
            text="开始输入",
            command=lambda: self.execute_text_input(
                dialog,
                file_path_var.get(),
                input_method.get(),
                overwrite_var.get(),
                False,
            ),
            style="Accent.TButton",
            width=10,
        ).pack(side=tk.RIGHT, padx=5)

    def execute_text_input(self, dialog, file_path, input_method, overwrite, delayed):
        if not file_path:
            messagebox.showwarning("警告", "请选择文本文件！")
            return

        if not os.path.exists(file_path):
            messagebox.showerror("错误", "文件不存在！")
            return

        dialog.destroy()

        self.input_text_from_file(file_path, input_method, overwrite, delayed)

    def input_text_from_file(self, file_path, input_method, overwrite, delayed):
        try:
            selected = self.get_selected_windows()

            if not selected:
                messagebox.showwarning("警告", "请先选择要操作的窗口！")
                return

            window_handles = []
            for item in selected:
                hwnd = int(self.get_window_item_value(item, "hwnd"))
                window_handles.append(hwnd)
            
            success = input_text_from_file(
                window_handles=window_handles,
                file_path=file_path,
                input_method=input_method,
                overwrite=overwrite,
                delayed=delayed
            )
            
            if not success:
                messagebox.showerror("错误", "输入文本失败")

        except Exception as e:
            messagebox.showerror("错误", f"操作失败: {str(e)}")
    
    def create_environments(self):
        try:
            cache_dir = self.cache_dir
            shortcut_dir = self.shortcut_path
            numbers = self.env_numbers.get().strip()

            if not all([cache_dir, shortcut_dir, numbers]):
                messagebox.showwarning(
                    "警告", "请先在设置中填写缓存目录和快捷方式目录!"
                )
                return

            os.makedirs(cache_dir, exist_ok=True)
            os.makedirs(shortcut_dir, exist_ok=True)

            chrome_path = self.manager.find_chrome_path()
            if not chrome_path:
                messagebox.showerror("错误", "未找到Chrome安装路径！")
                return

            shell = win32com.client.Dispatch("WScript.Shell")

            window_numbers = parse_window_numbers(numbers)

            created_count = 0
            for i in window_numbers:
                data_dir_name = str(i)  

                data_dir = os.path.join(cache_dir, data_dir_name)
                data_dir = data_dir.replace("\\", "/")

                os.makedirs(data_dir, exist_ok=True)
                
                shortcut_path = os.path.join(shortcut_dir, f"{i}.lnk")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.TargetPath = chrome_path
                shortcut.Arguments = (
                    f'--user-data-dir="{data_dir}"'
                )
                shortcut.WorkingDirectory = os.path.dirname(chrome_path)
                shortcut.WindowStyle = 1
                shortcut.IconLocation = f"{chrome_path},0"
                shortcut.save()
                created_count += 1

            try:
                ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
                print("已刷新Windows资源管理器图标缓存")
            except Exception as e:
                print(f"刷新图标缓存失败: {str(e)}")

            messagebox.showinfo(
                "成功", f"已成功创建 {created_count} 个Chrome环境！"
            )
            
        except Exception as e:
            print(f"创建环境失败: {str(e)}")
            messagebox.showerror("错误", f"创建环境失败: {str(e)}")
            
    def restore_default_icons(self):
        try:
            if not messagebox.askyesno("确认", "确定要还原所有快捷方式的默认图标吗？"):
                return
            
            shortcut_path = self.shortcut_path
            if not shortcut_path or not os.path.exists(shortcut_path):
                messagebox.showerror("错误", "快捷方式目录不存在！")
                return
            
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("还原默认图标")
            progress_dialog.geometry("300x150")
            progress_dialog.transient(self.root)
            progress_dialog.grab_set()
            center_window(progress_dialog, self.root)
            self.set_dialog_icon(progress_dialog)
            
            progress_label = ttk.Label(progress_dialog, text="正在还原快捷方式图标...\n请稍候")
            progress_label.pack(pady=(20, 10))
            
            progress = ttk.Progressbar(progress_dialog, orient=tk.HORIZONTAL, length=250, mode="indeterminate")
            progress.pack(pady=5)
            progress.start()
            
            def restore_thread():
                error_msg = None
                try:
                    chrome_path = self.manager.find_chrome_path()
                    if not chrome_path:
                        self.root.after(0, lambda: progress_dialog.destroy())
                        self.root.after(0, lambda: messagebox.showerror("错误", "未找到Chrome安装路径！"))
                        return
                    
                    shortcuts = []
                    for file in os.listdir(shortcut_path):
                        if file.endswith(".lnk"):
                            shortcuts.append(os.path.join(shortcut_path, file))
                    
                    shell = win32com.client.Dispatch("WScript.Shell")
                    
                    for path in shortcuts:
                        try:
                            shortcut = shell.CreateShortCut(path)
                            shortcut.IconLocation = f"{chrome_path},0"
                            shortcut.save()
                        except Exception as e:
                            print(f"还原图标失败: {path} - {str(e)}")
                    
                    os.system("ie4uinit.exe -show")
                    
                    self.root.after(0, lambda: progress_dialog.destroy())
                    self.root.after(0, lambda: messagebox.showinfo("成功", f"已成功还原 {len(shortcuts)} 个快捷方式的默认图标"))
                
                except Exception as e:
                    error_msg = str(e)
                    self.root.after(0, lambda: progress_dialog.destroy())
                    self.root.after(0, lambda msg=error_msg: messagebox.showerror("错误", f"还原默认图标失败: {msg}"))
            
            threading.Thread(target=restore_thread, daemon=True).start()
            
        except Exception as e:
            log_error("还原默认图标失败", e)
            messagebox.showerror("错误", f"还原默认图标失败: {str(e)}")

    def clean_icon_cache(self):
        try:
            if not messagebox.askyesno("确认清理图标缓存", 
                "清理图标缓存将执行以下操作：\n\n"
                "1. 关闭所有资源管理器窗口\n"
                "2. 删除系统图标缓存文件\n"
                "3. 重启资源管理器\n\n"
                "此操作可能导致窗口短暂关闭，但不会影响您的数据。\n\n"
                "是否继续？"):
                return
            
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("清理图标缓存")
            progress_dialog.geometry("300x150")
            progress_dialog.transient(self.root)
            progress_dialog.grab_set()
            center_window(progress_dialog, self.root)
            self.set_dialog_icon(progress_dialog)
            
            progress_label = ttk.Label(progress_dialog, text="正在清理图标缓存...\n请稍候，不要关闭此窗口")
            progress_label.pack(pady=(20, 10))
            
            progress = ttk.Progressbar(progress_dialog, orient=tk.HORIZONTAL, length=250, mode="determinate", maximum=100)
            progress.pack(pady=5)
            
            def update_progress(value, message=None):
                self.root.after(10, lambda: progress.configure(value=value))
                if message:
                    self.root.after(10, lambda: progress_label.configure(text=message))
                self.root.update_idletasks()
            
            def is_process_running(process_name):
                try:
                    result = subprocess.run(
                        ["tasklist", "/FI", f"IMAGENAME eq {process_name}"], 
                        capture_output=True, 
                        text=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    return process_name.lower() in result.stdout.lower()
                except Exception:
                    return False
            
            def clean_thread():
                try:
                    update_progress(10, "正在关闭资源管理器...")
                    
                    subprocess.run(
                        ["taskkill", "/f", "/im", "explorer.exe"], 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    for _ in range(10):
                        if not is_process_running("explorer.exe"):
                            break
                        time.sleep(0.2)
                    
                    update_progress(30, "正在清理图标缓存文件...")
                    
                    subprocess.run(
                        'attrib -h -s -r "%userprofile%\\AppData\\Local\\IconCache.db"', 
                        shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    subprocess.run(
                        'del /f "%userprofile%\\AppData\\Local\\IconCache.db"', 
                        shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    subprocess.run(
                        'del /f "%userprofile%\\AppData\\Local\\Microsoft\\Windows\\Explorer\\IconCache_*.db"', 
                        shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    update_progress(50, "正在清理缩略图缓存...")
                    subprocess.run(
                        'attrib /s /d -h -s -r "%userprofile%\\AppData\\Local\\Microsoft\\Windows\\Explorer\\*"', 
                        shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    subprocess.run(
                        'del /f "%userprofile%\\AppData\\Local\\Microsoft\\Windows\\Explorer\\thumbcache_*.db"', 
                        shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    update_progress(70, "正在清理系统托盘图标记忆...")
                    subprocess.run(
                        'reg delete "HKEY_CLASSES_ROOT\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\TrayNotify" /v IconStreams /f', 
                        shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    subprocess.run(
                        'reg delete "HKEY_CLASSES_ROOT\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\TrayNotify" /v PastIconsStream /f', 
                        shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    update_progress(80, "正在刷新系统图标缓存...")
                    subprocess.run(
                        "ie4uinit.exe -show", 
                        shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    update_progress(85, "正在重启Shell服务...")
                    try:
                        subprocess.run(
                            'net stop "Shell Hardware Detection" && net start "Shell Hardware Detection"',
                            shell=True,
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                    except Exception:
                        pass
                    
                    update_progress(90, "正在重启资源管理器...")
                    
                    subprocess.Popen(
                        ["start", "explorer.exe"],
                        shell=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    time.sleep(0.5)
                    
                    try:
                        subprocess.run(
                            'powershell -command "Start-Process shell:::{05d7b0f4-2121-4eff-bf6b-ed3f69b894d9}"',
                            shell=True,
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                    except Exception:
                        pass
                    
                    time.sleep(1)

                    try:
                        from ctypes import Structure, c_long, c_ulong, c_ushort, c_short, POINTER, sizeof, byref
                        
                        class MOUSEINPUT(Structure):
                            _fields_ = [
                                ("dx", c_long),
                                ("dy", c_long),
                                ("mouseData", c_ulong),
                                ("dwFlags", c_ulong),
                                ("time", c_ulong),
                                ("dwExtraInfo", POINTER(c_ulong))
                            ]
                            
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
                            
                        class _INPUT_UNION(ctypes.Union):
                            _fields_ = [
                                ("mi", MOUSEINPUT),
                                ("ki", KEYBDINPUT),
                                ("hi", HARDWAREINPUT)
                            ]
                            
                        class INPUT(Structure):
                            _anonymous_ = ("u",)
                            _fields_ = [
                                ("type", c_ulong),
                                ("u", _INPUT_UNION)
                            ]
                        
                        # 常量定义
                        INPUT_KEYBOARD = 1
                        KEYEVENTF_KEYUP = 0x0002
                        VK_LWIN = 0x5B
                        
                        inp_down = INPUT()
                        inp_down.type = INPUT_KEYBOARD
                        inp_down.ki.wVk = VK_LWIN
                        inp_down.ki.wScan = 0
                        inp_down.ki.dwFlags = 0
                        inp_down.ki.time = 0
                        inp_down.ki.dwExtraInfo = ctypes.pointer(ctypes.c_ulong(0))
                        
                        inp_up = INPUT()
                        inp_up.type = INPUT_KEYBOARD
                        inp_up.ki.wVk = VK_LWIN
                        inp_up.ki.wScan = 0
                        inp_up.ki.dwFlags = KEYEVENTF_KEYUP
                        inp_up.ki.time = 0
                        inp_up.ki.dwExtraInfo = ctypes.pointer(ctypes.c_ulong(0))
                        
                        inputs = (INPUT * 2)()
                        inputs[0] = inp_down
                        inputs[1] = inp_up
                        
                        ctypes.windll.user32.SendInput(2, ctypes.byref(inputs), sizeof(INPUT))
                        time.sleep(0.1)
                        
                        inp_down.ki.wVk = 0x1B
                        inp_up.ki.wVk = 0x1B
                        inputs[0] = inp_down
                        inputs[1] = inp_up
                        ctypes.windll.user32.SendInput(2, ctypes.byref(inputs), sizeof(INPUT))
                    except Exception:
                        pass
                    
                    started = False
                    for _ in range(6):
                        if is_process_running("explorer.exe"):
                            started = True
                            break
                        time.sleep(0.5)
                    
                    if not started:
                        subprocess.Popen("explorer.exe")
                        time.sleep(1)
                    
                    update_progress(100, "图标缓存清理完成!")
                    
                    def show_completion():
                        try:
                            if progress_dialog.winfo_exists():
                                progress_dialog.destroy()
                            
                            self.root.attributes('-topmost', True)
                            self.root.attributes('-topmost', False)
                            self.root.lift()
                            self.root.focus_force()
                            
                            messagebox.showinfo("成功", "图标缓存清理完成！任务栏应该很快就会出现")
                        except Exception as e:
                            print(f"显示完成消息时出错: {str(e)}")
                    
                    self.root.after(1500, show_completion)
                
                except Exception as e:
                    if not is_process_running("explorer.exe"):
                        try:
                            subprocess.Popen("explorer.exe")
                        except:
                            pass
                    
                    def show_error():
                        if progress_dialog.winfo_exists():
                            progress_dialog.destroy()
                        messagebox.showerror("错误", f"清理图标缓存时出错: {str(e)}")
                    
                    self.root.after(0, show_error)
            
            threading.Thread(target=clean_thread, daemon=True).start()
            
        except Exception as e:
            print(f"一键清理图标缓存失败: {str(e)}")
            messagebox.showerror("错误", f"一键清理图标缓存失败: {str(e)}")
    
    def show_settings_dialog(self):
        try:
            dialog = tk.Toplevel(self.root)
            dialog.title("设置")
            dialog.geometry("500x350")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            
            center_window(dialog, self.root)
            self.set_dialog_icon(dialog)
            
            frame = ttk.Frame(dialog, padding=10)
            frame.pack(fill=tk.BOTH, expand=True)
            
            path_frame = ttk.LabelFrame(frame, text="路径设置")
            path_frame.pack(fill=tk.X, pady=(0, 10))
            
            shortcut_frame = ttk.Frame(path_frame)
            shortcut_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(shortcut_frame, text="快捷方式路径:").pack(side=tk.LEFT)
            
            shortcut_path_var = tk.StringVar(value=self.shortcut_path)
            shortcut_entry = ttk.Entry(shortcut_frame, textvariable=shortcut_path_var, width=30)
            shortcut_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            def browse_shortcut():
                path = filedialog.askdirectory(initialdir=self.shortcut_path)
                if path:
                    shortcut_path_var.set(path)
            
            ttk.Button(shortcut_frame, text="浏览", command=browse_shortcut).pack(side=tk.RIGHT)
            
            cache_frame = ttk.Frame(path_frame)
            cache_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(cache_frame, text="缓存目录:").pack(side=tk.LEFT)
            
            cache_dir_var = tk.StringVar(value=self.cache_dir)
            cache_entry = ttk.Entry(cache_frame, textvariable=cache_dir_var, width=30)
            cache_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            def browse_cache():
                path = filedialog.askdirectory(initialdir=self.cache_dir)
                if path:
                    cache_dir_var.set(path)
            
            ttk.Button(cache_frame, text="浏览", command=browse_cache).pack(side=tk.RIGHT)
            
            screen_frame = ttk.LabelFrame(frame, text="屏幕设置")
            screen_frame.pack(fill=tk.X, pady=(0, 10))
            
            self.screens, self.screen_names = update_screen_list()
            
            screen_settings_frame = ttk.Frame(screen_frame)
            screen_settings_frame.pack(fill=tk.X, pady=5, padx=10)
            
            ttk.Label(screen_settings_frame, text="屏幕选择:").pack(side=tk.LEFT)
            
            screen_var = tk.StringVar(value=self.screen_selection)
            screen_combo = ttk.Combobox(screen_settings_frame, textvariable=screen_var, width=20, state="readonly")
            screen_combo["values"] = self.screen_names
            screen_combo.pack(side=tk.LEFT, padx=5)
            
            if self.screen_selection and self.screen_selection in self.screen_names:
                screen_combo.set(self.screen_selection)
            elif self.screen_names:
                screen_combo.current(0)
            icon_frame = ttk.LabelFrame(frame, text="图标设置")
            icon_frame.pack(fill=tk.X, pady=(0, 10))
            
            auto_modify_icon_var = tk.BooleanVar(value=self.manager.auto_modify_shortcut_icon)
            
            icon_options_frame = ttk.Frame(icon_frame)
            icon_options_frame.pack(fill=tk.X, padx=10, pady=5)
            
            auto_icon_chk = ttk.Checkbutton(
                icon_options_frame,
                variable=auto_modify_icon_var
            )
            auto_icon_chk.pack(side=tk.LEFT, padx=(0, 5))
            
            ttk.Label(
                icon_options_frame,
                text="自动修改快捷方式图标"
            ).pack(side=tk.LEFT, padx=(0, 10))
            
            ttk.Button(
                icon_options_frame,
                text="一键还原快捷方式图标",
                command=self.restore_default_icons
            ).pack(side=tk.LEFT)
            
            ttk.Button(
                icon_options_frame,
                text="一键清理图标缓存",
                command=self.clean_icon_cache
            ).pack(side=tk.LEFT, padx=(10, 0))
            
            button_frame = ttk.Frame(frame)
            button_frame.pack(fill=tk.X, pady=(10, 0))
            
            ttk.Button(
                button_frame,
                text="保存",
                command=lambda: self.save_settings_dialog(
                    dialog,
                    shortcut_path_var.get(),
                    cache_dir_var.get(),
                    screen_var.get(),
                    auto_modify_icon_var.get()
                ),
                style="Accent.TButton"
            ).pack(side=tk.RIGHT, padx=(5, 0))
            
            ttk.Button(
                button_frame,
                text="取消",
                command=dialog.destroy
            ).pack(side=tk.RIGHT)
        
        except Exception as e:
            log_error("显示设置对话框失败", e)
    
    def save_settings_dialog(self, dialog, shortcut_path, cache_dir, screen, auto_modify_icon):
        try:
            self.shortcut_path = shortcut_path
            self.cache_dir = cache_dir
            
            screen_configs = self.settings.get("screen_arrange_config", [])
            if screen_configs:
                enabled_screens = [config for config in screen_configs if config["enabled"]]
                if enabled_screens:
                    enabled_screens.sort(key=lambda x: x["priority"])
                    first_screen_id = enabled_screens[0]["screen_id"]
                    if first_screen_id < len(self.screen_names):
                        screen = self.screen_names[first_screen_id]
            
            self.screen_selection = screen
            self.manager.auto_modify_shortcut_icon = auto_modify_icon
            
            self.manager.shortcut_path = shortcut_path
            self.manager.cache_dir = cache_dir
            self.manager.screen_selection = screen
            
            self.settings["shortcut_path"] = shortcut_path
            self.settings["cache_dir"] = cache_dir
            self.settings["screen_selection"] = screen
            self.settings["auto_modify_shortcut_icon"] = auto_modify_icon
            self.settings["show_chrome_tip"] = self.show_chrome_tip
            
            save_settings(self.settings)
            
            # 关闭对话框
            dialog.destroy()
            
            messagebox.showinfo("成功", "设置已保存！")
        
        except Exception as e:
            log_error("保存设置失败", e)
            messagebox.showerror("错误", f"保存设置失败: {str(e)}")
    
    def get_window_item_value(self, item, column):
        try:
            values = self.window_list.item(item)["values"]
            if column == "select":
                return values[0]
            elif column == "number":
                return values[1]
            elif column == "title":
                return values[2]
            elif column == "master":
                return values[3]
            elif column == "hwnd":
                return values[4]
            return None
        except Exception as e:
            log_error(f"获取窗口项目值失败: {column}", e)
            return None
    
    def get_selected_windows(self):
        selected = []
        for item in self.window_list.get_children():
            if self.window_list.set(item, "select") == "√":
                selected.append(item)
        return selected
    
    def ask_yes_no(self, message):
        return messagebox.askyesno("确认", message)

    def set_dialog_icon(self, dialog):
        if hasattr(self, 'app_icon_path') and self.app_icon_path and os.path.exists(self.app_icon_path):
            try:
                dialog.iconbitmap(self.app_icon_path)
            except Exception as e:
                log_error("设置对话框图标失败", e)
    
    def show_chrome_settings_tip(self):
        tip_dialog = tk.Toplevel(self.root)
        tip_dialog.title("Chrome后台运行提示")
        tip_dialog.geometry("420x255")
        tip_dialog.transient(self.root)
        tip_dialog.grab_set()

        tip_dialog.focus_set()
        
        self.set_dialog_icon(tip_dialog)

        tip_text = '如果窗口关闭后，Chrome仍在后台运行（右下角系统托盘区域里有多个chrome图标），请批量在浏览器设置页面取消后台运行：\n\n1. 批量打开Chrome浏览器\n2. 在地址栏输入：chrome://settings/system，或者进入设置-系统\n3. 找到"关闭 Google Chrome 后继续运行后台应用"选项\n4. 关闭该选项'

        tip_label = ttk.Label(
            tip_dialog, text=tip_text, justify=tk.LEFT, wraplength=380
        )
        tip_label.pack(pady=20, padx=20)

        dont_show_var = tk.BooleanVar(value=False)
        dont_show_check = ttk.Checkbutton(
            tip_dialog, text="下次不再显示", variable=dont_show_var
        )
        dont_show_check.pack(pady=10)

        def on_ok():
            if dont_show_var.get():
                self.show_chrome_tip = False
                self.save_tip_settings()
            tip_dialog.destroy()

        ok_button = ttk.Button(
            tip_dialog, text="确定", command=on_ok, style="Accent.TButton"
        )
        ok_button.pack(pady=10)

        center_window(tip_dialog, self.root)

    def save_tip_settings(self):
        try:
            self.show_chrome_tip = False

            self.settings["show_chrome_tip"] = False

            save_settings(self.settings)

            print(f"成功保存Chrome提示设置: show_chrome_tip = {self.show_chrome_tip}")

        except Exception as e:
            print(f"保存提示设置失败: {str(e)}")
            messagebox.showerror("设置保存失败", f"无法保存提示设置: {str(e)}") 

    def show_multi_screen_config_dialog(self):
        pass

    def reset_close_behavior(self):
        self.close_behavior = None
        self.settings["close_behavior"] = None
        save_settings(self.settings)
        messagebox.showinfo("设置已重置", "窗口关闭行为设置已重置，下次关闭窗口时将再次询问")

