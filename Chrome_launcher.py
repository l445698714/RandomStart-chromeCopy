#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Chrome分身启动器 - PyQt5版本
功能：随机启动Chrome分身、指定范围启动、指定范围关闭、关闭所有、在已打开分身中打开网址、图标管理
作者：Claude AI
"""

import sys
import os
import random
import subprocess
import time
import re
import psutil
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                           QLabel, QLineEdit, QPushButton, QGroupBox, QProgressBar,
                           QMessageBox, QFrame, QStatusBar, QFileDialog, QCheckBox,
                           QGridLayout, QTabWidget, QToolBar, QSpinBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings, QSize
from PyQt5.QtGui import QFont, QIcon
import win32com.client

# 导入图标管理功能
try:
    from chrome_icon_manager import ChromeIconManager, quick_apply_icons_to_chrome_windows
    from utils import generate_color_icon, set_chrome_icon
    ICON_MANAGEMENT_AVAILABLE = True
except ImportError as e:
    print(f"图标管理功能不可用: {e}")
    ICON_MANAGEMENT_AVAILABLE = False


class BackgroundWorker(QThread):
    """后台工作线程，用于处理耗时操作，避免UI卡顿"""
    update_status = pyqtSignal(str, str)  # 状态更新信号：消息，颜色
    update_progress = pyqtSignal(int)  # 进度更新信号
    finished = pyqtSignal(str, str, list)  # 完成信号：消息，颜色，处理的编号列表

    def __init__(self, profiles_data, folder_path, delay_time, mode="launch", url=None):
        """初始化后台工作线程
        
        Args:
            profiles_data: 数据列表。
                           对于 "launch" 和 "close" 模式, 这是要处理的编号列表 [num1, num2, ...]。
                           对于 "open_url" 模式, 这是 (分身编号, 用户数据目录路径, chrome.exe路径) 元组的列表 [(num, path, exe_path), ...]。
            folder_path: 快捷方式文件夹路径 (主要用于 "launch" 模式)。对于 "open_url" 和 "close" 模式可以为 None。
            delay_time: 操作间延迟时间（秒）
            mode: 操作模式："launch"、"close"或"open_url"
            url: 要打开的URL（仅在mode为"open_url"时使用）
        """
        super().__init__()
        self.profiles_data = profiles_data
        self.folder_path = folder_path # folder_path 对于 open_url 模式将是 None
        self.delay_time = delay_time
        self.mode = mode
        self.url = url
        # self.chrome_exe_path 不再需要作为全局变量，因为 open_url 模式会自带 exe 路径

    def _initialize_chrome_exe_path(self):
        """尝试自动检测chrome.exe的路径 (主要用于 launch 模式，如果未在主界面指定)
           对于 open_url 模式，将使用 profiles_data 中提供的特定 exe_path。
        """
        # 这个方法现在主要服务于 launch 模式，如果 ChromeLauncher 没有传递一个有效的 chrome_exe_path
        # 但目前我们的 launch 模式直接使用快捷方式，所以这个方法可能不太会被直接依赖。
        # 保留它以防未来有不通过快捷方式的启动模式需要全局 chrome.exe
        
        # settings = QSettings("ChromeLauncher", "Settings")
        # configured_chrome_path = settings.value("chrome_exe_path", None)
        # if configured_chrome_path and os.path.exists(configured_chrome_path):
        #     return configured_chrome_path # 返回路径而不是设置实例变量

        possible_paths = [
            r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
            os.path.expanduser(r"~\\AppData\\Local\\Google\\Chrome\\Application\\chrome.exe")
        ]
        for path in possible_paths:
            if os.path.exists(path):
                # settings.setValue("chrome_exe_path", path) # 可选：保存首次检测到的路径
                return path
        
        QMessageBox.critical(None, "错误", "未能自动检测到chrome.exe，请考虑在程序设置中手动指定Chrome路径！")
        return None

    def run(self):
        """根据模式执行相应操作"""
        if self.mode == "launch":
            self.launch_browsers()
        elif self.mode == "close":
            self.close_browsers() 
        elif self.mode == "open_url":
            # 对于 open_url 模式，每个条目已包含 chrome.exe 路径，不需要全局初始化
            self.open_url_in_browsers()

    def launch_browsers(self):
        """启动Chrome分身 (self.profiles_data 是编号列表)"""
        success_count = 0
        successful_numbers = []
        
        if not self.profiles_data: # 检查列表是否为空
            self.finished.emit("没有选择任何分身进行启动", "orange", [])
            return
            
        total_steps = len(self.profiles_data)
        
        for i, n in enumerate(self.profiles_data): # self.profiles_data 是编号列表
            progress = int((i + 1) / total_steps * 100)
            self.update_progress.emit(progress)
            
            # 发送当前启动状态信息（格式化信息，用于主要日志）
            remaining_numbers = self.profiles_data[i+1:] if i+1 < len(self.profiles_data) else []
            self.update_status.emit(f"CURRENT_LAUNCHING:{n}|LAUNCHED:{','.join(map(str, successful_numbers))}|REMAINING:{','.join(map(str, remaining_numbers))}", "blue")
            
            if not self.folder_path: # 启动模式必须有快捷方式文件夹路径
                 # 错误信息可以覆盖主要日志
                 self.update_status.emit(f"错误: 未提供快捷方式文件夹路径，无法启动分身 {n}", "red")
                 continue
            shortcut_path = os.path.join(self.folder_path, f"{n}.lnk")
            try:
                if os.path.exists(shortcut_path):
                    subprocess.Popen(["cmd", "/c", "start", "", shortcut_path], shell=True)
                    success_count += 1
                    successful_numbers.append(n)
                    time.sleep(float(self.delay_time))
                else:
                    # 警告信息可以覆盖主要日志
                    self.update_status.emit(f"警告: 找不到快捷方式 {n}.lnk", "orange")
            except Exception as e:
                # 错误信息可以覆盖主要日志
                self.update_status.emit(f"警告: 启动 {n} 失败: {str(e)}", "red")
        
        if success_count > 0:
            status_text = f"成功启动{success_count}个Chrome浏览器!\n已启动编号: {', '.join(map(str, successful_numbers))}"
            self.finished.emit(status_text, "green", successful_numbers)
        else:
            self.finished.emit("错误: 没有成功启动任何浏览器 (或未提供快捷方式路径)", "red", [])

    def close_browsers(self):
        """关闭指定范围的Chrome分身 (self.profiles_data 是编号列表)"""
        closed_count = 0
        closed_numbers_set = set() # 使用集合避免重复
        
        chrome_processes = []
        for proc in psutil.process_iter(attrs=['pid', 'name', 'cmdline'], ad_value=None):
            try:
                # 确保 proc.info['name'] 和 proc.info['cmdline'] 都不是 None
                if proc.info['name'] and 'chrome.exe' in proc.info['name'].lower() and proc.info['cmdline']:
                    chrome_processes.append(proc)
            except (psutil.NoSuchProcess, psutil.AccessDenied, TypeError): # 添加TypeError以防info字段为None
                continue
        
        if not self.profiles_data: # 检查列表是否为空
            self.finished.emit("没有选择任何分身进行关闭", "orange", [])
            return
            
        total_steps = len(self.profiles_data)

        for i, n in enumerate(self.profiles_data): # self.profiles_data 是编号列表
            progress = int((i + 1) / total_steps * 100)
            self.update_progress.emit(progress)
            
            # 当前关闭逻辑依赖于命令行参数中包含 "chrome<N>" 的模式。
            # 这部分可以进一步优化为使用精确的 user_data_dir (如果能从某处获取映射关系)。
            # assumed_data_dir_fragment 用来在命令行参数中寻找与分身编号相关的部分。
            assumed_data_dir_fragment_pattern = re.compile(fr"chrome{n}(?:[\\/\"']|$)", re.IGNORECASE)
            
            terminated_this_iter = False
            for proc in list(chrome_processes): # Iterate over a copy of the list
                try:
                    # 确保 cmdline 不是 None
                    if proc.info['cmdline'] is None:
                        continue
                    cmd_line_str = " ".join(proc.info['cmdline'])
                    # 使用更精确的正则匹配 --user-data-dir="...\chromeN..." 或 --user-data-dir=...\chromeN...
                    user_data_dir_arg_match = re.search(r"--user-data-dir=(?:\"([^\"]*)\"|([^ ]+(?: [^ ]+)*?(?=(?: --|$))))", cmd_line_str)
                    if user_data_dir_arg_match:
                        user_data_path_from_cmd = user_data_dir_arg_match.group(1) or user_data_dir_arg_match.group(2)
                        if user_data_path_from_cmd and assumed_data_dir_fragment_pattern.search(user_data_path_from_cmd):
                            proc.terminate()
                            try:
                                proc.wait(timeout=3) # 增加超时时间到3秒
                            except psutil.TimeoutExpired:
                                proc.kill() # 强制杀死如果超时
                            closed_numbers_set.add(n)
                            terminated_this_iter = True
                            chrome_processes.remove(proc) # 从列表中移除，避免重复处理
                            status_text = f"已尝试关闭分身: {n}"
                            self.update_status.emit(status_text, "blue")
                            break # 假设一个编号对应一个主要用户数据目录实例
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    # 如果进程已经消失或无权访问，则从列表中移除
                    if proc in chrome_processes:
                         chrome_processes.remove(proc)
                    continue
                except Exception: # 其他潜在错误
                    if proc in chrome_processes:
                         chrome_processes.remove(proc)
                    continue
            
            # 如果通过 psutil 终止成功，可以尝试 taskkill 作为补充，但需小心误杀
            if terminated_this_iter:
                try:
                    # taskkill 基于窗口标题的关闭非常依赖窗口标题的命名规则
                    # 例如，如果窗口标题包含 "Profile N" 或 "chromeN"
                    # subprocess.run(f'taskkill /F /FI "IMAGENAME eq chrome.exe" /FI "WINDOWTITLE eq *chrome{n}*"', 
                    #               shell=True, check=False, capture_output=True, timeout=1)
                    pass # 暂时禁用taskkill，因为它可能不够精确，依赖psutil的终止
                except Exception:
                    pass 

            time.sleep(float(self.delay_time) / 4 if self.delay_time else 0.02) # 减少延迟

        closed_list = sorted(list(closed_numbers_set))
        if closed_list:
            status_text = f"成功关闭 {len(closed_list)} 个指定Chrome分身窗口!\n已关闭编号: {', '.join(map(str, closed_list))}"
            self.finished.emit(status_text, "green", closed_list)
        else:
            self.finished.emit("在指定范围内没有找到或未能关闭活动的Chrome分身窗口", "orange", [])


    def open_url_in_browsers(self):
        """在指定的Chrome分身中打开URL (self.profiles_data 是 (用户数据目录路径, chrome_exe路径) 元组的列表)"""
        success_count = 0
        successful_instances_info = [] # 用于记录成功操作的实例信息 (udd)

        if not self.profiles_data: # 检查列表是否为空
            self.finished.emit("没有已运行的Chrome实例来打开网址", "orange", [])
            return
            
        total_steps = len(self.profiles_data)

        for i, profile_entry in enumerate(self.profiles_data):
            progress = int((i + 1) / total_steps * 100)
            self.update_progress.emit(progress)
            
            # 解包 profile_entry - 新格式: (user_data_dir, specific_chrome_exe_path)
            if len(profile_entry) == 2:
                user_data_dir, specific_chrome_exe_path = profile_entry
            else:
                self.update_status.emit(f"警告: profiles_data 条目格式不正确，跳过。数据: {profile_entry}", "red")
                continue

            try:
                if not specific_chrome_exe_path or not os.path.exists(specific_chrome_exe_path):
                    self.update_status.emit(f"警告: Chrome.exe 路径无效或未找到: {specific_chrome_exe_path} (UDD: {os.path.basename(user_data_dir)})", "red")
                    continue

                if not isinstance(user_data_dir, str) or not os.path.isdir(user_data_dir):
                     # 对于"在已运行分身中打开"，user_data_dir 应该已存在且为目录。
                     self.update_status.emit(f"警告: 用户数据目录路径无效或不是目录: {user_data_dir}", "red")
                     continue

                cmd = [
                    specific_chrome_exe_path, 
                    f"--user-data-dir={user_data_dir}", 
                    self.url
                ]
                subprocess.Popen(cmd)
                success_count += 1
                successful_instances_info.append(os.path.basename(user_data_dir)) # 记录 UDD 的 basename 作为标识
                status_text = f"正在实例 (UDD: ...{os.path.basename(user_data_dir)}) 中打开网址: {self.url}"
                self.update_status.emit(status_text, "blue")
                time.sleep(float(self.delay_time))
            except Exception as e:
                self.update_status.emit(f"警告: 在实例 (UDD: {os.path.basename(user_data_dir)}) 中打开网址失败: {str(e)}", "red")
        
        if success_count > 0:
            status_text = f"成功在 {success_count} 个Chrome实例中打开网址!\n实例 (UDD Basenames): {', '.join(successful_instances_info)}"
            self.finished.emit(status_text, "green", successful_instances_info)
        else:
            self.finished.emit("错误: 未能在任何指定Chrome实例中打开网址", "red", [])


class IconManagementWorker(QThread):
    """图标管理专用工作线程"""
    update_status = pyqtSignal(str, str)  # 状态更新信号：消息，颜色
    update_progress = pyqtSignal(int)  # 进度更新信号
    finished = pyqtSignal(str, str)  # 完成信号：消息，颜色
    
    def __init__(self, operation_type, numbers=None, shortcut_dir=None):
        """
        初始化图标管理工作线程
        
        Args:
            operation_type: 操作类型 ('generate_icons', 'apply_icons', 'update_shortcuts', 'restore_defaults', 'clean_cache')
            numbers: 要处理的分身编号列表
            shortcut_dir: 快捷方式目录路径
        """
        super().__init__()
        self.operation_type = operation_type
        self.numbers = numbers or []
        self.shortcut_dir = shortcut_dir
        
        # 获取图标管理器实例
        if ICON_MANAGEMENT_AVAILABLE:
            from chrome_icon_manager import create_chrome_icon_manager
            self.icon_manager = create_chrome_icon_manager()
        else:
            self.icon_manager = None
    
    def run(self):
        """执行图标管理操作"""
        try:
            if not self.icon_manager:
                self.finished.emit("图标管理器不可用", "red")
                return
                
            if self.operation_type == "generate_icons":
                self.generate_icons()
            elif self.operation_type == "apply_icons":
                self.apply_icons()
            elif self.operation_type == "update_shortcuts":
                self.update_shortcuts()
            elif self.operation_type == "restore_defaults":
                self.restore_defaults()
            elif self.operation_type == "clean_cache":
                self.clean_cache()
            else:
                self.finished.emit(f"未知的操作类型: {self.operation_type}", "red")
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.finished.emit(f"图标管理操作失败: {str(e)}", "red")
    
    def generate_icons(self):
        """生成图标"""
        if not self.numbers:
            self.finished.emit("没有指定要生成的图标编号", "orange")
            return
        
        self.update_status.emit(f"开始生成 {len(self.numbers)} 个图标...", "blue")
        
        success_count = 0
        total = len(self.numbers)
        
        for i, number in enumerate(self.numbers):
            try:
                self.update_progress.emit(int((i / total) * 100))
                self.icon_manager.generate_numbered_icon(number)
                success_count += 1
                self.update_status.emit(f"已生成图标 {number} ({i+1}/{total})", "blue")
            except Exception as e:
                self.update_status.emit(f"生成图标 {number} 失败: {str(e)}", "red")
        
        self.update_progress.emit(100)
        
        if success_count == total:
            self.finished.emit(f"成功生成了 {success_count} 个图标", "green")
        else:
            self.finished.emit(f"生成完成，成功 {success_count}/{total} 个图标", "orange")
    
    def apply_icons(self):
        """应用图标到窗口"""
        if not self.numbers:
            self.finished.emit("没有指定要应用图标的窗口", "orange")
            return
        
        self.update_status.emit(f"开始为 {len(self.numbers)} 个分身应用图标...", "blue")
        
        # 查找Chrome窗口
        try:
            # find_chrome_windows 返回 {hwnd: number} 映射
            chrome_windows_map = self.icon_manager.find_chrome_windows()
            if not chrome_windows_map:
                self.finished.emit("没有找到Chrome窗口", "orange")
                return
            
            # 构建目标窗口映射：{hwnd: number} 只包含我们需要的编号
            target_windows_map = {}
            for hwnd, window_number in chrome_windows_map.items():
                if window_number in self.numbers:
                    target_windows_map[hwnd] = window_number
            
            if not target_windows_map:
                self.finished.emit(f"没有找到编号为 {', '.join(map(str, self.numbers))} 的Chrome窗口", "orange")
                return
            
            # 批量应用图标
            def progress_callback(current, total, message):
                self.update_progress.emit(int((current / total) * 100))
                self.update_status.emit(message, "blue")
            
            # batch_apply_icons_to_windows 期望 {hwnd: number} 映射
            results = self.icon_manager.batch_apply_icons_to_windows(
                target_windows_map, progress_callback=progress_callback
            )
            
            self.update_progress.emit(100)
            
            # 统计成功数量
            success_count = sum(1 for success in results.values() if success)
            
            if success_count == len(target_windows_map):
                self.finished.emit(f"成功为 {success_count} 个窗口应用了图标", "green")
            else:
                self.finished.emit(f"图标应用完成，成功 {success_count}/{len(target_windows_map)} 个窗口", "orange")
                
        except Exception as e:
            self.finished.emit(f"应用图标失败: {str(e)}", "red")
    
    def update_shortcuts(self):
        """更新快捷方式图标"""
        if not self.shortcut_dir or not os.path.exists(self.shortcut_dir):
            self.finished.emit("快捷方式目录无效", "red")
            return
        
        self.update_status.emit("开始更新快捷方式图标...", "blue")
        
        try:
            # 如果指定了编号，只更新这些编号的快捷方式
            if self.numbers:
                shortcuts_to_update = []
                for number in self.numbers:
                    shortcut_path = os.path.join(self.shortcut_dir, f"{number}.lnk")
                    if os.path.exists(shortcut_path):
                        shortcuts_to_update.append((shortcut_path, number))
                
                if not shortcuts_to_update:
                    self.finished.emit("没有找到对应编号的快捷方式文件", "orange")
                    return
            else:
                # 扫描目录中的所有快捷方式
                shortcuts_to_update = []
                for file in os.listdir(self.shortcut_dir):
                    if file.endswith('.lnk') and file[:-4].isdigit():
                        number = int(file[:-4])
                        shortcut_path = os.path.join(self.shortcut_dir, file)
                        shortcuts_to_update.append((shortcut_path, number))
                
                if not shortcuts_to_update:
                    self.finished.emit("没有找到数字命名的快捷方式文件", "orange")
                    return
            
            success_count = 0
            total = len(shortcuts_to_update)
            
            for i, (shortcut_path, number) in enumerate(shortcuts_to_update):
                try:
                    self.update_progress.emit(int((i / total) * 100))
                    self.icon_manager.update_shortcut_icons([(shortcut_path, number)])
                    success_count += 1
                    self.update_status.emit(f"已更新快捷方式 {number} ({i+1}/{total})", "blue")
                except Exception as e:
                    self.update_status.emit(f"更新快捷方式 {number} 失败: {str(e)}", "red")
            
            self.update_progress.emit(100)
            
            if success_count == total:
                self.finished.emit(f"成功更新了 {success_count} 个快捷方式图标", "green")
            else:
                self.finished.emit(f"更新完成，成功 {success_count}/{total} 个快捷方式", "orange")
                
        except Exception as e:
            self.finished.emit(f"更新快捷方式图标失败: {str(e)}", "red")
    
    def restore_defaults(self):
        """恢复默认图标"""
        if not self.shortcut_dir or not os.path.exists(self.shortcut_dir):
            self.finished.emit("快捷方式目录无效", "red")
            return
        
        self.update_status.emit("开始恢复默认图标...", "blue")
        
        try:
            self.icon_manager.restore_default_chrome_icons(self.shortcut_dir)
            self.update_progress.emit(100)
            self.finished.emit("成功恢复默认图标", "green")
        except Exception as e:
            self.finished.emit(f"恢复默认图标失败: {str(e)}", "red")
    
    def clean_cache(self):
        """清理图标缓存"""
        self.update_status.emit("开始清理图标缓存...", "blue")
        
        try:
            self.update_progress.emit(50)
            self.icon_manager.clean_system_icon_cache()
            self.update_progress.emit(100)
            self.finished.emit("成功清理图标缓存", "green")
        except Exception as e:
            self.finished.emit(f"清理图标缓存失败: {str(e)}", "red")


class ProfileCreationWorker(QThread):
    """分身创建专用工作线程"""
    update_status = pyqtSignal(str, str)  # 状态更新信号：消息，颜色
    update_progress = pyqtSignal(int)  # 进度更新信号
    finished = pyqtSignal(str, str, int)  # 完成信号：消息，颜色，创建数量
    
    def __init__(self, shortcut_path, cache_path, start_num, end_num):
        """
        初始化分身创建工作线程
        
        Args:
            shortcut_path: 快捷方式保存路径
            cache_path: 缓存数据保存路径
            start_num: 起始编号
            end_num: 结束编号
        """
        super().__init__()
        self.shortcut_path = shortcut_path
        self.cache_path = cache_path
        self.start_num = start_num
        self.end_num = end_num
        self.total_count = end_num - start_num + 1
        
        # 查找Chrome可执行文件路径
        self.chrome_exe_path = self._find_chrome_executable()
        
    def _find_chrome_executable(self):
        """查找Chrome可执行文件路径"""
        possible_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
            os.path.join(os.environ.get('PROGRAMFILES', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
            os.path.join(os.environ.get('PROGRAMFILES(X86)', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        return None
    
    def run(self):
        """执行分身创建任务"""
        try:
            if not self.chrome_exe_path:
                self.finished.emit("未找到Chrome安装路径！请确保已安装Chrome浏览器。", "red", 0)
                return
            
            self.update_status.emit(f"开始创建编号 {self.start_num}-{self.end_num} 的Chrome分身...", "blue")
            
            created_count = 0
            
            for i in range(self.start_num, self.end_num + 1):
                try:
                    # 更新进度
                    progress = int(((i - self.start_num + 1) / self.total_count) * 100)
                    self.update_progress.emit(progress)
                    self.update_status.emit(f"正在创建分身 {i}...", "blue")
                    
                    # 创建缓存目录
                    profile_cache_dir = os.path.join(self.cache_path, str(i))
                    os.makedirs(profile_cache_dir, exist_ok=True)
                    
                    # 创建快捷方式
                    shortcut_file_path = os.path.join(self.shortcut_path, f"{i}.lnk")
                    
                    # 创建快捷方式使用COM组件
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut = shell.CreateShortCut(shortcut_file_path)
                    shortcut.TargetPath = self.chrome_exe_path
                    shortcut.Arguments = f'--user-data-dir="{profile_cache_dir}"'
                    shortcut.WorkingDirectory = os.path.dirname(self.chrome_exe_path)
                    shortcut.IconLocation = f"{self.chrome_exe_path},0"
                    shortcut.Description = f"Chrome分身 {i}"
                    shortcut.save()
                    
                    created_count += 1
                    self.update_status.emit(f"分身 {i} 创建成功 ({created_count}/{self.total_count})", "green")
                    
                    # 短暂延迟，避免系统负载过高
                    time.sleep(0.1)
                    
                except Exception as e:
                    self.update_status.emit(f"创建分身 {i} 失败: {str(e)}", "red")
                    continue
            
            # 完成后刷新系统图标缓存
            try:
                subprocess.run('ie4uinit.exe -show', shell=True, capture_output=True, timeout=10)
            except Exception:
                pass  # 忽略刷新图标缓存的错误
            
            self.update_progress.emit(100)
            
            if created_count == self.total_count:
                self.finished.emit(f"成功创建了编号 {self.start_num}-{self.end_num} 的 {created_count} 个Chrome分身！", "green", created_count)
            elif created_count > 0:
                self.finished.emit(f"部分成功：创建了 {created_count}/{self.total_count} 个Chrome分身", "orange", created_count)
            else:
                self.finished.emit("创建失败：没有成功创建任何分身", "red", 0)
                
        except Exception as e:
            self.finished.emit(f"创建分身时发生严重错误: {str(e)}", "red", 0)


class ChromeLauncher(QMainWindow):
    """Chrome分身启动器主窗口"""
    
    def __init__(self):
        """初始化主窗口"""
        super().__init__()
        
        # 图标管理器初始化
        self.icon_manager = None
        self.auto_apply_icons = True
        
        if ICON_MANAGEMENT_AVAILABLE:
            try:
                from chrome_icon_manager import create_chrome_icon_manager
                self.icon_manager = create_chrome_icon_manager()
                print("图标管理器初始化成功")
            except Exception as e:
                print(f"图标管理器初始化失败: {e}")
                self.icon_manager = None
        
        # 初始化设置
        self.settings = QSettings("ChromeLauncher", "Settings")
        
        # 创建分身相关设置
        self.shortcut_creation_path = ""
        self.cache_creation_path = ""
        
        # 记录已启动的分身编号
        self.launched_numbers = set()
        
        # 依次启动相关变量
        self.sequential_launch_active = False
        self.sequential_launch_range = ""  # 存储依次启动的范围字符串
        self.sequential_launch_profiles = []  # 存储依次启动的分身编号列表
        self.sequential_launch_current_index = 0 # 指向下一个待启动分身的索引
        self._currently_attempting_sequential_profile = None # 临时存储当前尝试启动的分身号

        # 定义用于从命令行提取分身编号的正则表达式模式
        self.profile_num_patterns = [
            re.compile(r"chrome[\\/]?(\d+)", re.IGNORECASE),
            re.compile(r"profile[\\/]?(\d+)", re.IGNORECASE),
            re.compile(r"--user-data-dir(?:\"|\'|=|\s)+.*?chrome(\d+)", re.IGNORECASE),
            re.compile(r"--profile-directory=(?:\"Profile\s+(\d+)\"?|Default)", re.IGNORECASE)
        ]
        self.user_data_dir_pattern = re.compile(r'--user-data-dir=(?:\"(?P<path>[^"]+)\"|(?P<path_unquoted>[^\s]+(?:\s+[^\s]+)*?(?=\s*--|\s*$)))')
        
        # 设置窗口属性
        self.setWindowTitle("Chrome分身启动器 V3.0")
        self.setFixedSize(580, 380)
        
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ico.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"警告: 图标文件 {icon_path} 不存在")
        
        # 同步/刷新 launched_numbers 以匹配当前实际运行的Chrome分身
        self._sync_launched_numbers_with_running_processes()
        
        # 应用浅色主题样式
        self.apply_light_theme()
        
        # 初始化用户界面
        self.init_ui()
            
        initial_status_message = "准备就绪。"
        if self.launched_numbers:
            sorted_launched_numbers = sorted(list(self.launched_numbers))
            launched_numbers_str = ", ".join(map(str, sorted_launched_numbers))
            initial_status_message += f" 当前已通过编号识别并记录的已启动分身: {launched_numbers_str}。"
        else:
            initial_status_message += " 当前没有通过编号识别的已启动分身被记录。"
        
        # 添加图标管理状态信息
        if ICON_MANAGEMENT_AVAILABLE and self.icon_manager:
            initial_status_message += " 图标管理功能已启用。"
        else:
            initial_status_message += " 图标管理功能不可用。"
            
        self.set_status(initial_status_message, "blue")
        self.statusBar.showMessage(f"就绪。已记录编号分身: {len(self.launched_numbers)} 个。")
    
    def init_ui(self):
        """初始化用户界面"""
        # 设置中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(8)  # 减少间距
        main_layout.setContentsMargins(10, 10, 10, 10)  # 减少边距
        
        # === 基础设置区域（合并路径和核心参数） ===
        config_group = self.create_group_box("基础设置")
        config_layout = QVBoxLayout(config_group)
        config_layout.setContentsMargins(10, 20, 10, 10)
        config_layout.setSpacing(8)
        
        # 文件夹路径
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("快捷方式目录:"))
        self.folder_path = QLineEdit()
        self.folder_path.setMinimumWidth(250)
        folder_layout.addWidget(self.folder_path)
        
        browse_button = QPushButton("浏览...")
        browse_button.clicked.connect(self.browse_folder)
        browse_button.setFixedWidth(60)
        folder_layout.addWidget(browse_button)
        config_layout.addLayout(folder_layout)
        
        # 参数设置（一行显示）
        params_layout = QHBoxLayout()
        
        params_layout.addWidget(QLabel("范围:"))
        self.start_num = QLineEdit("1")
        self.start_num.setFixedWidth(45)
        self.start_num.setFixedHeight(24)
        params_layout.addWidget(self.start_num)
        
        params_layout.addWidget(QLabel("-"))
        self.end_num = QLineEdit("100")
        self.end_num.setFixedWidth(45)
        self.end_num.setFixedHeight(24)
        params_layout.addWidget(self.end_num)
        
        params_layout.addWidget(QLabel("  数量:"))
        self.num_browsers = QLineEdit("5")
        self.num_browsers.setFixedWidth(40)
        self.num_browsers.setFixedHeight(24)
        params_layout.addWidget(self.num_browsers)
        
        params_layout.addWidget(QLabel("  延迟:"))
        self.delay_time = QLineEdit("0.5")
        self.delay_time.setFixedWidth(40)
        self.delay_time.setFixedHeight(24)
        params_layout.addWidget(self.delay_time)
        params_layout.addWidget(QLabel("秒"))
        
        params_layout.addWidget(QLabel("  指定范围:"))
        self.specific_range = QLineEdit("1-10")
        self.specific_range.setFixedWidth(100)  # 设置合适的固定宽度
        self.specific_range.setFixedHeight(24)
        params_layout.addWidget(self.specific_range)
        
        params_layout.addStretch()
        config_layout.addLayout(params_layout)
        
        main_layout.addWidget(config_group)
        
        # === 分栏页面 ===
        tab_widget = QTabWidget()
        tab_widget.setFixedHeight(140)  # 限制分栏高度，保持界面紧凑
        
        # 设置分栏标题字体加粗且大小统一
        tab_widget.setStyleSheet("""
            QTabWidget::tab-bar {
                alignment: left;
            }
            QTabBar::tab {
                font-family: 'Microsoft YaHei';
                font-size: 10pt;
                font-weight: bold;
                padding: 8px 12px;
                margin-right: 2px;
                background-color: #E8E8E8;
                border: 1px solid #CCCCCC;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                color: #666666;
            }
            QTabBar::tab:selected {
                background-color: #4A90E2;
                color: white;
                border-bottom: 2px solid #4A90E2;
                font-weight: bold;
            }
            QTabBar::tab:hover:!selected {
                background-color: #D6E7F7;
                color: #2C5FA3;
            }
            QTabWidget::pane {
                border: 1px solid #CCCCCC;
                background-color: #F8F8F8;
                border-top: 2px solid #4A90E2;
            }
        """)
        
        # 功能操作分栏
        operations_tab = QWidget()
        operations_layout = QVBoxLayout(operations_tab)
        operations_layout.setContentsMargins(10, 10, 10, 10)
        operations_layout.setSpacing(6)
        
        # 操作按钮区域
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(6)
        self.create_compact_button("随机启动", self.launch_random_browsers, buttons_layout)
        self.create_compact_button("指定启动", self.launch_specific_range, buttons_layout)
        self.create_compact_button("依次启动", self.launch_sequentially, buttons_layout)
        self.create_compact_button("指定关闭", self.close_specific_range, buttons_layout)
        self.create_compact_button("关闭所有", self.close_all_chrome, buttons_layout)
        operations_layout.addLayout(buttons_layout)
        
        # 网址操作区域
        url_input_layout = QHBoxLayout()
        url_input_layout.addWidget(QLabel("网址:"))
        self.url_entry = QLineEdit("https://www.example.com")
        self.url_entry.setFixedHeight(24)
        url_input_layout.addWidget(self.url_entry)
        
        url_button = QPushButton("在已启动分身中打开网址")
        url_button.clicked.connect(self.open_url_in_running)
        url_button.setFixedHeight(26)
        url_input_layout.addWidget(url_button)
        operations_layout.addLayout(url_input_layout)
        
        tab_widget.addTab(operations_tab, "功能操作")
        
        # 图标管理分栏
        if ICON_MANAGEMENT_AVAILABLE:
            icon_tab = QWidget()
            icon_layout = QVBoxLayout(icon_tab)
            icon_layout.setContentsMargins(10, 10, 10, 10)
            icon_layout.setSpacing(6)
            
            # 自动图标应用设置
            auto_icon_layout = QHBoxLayout()
            self.auto_apply_checkbox = QCheckBox("启动Chrome时自动应用编号图标")
            self.auto_apply_checkbox.setChecked(self.auto_apply_icons)
            self.auto_apply_checkbox.stateChanged.connect(self.on_auto_apply_icons_changed)
            self.auto_apply_checkbox.setFont(QFont("Microsoft YaHei", 7))
            auto_icon_layout.addWidget(self.auto_apply_checkbox)
            auto_icon_layout.addStretch()
            icon_layout.addLayout(auto_icon_layout)
            
            # 图标管理按钮区域 (2行3列布局)
            buttons_frame = QFrame()
            buttons_grid = QGridLayout(buttons_frame)
            buttons_grid.setSpacing(6)
            buttons_grid.setContentsMargins(0, 0, 0, 0)
            
            # 第一行按钮
            buttons_grid.addWidget(self.create_compact_icon_button("生成图标", self.generate_all_icons), 0, 0)
            buttons_grid.addWidget(self.create_compact_icon_button("应用图标", self.apply_icons_to_windows), 0, 1)
            buttons_grid.addWidget(self.create_compact_icon_button("更新快捷方式", self.update_shortcut_icons), 0, 2)
            
            # 第二行按钮
            buttons_grid.addWidget(self.create_compact_icon_button("恢复默认", self.restore_default_icons), 1, 0)
            buttons_grid.addWidget(self.create_compact_icon_button("清理缓存", self.clean_icon_cache), 1, 1)
            buttons_grid.addWidget(self.create_compact_icon_button("缓存信息", self.show_cache_info), 1, 2)
            
            icon_layout.addWidget(buttons_frame)
            tab_widget.addTab(icon_tab, "图标管理")
        
        # 创建分身分栏
        profile_tab = QWidget()
        profile_layout = QVBoxLayout(profile_tab)
        profile_layout.setContentsMargins(10, 10, 10, 10)
        profile_layout.setSpacing(6)
        
        # 快捷方式路径设置
        shortcut_path_layout = QHBoxLayout()
        shortcut_path_layout.addWidget(QLabel("快捷方式路径:"))
        self.shortcut_path_entry = QLineEdit()
        self.shortcut_path_entry.setFixedHeight(24)
        self.shortcut_path_entry.setPlaceholderText("选择快捷方式保存目录")
        shortcut_path_layout.addWidget(self.shortcut_path_entry)
        
        shortcut_browse_button = QPushButton("浏览")
        shortcut_browse_button.clicked.connect(self.browse_shortcut_path)
        shortcut_browse_button.setFixedWidth(50)
        shortcut_browse_button.setFixedHeight(24)
        shortcut_path_layout.addWidget(shortcut_browse_button)
        profile_layout.addLayout(shortcut_path_layout)
        
        # 缓存储存路径设置
        cache_path_layout = QHBoxLayout()
        cache_path_layout.addWidget(QLabel("缓存储存路径:"))
        self.cache_path_entry = QLineEdit()
        self.cache_path_entry.setFixedHeight(24)
        self.cache_path_entry.setPlaceholderText("选择缓存数据保存目录")
        cache_path_layout.addWidget(self.cache_path_entry)
        
        cache_browse_button = QPushButton("浏览")
        cache_browse_button.clicked.connect(self.browse_cache_path)
        cache_browse_button.setFixedWidth(50)
        cache_browse_button.setFixedHeight(24)
        cache_path_layout.addWidget(cache_browse_button)
        profile_layout.addLayout(cache_path_layout)
        
        # 数量设置和创建按钮
        create_layout = QHBoxLayout()
        create_layout.addWidget(QLabel("创建编号:"))
        self.create_start_entry = QLineEdit("1")
        self.create_start_entry.setFixedWidth(45)
        self.create_start_entry.setFixedHeight(24)
        self.create_start_entry.setPlaceholderText("起始")
        create_layout.addWidget(self.create_start_entry)
        
        create_layout.addWidget(QLabel("-"))
        
        self.create_end_entry = QLineEdit("10")
        self.create_end_entry.setFixedWidth(45)
        self.create_end_entry.setFixedHeight(24)
        self.create_end_entry.setPlaceholderText("结束")
        create_layout.addWidget(self.create_end_entry)
        
        create_layout.addWidget(QLabel("号 (输入编号)"))
        
        create_layout.addStretch()
        
        create_profile_button = QPushButton("开始创建分身")
        create_profile_button.clicked.connect(self.create_chrome_profiles)
        create_profile_button.setFixedHeight(30)
        create_profile_button.setMinimumWidth(120)
        create_profile_button.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
        create_layout.addWidget(create_profile_button)
        
        profile_layout.addLayout(create_layout)
        
        tab_widget.addTab(profile_tab, "创建分身")
        
        main_layout.addWidget(tab_widget)
        
        # === 进度条 ===
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setFixedHeight(18)
        main_layout.addWidget(self.progress_bar)
        
        # === 状态显示（精简版） ===
        status_frame = QFrame()
        status_frame.setFrameShape(QFrame.StyledPanel)
        status_frame.setFrameShadow(QFrame.Sunken)
        status_frame.setMinimumHeight(45)
        status_layout = QVBoxLayout(status_frame)
        status_layout.setContentsMargins(8, 6, 8, 6)
        
        self.status_label = QLabel("准备就绪")
        self.status_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.status_label.setWordWrap(True)
        self.status_label.setFont(QFont("Microsoft YaHei", 8))
        status_layout.addWidget(self.status_label)
        
        main_layout.addWidget(status_frame)
        
        # 创建状态栏
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("准备就绪")
    
        # 在UI创建完成后加载设置
        self.load_settings()
    
    def create_group_box(self, title):
        """创建带样式的分组框
        
        Args:
            title: 分组框标题
            
        Returns:
            QGroupBox: 样式化的分组框
        """
        group = QGroupBox(title)
        # 设置标题字体加粗且大小统一
        group.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        return group
    
    def create_button(self, text, callback, layout):
        """创建按钮
        
        Args:
            text: 按钮文本
            callback: 按钮点击回调函数
            layout: 要添加到的布局
            
        Returns:
            QPushButton: 样式化的按钮
        """
        button = QPushButton(text)
        button.clicked.connect(callback)
        button.setFixedHeight(40)
        button.setFont(QFont("Microsoft YaHei", 10, QFont.Normal))  # 确保不加粗
        layout.addWidget(button)
        return button
    
    def apply_light_theme(self):
        """应用浅色主题样式"""
        # 构建样式表
        self.setStyleSheet("""
            QWidget {
                background-color: #F0F0F0; /* 整体背景改为浅灰色 */
                color: #202020; /* 文字颜色改为深灰色 */
                font-size: 10pt;
            }
            
            QMainWindow {
                background-color: #E8E8E8; /* 主窗口背景稍深一点的浅灰 */
            }
            
            QGroupBox {
                background-color: #F8F8F8; /* GroupBox 背景更浅的灰色 */
                border: 1px solid #C0C0C0; /* 边框改为浅灰色 */
                border-radius: 5px;
                margin-top: 1.5em;
                font-weight: bold;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #101010; /* 标题颜色改为深色 */
            }
            
            QLabel {
                color: #202020; /* Label 文字颜色 */
                background-color: transparent;
            }
            
            QLineEdit {
                background-color: #FFFFFF; /* 输入框背景白色 */
                color: #101010; /* 输入框文字深色 */
                border: 1px solid #B0B0B0; /* 边框浅灰色 */
                border-radius: 3px;
                padding: 4px;
                selection-background-color: #ADD8E6; /* 选中文本背景浅蓝色 */
            }
            
            QPushButton {
                background-color: #0078D7; /* 按钮背景保持蓝色 */
                color: white; /* 按钮文字白色 */
                border: 1px solid #005FA3; /* 按钮边框深一点的蓝色 */
                border-radius: 3px;
                padding: 6px 12px;
                font-weight: normal; /* 按钮文字不加粗 */
            }
            
            QPushButton:hover {
                background-color: #1C97EA;
            }
            
            QPushButton:pressed {
                background-color: #00569C;
            }
            
            QProgressBar {
                border: 1px solid #C0C0C0; /* 边框浅灰色 */
                border-radius: 3px;
                background-color: #E0E0E0; /* 进度条背景浅灰色 */
                text-align: center;
                color: #202020; /* 进度条文字深色 */
            }
            
            QProgressBar::chunk {
                background-color: #0078D7; /* 进度条块保持蓝色 */
                width: 10px;
                margin: 0.5px;
            }
            
            QStatusBar {
                background-color: #D8D8D8; /* 状态栏背景浅灰色 */
                color: #202020; /* 状态栏文字深色 */
            }
            
            QScrollBar:vertical {
                border: 1px solid #C0C0C0;
                background: #F0F0F0;
                width: 10px;
                margin: 0px 0px 0px 0px;
            }
            
            QScrollBar::handle:vertical {
                background: #C0C0C0; /* 滚动条滑块颜色 */
                min-height: 20px;
            }
            
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: 1px solid #C0C0C0;
                background: #E0E0E0;
                height: 0px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }
        """)
    
    def set_status(self, message, color="white"):
        """设置状态标签的文本和颜色
        
        Args:
            message: 状态消息文本
            color: 状态颜色
        """
        self.status_label.setText(message)
        
        # 设置颜色
        if color == "green":
            self.status_label.setStyleSheet("color: #4caf50; font-size: 9pt;")
        elif color == "red":
            self.status_label.setStyleSheet("color: #f44336; font-size: 9pt;")
        elif color == "blue":
            self.status_label.setStyleSheet("color: #2196f3; font-size: 9pt;")
        elif color == "orange":
            self.status_label.setStyleSheet("color: #ff9800; font-size: 9pt;")
        else: # 默认颜色
            self.status_label.setStyleSheet("color: #202020; font-size: 9pt;") # 改为深色以适应浅色主题
    
    def set_status_formatted(self, current_launching=None, launched_list=None, not_launched_list=None, main_color="blue"):
        """设置格式化的三行状态显示
        
        Args:
            current_launching: 当前正在启动的分身编号
            launched_list: 已启动的分身编号列表
            not_launched_list: 未启动的分身编号列表
            main_color: 主要颜色
        """
        status_lines = []
        
        # 第一行：正在启动的分身
        if current_launching is not None:
            status_lines.append(f"🎯 正在启动分身：{current_launching}")
        else:
            status_lines.append("🎯 启动操作完成")
        
        # 第二行：已启动分身
        if launched_list and len(launched_list) > 0:
            launched_str = "、".join(map(str, sorted(launched_list)))
            status_lines.append(f"✅已启动分身：{launched_str}")
        else:
            status_lines.append("✅已启动分身：无")
        
        # 第三行：未启动分身
        if not_launched_list and len(not_launched_list) > 0:
            not_launched_str = "、".join(map(str, sorted(not_launched_list)))
            status_lines.append(f"❌未启动分身：{not_launched_str}")
        else:
            status_lines.append("❌未启动分身：无")
        
        # 合并为多行文本
        formatted_message = "\n".join(status_lines)
        self.set_status(formatted_message, main_color)
    
    def launch_random_browsers(self):
        """随机启动指定数量的Chrome分身"""
        try:
            # 获取范围设置
            start = int(self.start_num.text().strip())
            end = int(self.end_num.text().strip())
            count = int(self.num_browsers.text().strip())
            
            # 获取指定的文件夹路径
            folder_path = self.folder_path.text().strip()
            
            # 确认文件夹存在
            if not os.path.exists(folder_path):
                self.set_status(f"错误: 文件夹路径 {folder_path} 不存在", "red")
                self.statusBar.showMessage(f"错误: 文件夹路径不存在")
                return
                
            # 文件夹中的可用快捷方式
            available_files = [f for f in os.listdir(folder_path) if f.endswith('.lnk')]
            available_numbers = [int(f.split('.')[0]) for f in available_files if f.split('.')[0].isdigit()]
            
            # 过滤范围内的编号
            in_range_numbers = [n for n in available_numbers if start <= n <= end]
            
            # 从已有编号中排除已启动的编号
            available_to_select_from = [n for n in in_range_numbers if n not in self.launched_numbers]
            
            # 记录本次操作的完整目标范围（去重前，有效快捷方式）
            self._last_operation_scope_profiles = list(in_range_numbers) # 应该是 in_range_numbers，代表范围内的所有有效快捷方式

            if not available_to_select_from:
                self.set_status("没有可用的分身编号，所有分身可能已经启动", "orange")
                self.statusBar.showMessage("没有可用的分身编号")
                return
            
            # 如果可用数量少于请求的数量，调整数量
            count = min(count, len(available_to_select_from))
            
            # 随机选择指定数量的编号
            selected_numbers = random.sample(available_to_select_from, count)
            selected_numbers.sort()  # 排序以便按顺序启动
            
            # 计算未启动的分身（范围内的所有可用分身减去已启动的）
            not_launched_in_range = [n for n in in_range_numbers if n not in self.launched_numbers and n not in selected_numbers]
            
            # 显示初始状态
            self.set_status_formatted(
                current_launching="准备中",
                launched_list=list(self.launched_numbers),
                not_launched_list=not_launched_in_range + selected_numbers,
                main_color="blue"
            )
            
            # 显示进度条
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            
            # 创建并启动工作线程
            self.worker = BackgroundWorker(selected_numbers, folder_path, self.delay_time.text())
            self.worker.update_status.connect(self.on_launch_status_update)
            self.worker.update_progress.connect(self.progress_bar.setValue)
            self.worker.finished.connect(self.on_launch_finished)
            self.worker.start()
            
        except ValueError:
            self.set_status("错误: 请输入有效的数字", "red")
        except Exception as e:
            self.set_status(f"错误: {str(e)}", "red")
            self.statusBar.showMessage(f"错误: {str(e)}")
    
    def on_launch_status_update(self, message, color):
        """处理启动过程中的状态更新"""
        # 解析特殊格式的状态消息
        if "CURRENT_LAUNCHING:" in message:
            try:
                # 解析消息格式: CURRENT_LAUNCHING:5|LAUNCHED:1,2,3|REMAINING:6,7,8
                parts = message.split("|")
                current_launching = None
                launched_list = []
                remaining_list = []
                
                for part in parts:
                    if part.startswith("CURRENT_LAUNCHING:"):
                        current_launching = int(part.split(":")[1])
                    elif part.startswith("LAUNCHED:"):
                        launched_str = part.split(":", 1)[1]
                        if launched_str:
                            launched_list = [int(x) for x in launched_str.split(",") if x.strip()]
                    elif part.startswith("REMAINING:"):
                        remaining_str = part.split(":", 1)[1] 
                        if remaining_str:
                            remaining_list = [int(x) for x in remaining_str.split(",") if x.strip()]
                
                # 合并当前已启动的分身和新启动的分身
                all_launched = list(self.launched_numbers) + launched_list
                
                # 显示格式化状态（这是主要日志，不会被覆盖）
                self.set_status_formatted(
                    current_launching=current_launching,
                    launched_list=all_launched,
                    not_launched_list=remaining_list,
                    main_color="blue"
                )
                
                # 过程信息只在状态栏显示，不覆盖主要日志
                if current_launching is not None:
                    self.statusBar.showMessage(f"正在启动分身 {current_launching}...")
                
            except Exception as e:
                # 如果解析失败，过程日志只在状态栏显示
                self.statusBar.showMessage(f"启动过程: {message}")
        else:
            # 普通的警告/错误消息可以显示在主区域，但要简短
            if any(keyword in message for keyword in ["错误", "警告", "失败"]):
                self.set_status(message, color)
            else:
                # 其他过程信息只在状态栏显示
                self.statusBar.showMessage(message)
    
    def on_launch_finished(self, status_text_from_worker, color, successful_numbers):
        """启动完成后的回调
        
        Args:
            status_text_from_worker: 来自 BackgroundWorker 的原始状态文本
            color: 状态颜色
            successful_numbers: 成功启动的编号列表
        """
        self.progress_bar.setVisible(False)
        
        # 更新已启动编号集合
        for num in successful_numbers:
            self.launched_numbers.add(num)
        if successful_numbers: # 只有成功启动了才保存
            self.save_settings()

        # 计算当前范围内未启动的分身
        remaining_count = self._get_remaining_in_range_count(self._last_operation_scope_profiles)
        not_launched_in_range = []
        if self._last_operation_scope_profiles:
            not_launched_in_range = [n for n in self._last_operation_scope_profiles if n not in self.launched_numbers]
        
        # 显示最终的三行格式化状态（主要日志）
        if successful_numbers:
            self.set_status_formatted(
                current_launching=None,  # 启动完成
                launched_list=list(self.launched_numbers),
                not_launched_list=not_launched_in_range,
                main_color="green"
            )
            
            # 成功消息只在状态栏显示，不覆盖主要日志
            launched_str = "、".join(map(str, sorted(successful_numbers)))
            self.statusBar.showMessage(f"✅ 成功启动分身: {launched_str} (总计: {len(self.launched_numbers)}个)")
        else:
            self.set_status_formatted(
                current_launching=None,
                launched_list=list(self.launched_numbers),
                not_launched_list=not_launched_in_range,
                main_color="orange"
            )
            # 失败消息在状态栏显示
            self.statusBar.showMessage("❌ 未成功启动任何新分身")
        
        self._last_operation_scope_profiles = [] # 清理
        
        # 自动应用图标（如果启用且有成功启动的分身）
        if successful_numbers and ICON_MANAGEMENT_AVAILABLE and self.icon_manager and self.auto_apply_icons:
            self.apply_icons_for_numbers(successful_numbers)
    
    def launch_specific_range(self):
        """启动指定范围的Chrome分身"""
        try:
            range_text = self.specific_range.text().strip()
            if '-' not in range_text:
                self.set_status("错误: 请使用正确的格式 (例如: 1-10)", "red")
                return
                
            start, end = map(int, range_text.split('-'))
            
            # 获取指定的文件夹路径
            folder_path = self.folder_path.text().strip()
            
            # 确认文件夹存在
            if not os.path.exists(folder_path):
                self.set_status(f"错误: 文件夹路径 {folder_path} 不存在", "red")
                self.statusBar.showMessage(f"错误: 文件夹路径不存在")
                return
            
            # 动态获取文件夹中的快捷方式数量
            available_files = [f for f in os.listdir(folder_path) if f.endswith('.lnk')]
            available_numbers = set()
            for f in available_files:
                name = f.split('.')[0]
                if name.isdigit():
                    available_numbers.add(int(name))
            
            # 输入验证
            if not available_numbers:
                self.set_status("错误: 文件夹中没有可用的快捷方式", "red")
                return
            
            # 检查范围是否有效
            if start < 1 or end < 1:
                self.set_status("错误: 范围必须为正整数", "red")
                return
            if start > end:
                self.set_status("错误: 起始值不能大于结束值", "red")
                return
            
            # 获取要打开的编号列表（但仅限于文件夹中存在的快捷方式）
            numbers_in_specified_range = [n for n in range(start, end + 1) if n in available_numbers]
            
            self._last_operation_scope_profiles = list(numbers_in_specified_range)

            if not numbers_in_specified_range:
                self.set_status(f"错误: 指定范围内 {range_text} 没有找到有效的快捷方式", "red")
                self._last_operation_scope_profiles = [] # 清理以防误报
                return
            
            # 从 identified numbers_in_specified_range 中选出真正要启动的 (那些尚未在 self.launched_numbers 中的)
            numbers_to_launch = [n for n in numbers_in_specified_range if n not in self.launched_numbers]

            if not numbers_to_launch:
                # 显示当前状态（全部已启动）
                self.set_status_formatted(
                    current_launching=None,
                    launched_list=list(self.launched_numbers),
                    not_launched_list=[],
                    main_color="orange"
                )
                self.statusBar.showMessage("范围内均已启动")
                self._last_operation_scope_profiles = [] # 清理
                return

            # 计算未启动的分身（范围内减去即将启动的）
            not_launched_in_range = [n for n in numbers_in_specified_range if n not in self.launched_numbers]
            
            # 显示初始状态
            self.set_status_formatted(
                current_launching="准备中",
                launched_list=list(self.launched_numbers),
                not_launched_list=not_launched_in_range,
                main_color="blue"
            )
            self.statusBar.showMessage(f"正在启动Chrome分身，范围: {start}-{end}")
            
            # 显示进度条
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            
            # 创建并启动工作线程
            self.worker = BackgroundWorker(numbers_to_launch, folder_path, self.delay_time.text())
            self.worker.update_status.connect(self.on_launch_status_update)
            self.worker.update_progress.connect(self.progress_bar.setValue)
            self.worker.finished.connect(self.on_launch_finished)
            self.worker.start()
                
        except ValueError:
            self.set_status("错误: 请使用正确的格式 (例如: 1-10)", "red")
            self.statusBar.showMessage("错误: 请使用正确的格式 (例如: 1-10)")
        except Exception as e:
            self.set_status(f"错误: {str(e)}", "red")
            self.statusBar.showMessage(f"错误: {str(e)}")
    
    def close_all_chrome(self):
        """关闭所有Chrome窗口"""
        try:
            reply = QMessageBox.question(self, '确认操作', 
                                      '确定要关闭所有Chrome窗口吗?',
                                      QMessageBox.Yes | QMessageBox.No, 
                                      QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                # 尝试先进行优雅关闭，去掉 /F 参数
                os.system('taskkill /IM chrome.exe')
                # 可以选择性地在这里加一个短暂的 sleep，给进程响应时间，例如 time.sleep(3)
                self.set_status("已发送关闭所有Chrome窗口的请求", "green")
                self.statusBar.showMessage("已发送关闭所有Chrome窗口的请求")
                # 清空已启动编号集合
                self.launched_numbers.clear()
                self.save_settings() # 保存清空后的状态
                
        except Exception as e:
            self.set_status(f"关闭Chrome失败: {str(e)}", "red")
            self.statusBar.showMessage(f"关闭Chrome失败: {str(e)}")
    
    def close_specific_range(self):
        """关闭指定范围的Chrome分身"""
        try:
            range_text = self.specific_range.text().strip()
            if '-' not in range_text:
                self.set_status("错误: 请使用正确的格式 (例如: 1-10)", "red")
                return
                
            start, end = map(int, range_text.split('-'))
            
            # 获取指定的文件夹路径
            folder_path = self.folder_path.text().strip()
            
            # 确认文件夹存在
            if not os.path.exists(folder_path):
                self.set_status(f"错误: 文件夹路径 {folder_path} 不存在", "red")
                self.statusBar.showMessage(f"错误: 文件夹路径不存在")
                return
            
            # 动态获取文件夹中的快捷方式数量
            available_files = [f for f in os.listdir(folder_path) if f.endswith('.lnk')]
            available_numbers = set()
            for f in available_files:
                name = f.split('.')[0]
                if name.isdigit():
                    available_numbers.add(int(name))
            
            # 输入验证
            if not available_numbers:
                self.set_status("错误: 文件夹中没有可用的快捷方式", "red")
                return
            
            # 检查范围是否有效
            if start < 1 or end < 1:
                self.set_status("错误: 范围必须为正整数", "red")
                return
            if start > end:
                self.set_status("错误: 起始值不能大于结束值", "red")
                return
            
            # 获取要关闭的编号列表
            numbers = list(range(start, end + 1))
            
            # 更新状态
            self.set_status(f"正在关闭Chrome分身 (范围: {start}-{end})...", "blue")
            self.statusBar.showMessage(f"正在关闭Chrome分身，范围: {start}-{end}")
            

            # 显0.示进度条
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            
            # 创建并启动工作线程
            self.worker = BackgroundWorker(numbers, folder_path, self.delay_time.text(), mode="close")
            self.worker.update_status.connect(lambda msg, color: self.set_status(msg, color))
            self.worker.update_progress.connect(self.progress_bar.setValue)
            self.worker.finished.connect(self.on_close_finished)
            self.worker.start()
                
        except ValueError:
            self.set_status("错误: 请使用正确的格式 (例如: 1-10)", "red")
            self.statusBar.showMessage("错误: 请使用正确的格式 (例如: 1-10)")
        except Exception as e:
            self.set_status(f"错误: {str(e)}", "red")
            self.statusBar.showMessage(f"错误: {str(e)}")
    
    def on_close_finished(self, status_text, color, closed_numbers):
        """关闭完成后的回调
        
        Args:
            status_text: 状态文本
            color: 状态颜色
            closed_numbers: 成功关闭的编号列表
        """
        self.progress_bar.setVisible(False)
        
        # 从已启动编号集合中移除已关闭的编号
        for num in closed_numbers:
            if num in self.launched_numbers:
                self.launched_numbers.remove(num)
        if closed_numbers: # 只有成功关闭了才保存
            self.save_settings()

        # 关闭操作的结果消息在状态栏显示，主要状态区域显示当前已启动分身状态
        if closed_numbers:
            # 显示当前剩余已启动分身的状态
            self.set_status_formatted(
                current_launching=None,
                launched_list=list(self.launched_numbers),
                not_launched_list=[],  # 关闭操作后不显示未启动列表
                main_color="green"
            )
            
            # 关闭成功的消息在状态栏显示
            closed_str = "、".join(map(str, sorted(closed_numbers)))
            self.statusBar.showMessage(f"✅ 成功关闭分身: {closed_str} (当前剩余: {len(self.launched_numbers)}个)")
        else:
            # 如果没有成功关闭任何分身，显示当前状态
            self.set_status_formatted(
                current_launching=None,
                launched_list=list(self.launched_numbers),
                not_launched_list=[],
                main_color="orange"
            )
            self.statusBar.showMessage("❌ 未成功关闭任何指定分身")
    
    def on_open_url_finished(self, status_text, color, successful_instances_info):
        """打开URL完成后的回调
        
        Args:
            status_text: 状态文本
            color: 状态颜色
            successful_instances_info: 成功打开URL的实例信息列表 (UDD basenames)
        """
        self.progress_bar.setVisible(False)
        
        # URL操作完成后，保持当前主要状态不变，只在状态栏显示操作结果
        if "成功" in status_text:
            self.statusBar.showMessage(f"✅ {status_text.split('!')[0]}! (编号分身: {len(self.launched_numbers)}个)")
        else:
            self.statusBar.showMessage(f"❌ {status_text} (编号分身: {len(self.launched_numbers)}个)")
    
    def _get_remaining_in_range_count(self, profiles_in_scope_list):
        """计算指定分身编号列表中，还有多少个是未被全局启动的。"""
        if not profiles_in_scope_list:
            return 0
        # 确保 self.launched_numbers 是最新的
        return len([p for p in profiles_in_scope_list if p not in self.launched_numbers])

    def launch_sequentially(self):
        """依次启动指定范围内的Chrome分身"""
        try:
            folder_path = self.folder_path.text().strip()
            if not os.path.exists(folder_path):
                self.set_status(f"错误: 快捷方式文件夹路径 {folder_path} 不存在", "red")
                self.statusBar.showMessage("错误: 快捷方式文件夹路径不存在")
                return

            range_text = self.specific_range.text().strip()
            if '-' not in range_text:
                self.set_status("错误: 请使用正确的范围格式 (例如: 1-10)", "red")
                self.statusBar.showMessage("错误: 范围格式不正确")
                return
            
            start_s, end_s = range_text.split('-')
            start_num = int(start_s)
            end_num = int(end_s)

            if start_num < 1 or end_num < 1 or start_num > end_num:
                self.set_status("错误: 范围值无效，必须为正整数且起始不大于结束", "red")
                self.statusBar.showMessage("错误: 范围值无效")
                return
            
            # 序列初始化/重新初始化逻辑
            if not self.sequential_launch_range_active or self.sequential_launch_active_range_str != range_text:
                self.set_status(f"正在初始化依次启动序列 (范围: {range_text})...", "blue")
                available_files = [f for f in os.listdir(folder_path) if f.endswith('.lnk')]
                available_numbers_in_folder = sorted([int(f.split('.')[0]) for f in available_files if f.split('.')[0].isdigit()])
                
                self.sequential_launch_profiles = [n for n in available_numbers_in_folder if start_num <= n <= end_num]
                
                if not self.sequential_launch_profiles:
                    self.set_status(f"错误: 在指定范围 {range_text} 内没有找到有效的快捷方式。", "orange")
                    self.statusBar.showMessage("指定范围内无有效分身")
                    self.sequential_launch_range_active = False # 确保重置
                    return
                
                self.sequential_launch_current_index = 0
                self.sequential_launch_range_active = True
                self.sequential_launch_active_range_str = range_text
                
                remaining_to_initially_launch = self._get_remaining_in_range_count(self.sequential_launch_profiles)
                if remaining_to_initially_launch > 0:
                    self.set_status(f"依次启动序列已初始化 (范围: {range_text})。下一个: {self.sequential_launch_profiles[0]}。此范围还剩 {remaining_to_initially_launch} 个待启动。", "blue")
                    self.statusBar.showMessage(f"序列初始化。下一个: {self.sequential_launch_profiles[0]}. 还剩 {remaining_to_initially_launch}")
                else:
                    self.set_status(f"依次启动序列 (范围: {range_text}) 内分身均已启动或无有效分身。还剩 0 个待启动。", "orange")
                    self.statusBar.showMessage("序列中无待启动分身")
                    self.sequential_launch_range_active = False # 无可启动，则序列非激活
                    return # 避免后续逻辑出错
            
            # 执行序列启动逻辑
            if not self.sequential_launch_range_active: # 如果初始化失败或序列已被标记为非激活
                self.set_status("请先使用有效范围初始化依次启动序列。", "orange")
                self.statusBar.showMessage("序列未激活")
                return

            # 跳过已启动的分身
            while self.sequential_launch_current_index < len(self.sequential_launch_profiles) and \
                  self.sequential_launch_profiles[self.sequential_launch_current_index] in self.launched_numbers:
                skipped_profile = self.sequential_launch_profiles[self.sequential_launch_current_index]
                self.sequential_launch_current_index += 1
                remaining_after_skip = 0
                if self.sequential_launch_range_active: # 仅当序列仍激活时计算
                    # 从当前索引开始计算剩余
                    remaining_after_skip = len([p for p in self.sequential_launch_profiles[self.sequential_launch_current_index:] if p not in self.launched_numbers])
                self.set_status(f"分身 {skipped_profile} 已启动，自动跳过。还剩 {remaining_after_skip} 个待启动。", "orange")
                self.statusBar.showMessage(f"已跳过 {skipped_profile}。还剩 {remaining_after_skip}")
            
            # 检查序列是否完成
            if self.sequential_launch_current_index >= len(self.sequential_launch_profiles):
                self.set_status(f"依次启动序列 (范围: {self.sequential_launch_active_range_str}) 已全部处理完毕。还剩 0 个待启动。", "green")
                self.statusBar.showMessage("依次启动序列完成")
                self.sequential_launch_range_active = False # 重置序列
                return
            
            # 启动下一个分身
            profile_to_launch = self.sequential_launch_profiles[self.sequential_launch_current_index]
            self._currently_attempting_sequential_profile = profile_to_launch

            self.set_status(f"正在尝试依次启动分身: {profile_to_launch}...", "blue")
            self.statusBar.showMessage(f"正在启动: {profile_to_launch}")
            self.progress_bar.setValue(0) # 重置进度条以用于单次启动
            self.progress_bar.setVisible(True)

            self.worker = BackgroundWorker([profile_to_launch], folder_path, self.delay_time.text())
            self.worker.update_status.connect(lambda msg, color: self.set_status(msg, color)) # 可以简化或移除，因为主要状态由 on_sequential_launch_item_finished 控制
            self.worker.update_progress.connect(self.progress_bar.setValue)
            self.worker.finished.connect(self.on_sequential_launch_item_finished)
            self.worker.start()

        except ValueError:
            self.set_status("错误: 范围中请输入有效的数字", "red")
            self.statusBar.showMessage("错误: 范围数字无效")
            # 重置序列状态，以防部分初始化
            self.sequential_launch_range_active = False
            self.sequential_launch_active_range_str = None
            self.sequential_launch_profiles = []
            self.sequential_launch_current_index = 0
        except Exception as e:
            self.set_status(f"依次启动时发生意外错误: {str(e)}", "red")
            self.statusBar.showMessage(f"意外错误: {str(e)}")
            # 重置序列状态
            self.sequential_launch_range_active = False
            self.sequential_launch_active_range_str = None
            self.sequential_launch_profiles = []
            self.sequential_launch_current_index = 0
            import traceback
            traceback.print_exc()

    def on_sequential_launch_item_finished(self, status_text_from_worker, color_from_worker, successful_numbers_from_worker):
        """单个依次启动项完成后的回调"""
        profile_launched_attempt = self._currently_attempting_sequential_profile
        self._currently_attempting_sequential_profile = None # 清理
        self.progress_bar.setVisible(False)

        if profile_launched_attempt is None: # 不应该发生，但作为保险
            self.set_status("依次启动回调错误：未记录尝试的分身", "red")
            return

        if successful_numbers_from_worker and profile_launched_attempt in successful_numbers_from_worker:
            self.launched_numbers.add(profile_launched_attempt)
            self.sequential_launch_current_index += 1
            self.save_settings() # 保存状态
            
            remaining_in_seq = 0
            next_profile_to_show = None
            remaining_profiles = []
            if self.sequential_launch_range_active: # 仅当序列仍认为自己是激活状态时计算
                # 计算从当前 self.sequential_launch_current_index 开始的剩余未启动项
                for i in range(self.sequential_launch_current_index, len(self.sequential_launch_profiles)):
                    p_num = self.sequential_launch_profiles[i]
                    if p_num not in self.launched_numbers:
                        remaining_profiles.append(p_num)
                remaining_in_seq = len(remaining_profiles)
                if remaining_profiles:
                    next_profile_to_show = remaining_profiles[0]
            
            # 显示成功启动的三行格式化状态
            if self.sequential_launch_range_active and remaining_in_seq > 0 and next_profile_to_show is not None:
                self.set_status_formatted(
                    current_launching=f"下一个: {next_profile_to_show}",
                    launched_list=list(self.launched_numbers),
                    not_launched_list=remaining_profiles,
                    main_color="green"
                )
                # 过程信息在状态栏显示
                self.statusBar.showMessage(f"✅ 成功启动 {profile_launched_attempt}，下一个: {next_profile_to_show} (还剩 {remaining_in_seq})")
            elif self.sequential_launch_range_active and remaining_in_seq == 0: # 刚启动的是最后一个
                self.set_status_formatted(
                    current_launching=None,
                    launched_list=list(self.launched_numbers),
                    not_launched_list=[],
                    main_color="green"
                )
                # 完成信息在状态栏显示
                self.statusBar.showMessage(f"✅ 成功启动 {profile_launched_attempt}，依次启动序列完成")
                self.sequential_launch_range_active = False # 重置序列
            else: # 序列已完成或不再激活 (例如在 launch_sequentially 中被重置)
                self.set_status_formatted(
                    current_launching=None,
                    launched_list=list(self.launched_numbers),
                    not_launched_list=[],
                    main_color="green"
                )
                # 完成信息在状态栏显示
                self.statusBar.showMessage(f"✅ 成功启动 {profile_launched_attempt}，序列完成")
                self.sequential_launch_range_active = False # 确保重置
            
            # 自动应用图标（如果启用）
            if ICON_MANAGEMENT_AVAILABLE and self.icon_manager and self.auto_apply_icons:
                self.apply_icons_for_numbers([profile_launched_attempt])

        else: # 启动失败
            remaining_on_fail = 0
            remaining_profiles_on_fail = []
            if self.sequential_launch_range_active:
                 # 从当前索引（未改变）开始计算
                 for i in range(self.sequential_launch_current_index, len(self.sequential_launch_profiles)):
                    p_num = self.sequential_launch_profiles[i]
                    if p_num not in self.launched_numbers:
                        remaining_profiles_on_fail.append(p_num)
                 remaining_on_fail = len(remaining_profiles_on_fail)
            
            # 显示失败的三行格式化状态
            self.set_status_formatted(
                current_launching=f"失败: {profile_launched_attempt}",
                launched_list=list(self.launched_numbers),
                not_launched_list=remaining_profiles_on_fail,
                main_color="red"
            )
            
            # 失败信息在状态栏显示
            if self.sequential_launch_range_active and remaining_on_fail > 0:
                self.statusBar.showMessage(f"❌ 分身 {profile_launched_attempt} 启动失败，序列还剩 {remaining_on_fail}")
            elif self.sequential_launch_range_active: # 意味着 remaining_on_fail is 0, but sequence was active
                 self.statusBar.showMessage(f"❌ 分身 {profile_launched_attempt} 启动失败，序列可能已无后续")
            else: # sequence not active
                self.statusBar.showMessage(f"❌ 分身 {profile_launched_attempt} 启动失败")

    def open_url_in_running(self):
        """在所有已打开的Chrome实例中打开网址 (新思路：直接操作窗口/进程，不再依赖分身编号)"""
        try:
            url = self.url_entry.text().strip()
            if not url:
                self.set_status("错误: 请输入有效的网址", "red")
                self.statusBar.showMessage("错误: 网址不能为空")
                return
                
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
            
            # target_profiles_info 将存储 (用户数据目录路径, 该实例的chrome.exe路径) 元组
            # 使用字典以 user_data_dir 作为键来自动去重，值为 chrome_exe_path
            unique_running_instances = {}
            
            user_data_dir_pattern = self.user_data_dir_pattern # 使用在 __init__ 中定义的模式
            processed_pids = set()

            print("DEBUG_OPEN_URL_NEW: Starting to scan running Chrome processes...")

            for proc in psutil.process_iter(attrs=['pid', 'name', 'cmdline', 'exe'], ad_value=None):
                try:
                    if proc.pid in processed_pids:
                        continue
                    
                    proc_info = proc.info
                    if not proc_info or \
                       not proc_info.get('name') or \
                       'chrome.exe' not in proc_info['name'].lower() or \
                       not proc_info.get('cmdline') or \
                       not proc_info.get('exe'):
                        continue
                    
                    cmd_line_list = proc_info['cmdline']
                    cmd_line_str = " ".join(cmd_line_list)
                    proc_exe_path = proc_info['exe']

                    print(f"DEBUG_OPEN_URL_NEW: PID={proc.pid}, EXE='{proc_exe_path}', CMD_LINE='{cmd_line_str}'")

                    actual_user_data_dir = None
                    match_ud_dir = user_data_dir_pattern.search(cmd_line_str)

                    if match_ud_dir:
                        # 提取路径，优先尝试带引号的路径组，然后尝试不带引号的路径组
                        raw_udd = match_ud_dir.group('path') or match_ud_dir.group('path_unquoted')
                        if raw_udd:
                            actual_user_data_dir = os.path.normpath(raw_udd.strip())
                            print(f"DEBUG_OPEN_URL_NEW: PID={proc.pid}, Extracted UDD: '{actual_user_data_dir}'")
                            
                            # 确保提取到的 user_data_dir 是一个有效的目录，并且 proc_exe_path 存在
                            if actual_user_data_dir and os.path.isdir(actual_user_data_dir) and \
                               proc_exe_path and os.path.exists(proc_exe_path):
                                # 以 user_data_dir 为键，如果已存在，则不更新（保留第一个遇到的exe_path）
                                # 或者可以更新为最新的，或者基于某种逻辑选择，这里简单保留第一个
                                if actual_user_data_dir not in unique_running_instances:
                                    unique_running_instances[actual_user_data_dir] = proc_exe_path
                                    print(f"DEBUG_OPEN_URL_NEW: PID={proc.pid}, ADDED instance: UDD='{actual_user_data_dir}', EXE='{proc_exe_path}'")
                                else:
                                    print(f"DEBUG_OPEN_URL_NEW: PID={proc.pid}, Instance already recorded for UDD: '{actual_user_data_dir}'")
                            else:
                                print(f"DEBUG_OPEN_URL_NEW: PID={proc.pid}, SKIPPED - Invalid UDD ('{actual_user_data_dir}', is_dir={os.path.isdir(actual_user_data_dir) if actual_user_data_dir else 'N/A'}) or EXE path ('{proc_exe_path}', exists={os.path.exists(proc_exe_path) if proc_exe_path else 'N/A'}).")
                        else:
                            print(f"DEBUG_OPEN_URL_NEW: PID={proc.pid}, SKIPPED - Could not extract UDD string from regex match.")
                    else:
                        print(f"DEBUG_OPEN_URL_NEW: PID={proc.pid}, SKIPPED - user_data_dir_pattern did not match.")

                    processed_pids.add(proc.pid)

                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue 
                except (TypeError, ValueError, AttributeError) as e_data_handling:
                    print(f"DEBUG_OPEN_URL_NEW_DATA_HANDLING_ERROR: PID (approx {proc.pid if proc and hasattr(proc, 'pid') else 'N/A'}): {str(e_data_handling)}")
                    continue 
                except Exception as e_general_proc:
                    print(f"DEBUG_OPEN_URL_NEW_GENERAL_ERROR: PID (approx {proc.pid if proc and hasattr(proc, 'pid') else 'N/A'}): {str(e_general_proc)}")
                    continue
            
            final_target_instances = list(unique_running_instances.items()) # 转换为 [(udd, exe_path), ...] 列表
            # Sort by UDD path for consistent ordering, though not strictly necessary for functionality
            final_target_instances.sort(key=lambda x: x[0]) 
            
            print(f"DEBUG_OPEN_URL_NEW: FINAL target instances for URL opening: {final_target_instances}")

            if not final_target_instances:
                self.set_status("没有找到正在运行的、可操作的Chrome实例", "orange")
                self.statusBar.showMessage("没有找到正在运行的Chrome实例")
                return

            self.set_status(f"准备在 {len(final_target_instances)} 个已运行Chrome实例中打开网址...", "blue")
            self.statusBar.showMessage(f"正在打开网址: {url}")
            
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            
            # BackgroundWorker 的 profiles_data 现在是 [(user_data_dir, exe_path), ...]
            self.worker = BackgroundWorker(profiles_data=final_target_instances, 
                                         folder_path=None, 
                                         delay_time=self.delay_time.text(), 
                                         mode="open_url", url=url)
            self.worker.update_status.connect(lambda msg, color: self.set_status(msg, color))
            self.worker.update_progress.connect(self.progress_bar.setValue)
            self.worker.finished.connect(self.on_open_url_finished)
            self.worker.start()
                
        except ValueError as ve:
            self.set_status(f"输入错误: {str(ve)}", "red")
            self.statusBar.showMessage(f"输入错误: {str(ve)}")
        except Exception as e_global:
            self.set_status(f"打开URL时发生全局错误: {str(e_global)}", "red")
            self.statusBar.showMessage(f"全局错误: {str(e_global)}")
            import traceback
            traceback.print_exc()
    
    def browse_folder(self):
        """打开文件夹浏览对话框选择分身快捷方式目录"""
        current_path = self.folder_path.text().strip()
        folder = QFileDialog.getExistingDirectory(self, "选择分身快捷方式目录", 
                                                current_path)
        if folder:
            self.folder_path.setText(folder)
            # 保存设置
            self.save_settings()
            
    def save_settings(self):
        """保存设置到配置文件"""
        try:
            # 检查基础UI组件是否存在
            if hasattr(self, 'folder_path') and self.folder_path:
                self.settings.setValue("folder_path", self.folder_path.text())
            if hasattr(self, 'start_num') and self.start_num:
                self.settings.setValue("start_num", self.start_num.text())
            if hasattr(self, 'end_num') and self.end_num:
                self.settings.setValue("end_num", self.end_num.text())
            if hasattr(self, 'num_browsers') and self.num_browsers:
                self.settings.setValue("num_browsers", self.num_browsers.text())
            if hasattr(self, 'delay_time') and self.delay_time:
                self.settings.setValue("delay_time", self.delay_time.text())
            if hasattr(self, 'specific_range') and self.specific_range:
                self.settings.setValue("specific_range", self.specific_range.text())
            if hasattr(self, 'url_entry') and self.url_entry:
                self.settings.setValue("url_entry", self.url_entry.text())
            
            # 保存依次启动的状态
            if hasattr(self, 'sequential_launch_range'):
                self.settings.setValue("sequential_launch_range", self.sequential_launch_range or "")
            if hasattr(self, 'sequential_launch_current_index'):
                self.settings.setValue("sequential_launch_index", self.sequential_launch_current_index)
            if hasattr(self, 'sequential_launch_active'):
                self.settings.setValue("sequential_launch_active", self.sequential_launch_active)
            
            # 保存图标设置
            if hasattr(self, 'auto_apply_checkbox') and self.auto_apply_checkbox:
                self.settings.setValue("auto_apply_icons", self.auto_apply_checkbox.isChecked())
            
            # 保存创建分身的路径设置
            if hasattr(self, 'shortcut_path_entry') and self.shortcut_path_entry:
                self.settings.setValue("shortcut_creation_path", self.shortcut_path_entry.text())
            if hasattr(self, 'cache_path_entry') and self.cache_path_entry:
                self.settings.setValue("cache_creation_path", self.cache_path_entry.text())
            if hasattr(self, 'create_start_entry') and self.create_start_entry:
                self.settings.setValue("create_start", self.create_start_entry.text())
            if hasattr(self, 'create_end_entry') and self.create_end_entry:
                self.settings.setValue("create_end", self.create_end_entry.text())
            
            print("设置已保存")
        except Exception as e:
            print(f"保存设置失败: {e}")
    
    def load_settings(self):
        """从配置文件加载设置"""
        try:
            # 等待UI组件创建完成后再加载设置
            if hasattr(self, 'folder_path'):
                self.folder_path.setText(self.settings.value("folder_path", ""))
                self.start_num.setText(self.settings.value("start_num", "1"))
                self.end_num.setText(self.settings.value("end_num", "100"))
                self.num_browsers.setText(self.settings.value("num_browsers", "5"))
                self.delay_time.setText(self.settings.value("delay_time", "0.5"))
                self.specific_range.setText(self.settings.value("specific_range", "1-10"))
                self.url_entry.setText(self.settings.value("url_entry", "https://www.example.com"))
                
                # 加载依次启动的状态
                self.sequential_launch_range = self.settings.value("sequential_launch_range", "")
                self.sequential_launch_current_index = self.settings.value("sequential_launch_index", 0, type=int)
                self.sequential_launch_active = self.settings.value("sequential_launch_active", False, type=bool)
                
                # 加载图标设置
                if hasattr(self, 'auto_apply_checkbox'):
                    self.auto_apply_icons = self.settings.value("auto_apply_icons", True, type=bool)
                    self.auto_apply_checkbox.setChecked(self.auto_apply_icons)
                
                # 加载创建分身的路径设置
                if hasattr(self, 'shortcut_path_entry'):
                    saved_shortcut_path = self.settings.value("shortcut_creation_path", "")
                    self.shortcut_path_entry.setText(saved_shortcut_path)
                    self.shortcut_creation_path = saved_shortcut_path
                
                if hasattr(self, 'cache_path_entry'):
                    saved_cache_path = self.settings.value("cache_creation_path", "")
                    self.cache_path_entry.setText(saved_cache_path)
                    self.cache_creation_path = saved_cache_path
                
                if hasattr(self, 'create_start_entry'):
                    self.create_start_entry.setText(self.settings.value("create_start", "1"))
                
                if hasattr(self, 'create_end_entry'):
                    self.create_end_entry.setText(self.settings.value("create_end", "10"))
                
                print("设置已加载")
        except Exception as e:
            print(f"加载设置失败: {e}")
    
    def closeEvent(self, event):
        """程序关闭时的处理"""
        # 保存设置
        self.save_settings()
        
        # 停止所有正在运行的工作线程
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
        
        if hasattr(self, 'icon_worker') and self.icon_worker.isRunning():
            self.icon_worker.terminate()
            self.icon_worker.wait()
        
        if hasattr(self, 'profile_worker') and self.profile_worker.isRunning():
            self.profile_worker.terminate()
            self.profile_worker.wait()
        
        event.accept()

    def _sync_launched_numbers_with_running_processes(self):
        """
        扫描当前运行的Chrome进程，并用实际运行的分身编号更新 self.launched_numbers。
        这确保了程序状态与系统实际状态的一致性，特别是在手动关闭分身或程序异常退出后。
        """
        print("DEBUG: Synchronizing launched_numbers with running processes...")
        actually_running_profiles = set()
        try:
            user_data_dir_pattern = self.user_data_dir_pattern # 使用实例属性
            profile_num_patterns = self.profile_num_patterns   # 使用实例属性
            processed_pids_sync = set() # 避免重复处理同一进程

            for proc in psutil.process_iter(attrs=['pid', 'name', 'cmdline', 'exe'], ad_value=None):
                try:
                    # 在每次迭代开始时初始化这些变量
                    actual_user_data_dir_sync = None
                    extracted_profile_num_sync = None

                    if proc.pid in processed_pids_sync:
                        continue
                    
                    proc_info = proc.info
                    if not proc_info or \
                       not proc_info.get('name') or \
                       'chrome.exe' not in proc_info['name'].lower() or \
                       not proc_info.get('cmdline'):
                        continue
                    
                    cmd_line_list = proc_info['cmdline']
                    cmd_line_str = " ".join(cmd_line_list)
                    
                    # 尝试从 user-data-dir 路径中提取编号 (更可靠)
                    # actual_user_data_dir_sync # 已在循环顶部初始化为None
                    # extracted_profile_num_sync # 已在循环顶部初始化为None
                    match_ud_dir_sync = user_data_dir_pattern.search(cmd_line_str)
                    if match_ud_dir_sync:
                        print(f"DEBUG_SYNC_UDD_PARSE: PID={proc.pid}, user_data_dir_pattern matched. Groups: {match_ud_dir_sync.groups()}")
                        # 提取路径，优先尝试带引号的路径组，然后尝试不带引号的路径组
                        raw_udd_sync = match_ud_dir_sync.group('path') or match_ud_dir_sync.group('path_unquoted')
                        
                        if raw_udd_sync:
                            print(f"DEBUG_SYNC_UDD_PARSE: PID={proc.pid}, Raw UDD extracted: '{raw_udd_sync}', type: {type(raw_udd_sync)}")
                            # 不需要再手动去除引号，因为 (?P<path>[^"]+) 组已不包含引号
                            actual_user_data_dir_sync = os.path.normpath(raw_udd_sync.strip())
                            print(f"DEBUG_SYNC_UDD_PARSE: PID={proc.pid}, Normalized actual_user_data_dir: '{actual_user_data_dir_sync}', type: {type(actual_user_data_dir_sync)}")
                            
                            if actual_user_data_dir_sync: # Ensure actual_user_data_dir_sync is not None or empty
                                path_basename_sync = os.path.basename(actual_user_data_dir_sync)
                                # 尝试从当前目录名提取，例如 "chrome123" 或 "profile123"
                                match_name_num_sync = re.search(r"^(?:chrome|profile)(\\d+)$", path_basename_sync, re.IGNORECASE)
                                if match_name_num_sync:
                                    extracted_profile_num_sync = int(match_name_num_sync.group(1))
                                    print(f"DEBUG_SYNC_UDD_EXTRACT: PID={proc.pid}, Extracted {extracted_profile_num_sync} from basename '{path_basename_sync}' using (chrome|profile)N pattern.")
                                else:
                                    # 如果目录名本身就是数字 (例如，父目录是 "chrome_profiles", 子目录是 "1", "2", "3")
                                    if path_basename_sync.isdigit():
                                        parent_dir_name_sync = os.path.basename(os.path.dirname(actual_user_data_dir_sync)).lower()
                                        if "chrome" in parent_dir_name_sync or "profile" in parent_dir_name_sync:
                                             extracted_profile_num_sync = int(path_basename_sync)
                                             print(f"DEBUG_SYNC_UDD_EXTRACT: PID={proc.pid}, Extracted {extracted_profile_num_sync} from numeric basename '{path_basename_sync}' with parent '{parent_dir_name_sync}'")

                                    # 如果还没找到，尝试从父目录名提取，例如路径是 ".../chrome123/Default"
                                    if extracted_profile_num_sync is None:
                                        parent_basename_sync = os.path.basename(os.path.dirname(actual_user_data_dir_sync))
                                        match_parent_name_num_sync = re.search(r"^(?:chrome|profile)(\\d+)$", parent_basename_sync, re.IGNORECASE)
                                        if match_parent_name_num_sync:
                                             extracted_profile_num_sync = int(match_parent_name_num_sync.group(1))
                                             print(f"DEBUG_SYNC_UDD_EXTRACT: PID={proc.pid}, Extracted {extracted_profile_num_sync} from parent basename '{parent_basename_sync}'")
                        # No explicit else here, if raw_udd_sync is None, actual_user_data_dir_sync remains None
                    # No specific error caught here for match_ud_dir_sync being None or regex failing early for user-data-dir

                    # If extracted_profile_num_sync is still None after UDD parsing, try other patterns
                    if extracted_profile_num_sync is None:
                        for p_pattern_sync in profile_num_patterns:
                            match_profile_sync = p_pattern_sync.search(cmd_line_str)
                            if match_profile_sync:
                                try:
                                    num_str_sync = match_profile_sync.group(1)
                                    if p_pattern_sync.pattern == r"--profile-directory=(?:\\\"?Profile\\s+(\\\d+)\\\"?|Default)":
                                        print(f"DEBUG_SYNC_PROFILE_DIR_MATCH: PID={proc.pid}, Pattern='{p_pattern_sync.pattern}', FullMatch='{match_profile_sync.group(0)}', Group1='{num_str_sync}', AllGroups={match_profile_sync.groups()}")
                                    if num_str_sync:
                                        extracted_profile_num_sync = int(num_str_sync)
                                        print(f"DEBUG_SYNC_PROFILE_ASSIGNED: PID={proc.pid}, extracted_profile_num_sync assigned: {extracted_profile_num_sync}")
                                        break
                                    elif "Default" in match_profile_sync.group(0) and "Profile" not in match_profile_sync.group(0) and p_pattern_sync.pattern == r"--profile-directory=(?:\\\"?Profile\\s+(\\\d+)\\\"?|Default)":
                                        print(f"DEBUG_SYNC_PROFILE_DIR_MATCH: PID={proc.pid}, Matched Default profile.")
                                        pass # Default usually not counted or handled as a numbered profile here
                                except (IndexError, ValueError, AttributeError) as e_parse_group: # Added AttributeError
                                    print(f"DEBUG_SYNC_PROFILE_DIR_MATCH_ERROR: PID={proc.pid}, Pattern='{p_pattern_sync.pattern}', Error: {e_parse_group} for match '{match_profile_sync.group(0) if match_profile_sync else 'NO MATCH'}'")
                                    continue # Error in group extraction or int conversion

                    if extracted_profile_num_sync is not None:
                        actually_running_profiles.add(extracted_profile_num_sync)
                    # No "NOT ADDED" print here in sync function, only in open_url
                    
                    processed_pids_sync.add(proc.pid)

                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
                except (TypeError, ValueError, AttributeError) as e_data_handling_sync: # Added AttributeError here too
                    print(f"DEBUG_SYNC_DATA_HANDLING_ERROR: PID (approx {proc.pid if proc and hasattr(proc, 'pid') else 'N/A'}): {str(e_data_handling_sync)}")
                    continue
                except Exception as e_general_proc_sync:
                    print(f"DEBUG_SYNC_GENERAL_ERROR: PID (approx {proc.pid if proc and hasattr(proc, 'pid') else 'N/A'}): {str(e_general_proc_sync)}")
                    # traceback.print_exc() # Optionally re-enable for very deep debugging
                    continue
            
            self.launched_numbers = actually_running_profiles
            print(f"DEBUG: Synchronized. self.launched_numbers is now: {self.launched_numbers}")
            # 保存同步后的状态
            self.save_settings()

        except Exception as e:
            print(f"错误: 同步已运行分身状态失败: {str(e)}")
            import traceback
            traceback.print_exc()

    def create_compact_icon_management_section(self, main_layout):
        """创建精简的图标管理区域
        
        Args:
            main_layout: 主布局，用于添加图标管理区域
        """
        icon_group = self.create_group_box("图标管理")
        icon_layout = QVBoxLayout(icon_group)
        icon_layout.setContentsMargins(10, 20, 10, 10)
        icon_layout.setSpacing(8)
        
        # 自动图标应用设置
        auto_icon_layout = QHBoxLayout()
        self.auto_apply_checkbox = QCheckBox("启动Chrome时自动应用编号图标")
        self.auto_apply_checkbox.setChecked(self.auto_apply_icons)
        self.auto_apply_checkbox.stateChanged.connect(self.on_auto_apply_icons_changed)
        self.auto_apply_checkbox.setFont(QFont("Microsoft YaHei", 7))
        auto_icon_layout.addWidget(self.auto_apply_checkbox)
        auto_icon_layout.addStretch()
        icon_layout.addLayout(auto_icon_layout)
        
        # 图标管理按钮 - 紧凑布局，2行3列
        buttons_widget = QWidget()
        buttons_layout = QGridLayout(buttons_widget)
        buttons_layout.setSpacing(6)  # 减少按钮间距
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        
        # 按钮数据：(文本, 回调函数)
        button_data = [
            ("生成图标", self.generate_all_icons),
            ("应用图标", self.apply_icons_to_windows), 
            ("更新快捷方式", self.update_shortcut_icons),
            ("恢复默认", self.restore_default_icons),
            ("清理缓存", self.clean_icon_cache),
            ("缓存信息", self.show_cache_info)
        ]
        
        # 将按钮排列为2行3列，使用更紧凑的尺寸
        for i, (text, callback) in enumerate(button_data):
            row = i // 3
            col = i % 3
            button = self.create_compact_icon_button(text, callback)
            buttons_layout.addWidget(button, row, col)
        
        icon_layout.addWidget(buttons_widget)
        main_layout.addWidget(icon_group)
    
    def create_compact_icon_button(self, text, callback):
        """创建紧凑的图标管理按钮
        
        Args:
            text: 按钮文本
            callback: 按钮点击回调函数
            
        Returns:
            QPushButton: 样式化的紧凑图标管理按钮
        """
        button = QPushButton(text)
        button.clicked.connect(callback)
        button.setFixedHeight(26)  # 更小的按钮高度
        button.setFixedWidth(120)  # 更小的按钮宽度
        button.setFont(QFont("Microsoft YaHei", 7))  # 更小的字体
        return button
    
    # === 图标管理回调方法 ===
    
    def on_auto_apply_icons_changed(self, state):
        """自动应用图标选项状态改变时的回调"""
        self.auto_apply_icons = state == Qt.Checked
        self.settings.setValue("auto_apply_icons", self.auto_apply_icons)
        status_text = "已启用自动应用图标" if self.auto_apply_icons else "已禁用自动应用图标"
        self.set_status(status_text, "green")
    
    def generate_all_icons(self):
        """生成所有图标"""
        if not self.icon_manager:
            QMessageBox.warning(self, "警告", "图标管理功能不可用！")
            return
        
        try:
            # 获取范围
            start = int(self.start_num.text())
            end = int(self.end_num.text())
            numbers = list(range(start, end + 1))
            
            self.set_status("开始生成图标...", "blue")
            self.progress_bar.setVisible(True)
            
            self.icon_worker = IconManagementWorker("generate_icons", numbers)
            self.icon_worker.update_status.connect(lambda msg, color: self.set_status(msg, color))
            self.icon_worker.update_progress.connect(self.progress_bar.setValue)
            self.icon_worker.finished.connect(self.on_icon_operation_finished)
            self.icon_worker.start()
            
        except ValueError:
            QMessageBox.warning(self, "错误", "请输入有效的数字范围！")
    
    def apply_icons_to_windows(self):
        """手动应用图标到当前窗口"""
        if not self.icon_manager:
            QMessageBox.warning(self, "警告", "图标管理功能不可用！")
            return
        
        if not self.launched_numbers:
            QMessageBox.information(self, "提示", "当前没有已启动的Chrome分身！")
            return
        
        self.set_status("开始应用图标到窗口...", "blue")
        self.progress_bar.setVisible(True)
        
        self.icon_worker = IconManagementWorker("apply_icons", list(self.launched_numbers))
        self.icon_worker.update_status.connect(lambda msg, color: self.set_status(msg, color))
        self.icon_worker.update_progress.connect(self.progress_bar.setValue)
        self.icon_worker.finished.connect(self.on_icon_operation_finished)
        self.icon_worker.start()
    
    def update_shortcut_icons(self):
        """更新快捷方式图标"""
        if not self.icon_manager:
            QMessageBox.warning(self, "警告", "图标管理功能不可用！")
            return
        
        shortcut_dir = self.folder_path.text().strip()
        if not shortcut_dir or not os.path.exists(shortcut_dir):
            QMessageBox.warning(self, "警告", "请设置有效的快捷方式目录！")
            return
        
        if not self.launched_numbers:
            # 如果没有已启动的分身，使用范围设置
            try:
                start = int(self.start_num.text())
                end = int(self.end_num.text())
                numbers = list(range(start, end + 1))
            except ValueError:
                QMessageBox.warning(self, "错误", "请输入有效的数字范围！")
                return
        else:
            numbers = list(self.launched_numbers)
        
        self.set_status("开始更新快捷方式图标...", "blue")
        self.progress_bar.setVisible(True)
        
        self.icon_worker = IconManagementWorker("update_shortcuts", numbers, shortcut_dir)
        self.icon_worker.update_status.connect(lambda msg, color: self.set_status(msg, color))
        self.icon_worker.update_progress.connect(self.progress_bar.setValue)
        self.icon_worker.finished.connect(self.on_icon_operation_finished)
        self.icon_worker.start()
    
    def restore_default_icons(self):
        """恢复默认图标"""
        if not self.icon_manager:
            QMessageBox.warning(self, "警告", "图标管理功能不可用！")
            return
        
        reply = QMessageBox.question(self, '确认操作', 
                                    '确定要恢复所有快捷方式的默认图标吗？',
                                    QMessageBox.Yes | QMessageBox.No, 
                                    QMessageBox.No)
        
        if reply != QMessageBox.Yes:
            return
        
        shortcut_dir = self.folder_path.text().strip()
        if not shortcut_dir or not os.path.exists(shortcut_dir):
            QMessageBox.warning(self, "警告", "请设置有效的快捷方式目录！")
            return
        
        self.set_status("开始恢复默认图标...", "blue")
        self.progress_bar.setVisible(True)
        
        self.icon_worker = IconManagementWorker("restore_defaults", None, shortcut_dir)
        self.icon_worker.update_status.connect(lambda msg, color: self.set_status(msg, color))
        self.icon_worker.update_progress.connect(self.progress_bar.setValue)
        self.icon_worker.finished.connect(self.on_icon_operation_finished)
        self.icon_worker.start()
    
    def clean_icon_cache(self):
        """清理系统图标缓存"""
        if not self.icon_manager:
            QMessageBox.warning(self, "警告", "图标管理功能不可用！")
            return
        
        reply = QMessageBox.question(self, '确认清理图标缓存', 
                                    '清理图标缓存将重启资源管理器，是否继续？\n\n'
                                    '注意：此操作会短暂关闭所有资源管理器窗口。',
                                    QMessageBox.Yes | QMessageBox.No, 
                                    QMessageBox.No)
        
        if reply != QMessageBox.Yes:
            return
        
        self.set_status("开始清理系统图标缓存...", "blue")
        self.progress_bar.setVisible(True)
        
        self.icon_worker = IconManagementWorker("clean_cache", None)
        self.icon_worker.update_status.connect(lambda msg, color: self.set_status(msg, color))
        self.icon_worker.update_progress.connect(self.progress_bar.setValue)
        self.icon_worker.finished.connect(self.on_icon_operation_finished)
        self.icon_worker.start()
    
    def show_cache_info(self):
        """显示图标缓存信息"""
        if not self.icon_manager:
            QMessageBox.warning(self, "警告", "图标管理功能不可用！")
            return
        
        try:
            # 获取缓存目录信息
            cache_dir = self.icon_manager.icon_cache_dir
            if not os.path.exists(cache_dir):
                QMessageBox.information(self, "缓存信息", "图标缓存目录不存在。")
                return
            
            # 统计图标文件
            icon_files = []
            cache_size = 0
            
            for file in os.listdir(cache_dir):
                if file.endswith(('.ico', '.png')):
                    file_path = os.path.join(cache_dir, file)
                    icon_files.append(file)
                    try:
                        cache_size += os.path.getsize(file_path)
                    except:
                        pass
            
            cache_size_mb = cache_size / (1024 * 1024)
            
            info_text = f"""图标缓存信息:

缓存目录: {cache_dir}
生成的图标数量: {len(icon_files)}
缓存大小: {cache_size_mb:.2f} MB

生成的图标文件:
{chr(10).join(icon_files[:15])}
{'...' if len(icon_files) > 15 else ''}
            """
            
            QMessageBox.information(self, "图标缓存信息", info_text)
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"获取缓存信息失败: {str(e)}")
    
    def on_icon_operation_finished(self, message, color):
        """图标操作完成回调"""
        self.progress_bar.setVisible(False)
        
        # 手动图标操作的结果只在状态栏显示，不覆盖主要日志
        if "成功" in message:
            self.statusBar.showMessage(f"✅ {message}")
        elif "失败" in message:
            self.statusBar.showMessage(f"❌ {message}")
        else:
            self.statusBar.showMessage(f"🔧 {message}")
        
        # 显示完成提示
        if "成功" in message:
            QMessageBox.information(self, "操作完成", message)
        elif "失败" in message:
            QMessageBox.warning(self, "操作失败", message)
    
    def apply_icons_for_numbers(self, numbers):
        """为指定的分身编号应用图标（在启动完成后自动调用）"""
        if not self.icon_manager or not self.auto_apply_icons:
            return
        
        if not numbers:
            return
        
        # 图标应用进度只在状态栏显示，不覆盖主要日志
        self.statusBar.showMessage(f"🎨 正在为 {len(numbers)} 个分身自动应用图标...")
        
        # 延迟一下，确保Chrome窗口已经创建
        def delayed_apply():
            self.icon_worker = IconManagementWorker("apply_icons", numbers)
            # 图标应用的过程信息只在状态栏显示
            self.icon_worker.update_status.connect(lambda msg, color: self.statusBar.showMessage(f"🎨 {msg}"))
            self.icon_worker.finished.connect(lambda msg, color: self.statusBar.showMessage(f"🎨 自动图标应用完成: {msg}"))
            self.icon_worker.start()
        
        # 等待2秒让Chrome窗口完全加载
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(2000, delayed_apply)
    
    def create_compact_button(self, text, callback, layout):
        """创建紧凑的按钮
        
        Args:
            text: 按钮文本
            callback: 按钮点击回调函数
            layout: 要添加到的布局
        """
        button = QPushButton(text)
        button.clicked.connect(callback)
        button.setFixedHeight(28)
        button.setMinimumWidth(85)
        button.setFont(QFont("Microsoft YaHei", 8))
        layout.addWidget(button)
        return button
    
    # === 创建分身相关方法 ===
    
    def browse_shortcut_path(self):
        """浏览选择快捷方式保存路径"""
        try:
            folder_path = QFileDialog.getExistingDirectory(
                self, 
                "选择快捷方式保存目录",
                self.shortcut_path_entry.text() or os.path.expanduser("~/Desktop")
            )
            if folder_path:
                self.shortcut_path_entry.setText(folder_path)
                # 路径选择信息只在状态栏显示，不覆盖主要日志
                self.statusBar.showMessage(f"📁 已选择快捷方式路径: {os.path.basename(folder_path)}")
        except Exception as e:
            # 错误信息可以覆盖主要日志
            self.set_status(f"选择快捷方式路径失败: {str(e)}", "red")
    
    def browse_cache_path(self):
        """浏览选择缓存储存路径"""
        try:
            folder_path = QFileDialog.getExistingDirectory(
                self, 
                "选择缓存数据保存目录",
                self.cache_path_entry.text() or os.path.expanduser("~/AppData/Local/Google/Chrome")
            )
            if folder_path:
                self.cache_path_entry.setText(folder_path)
                # 路径选择信息只在状态栏显示，不覆盖主要日志
                self.statusBar.showMessage(f"📁 已选择缓存路径: {os.path.basename(folder_path)}")
        except Exception as e:
            # 错误信息可以覆盖主要日志
            self.set_status(f"选择缓存路径失败: {str(e)}", "red")
    
    def create_chrome_profiles(self):
        """创建Chrome分身"""
        try:
            # 获取输入参数
            shortcut_path = self.shortcut_path_entry.text().strip()
            cache_path = self.cache_path_entry.text().strip()
            start_text = self.create_start_entry.text().strip()
            end_text = self.create_end_entry.text().strip()
            
            # 验证输入参数
            if not shortcut_path:
                QMessageBox.warning(self, "警告", "请选择快捷方式保存路径！")
                return
                
            if not cache_path:
                QMessageBox.warning(self, "警告", "请选择缓存储存路径！")
                return
                
            if not start_text or not end_text:
                QMessageBox.warning(self, "警告", "请输入创建编号范围！")
                return
            
            try:
                start = int(start_text)
                end = int(end_text)
                if start < 1 or end < 1 or start > end:
                    QMessageBox.warning(self, "警告", "创建编号范围无效，请输入有效的编号范围！")
                    return
            except ValueError:
                QMessageBox.warning(self, "警告", "请输入有效的数字！")
                return
            
            # 检查路径是否存在
            if not os.path.exists(shortcut_path):
                try:
                    os.makedirs(shortcut_path, exist_ok=True)
                except Exception as e:
                    QMessageBox.warning(self, "错误", f"创建快捷方式目录失败: {str(e)}")
                    return
            
            if not os.path.exists(cache_path):
                try:
                    os.makedirs(cache_path, exist_ok=True)
                except Exception as e:
                    QMessageBox.warning(self, "错误", f"创建缓存目录失败: {str(e)}")
                    return
            
            # 启动创建任务
            self.set_status("开始创建Chrome分身...", "blue")
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # 创建工作线程
            self.profile_worker = ProfileCreationWorker(shortcut_path, cache_path, start, end)
            self.profile_worker.update_status.connect(lambda msg, color: self.set_status(msg, color))
            self.profile_worker.update_progress.connect(self.progress_bar.setValue)
            self.profile_worker.finished.connect(self.on_profile_creation_finished)
            self.profile_worker.start()
            
        except Exception as e:
            self.set_status(f"创建分身失败: {str(e)}", "red")
            QMessageBox.critical(self, "错误", f"创建分身时发生错误: {str(e)}")
    
    def on_profile_creation_finished(self, status_text, color, created_count):
        """分身创建完成回调"""
        self.set_status(status_text, color)
        self.statusBar.showMessage(f"分身创建完成。创建数量: {created_count}")
        
        # 重新启用按钮
        self.setEnabled(True)


if __name__ == "__main__":
    import sys
    
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion样式获得更好的外观
    
    # 创建主窗口
    launcher = ChromeLauncher()
    launcher.show()
    
    sys.exit(app.exec_())