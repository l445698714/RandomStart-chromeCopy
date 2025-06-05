import os
import subprocess
import threading
import time
import win32gui
import win32process
import win32con
import win32api
import win32com.client
import re
import ctypes
import math
import gc
import psutil
from ctypes import wintypes
from typing import List, Dict, Tuple, Optional, Any
import pythoncom
from hooks import InputHookManager
from utils import (
    log_error,
    load_settings,
    save_settings,
    parse_window_numbers,
    generate_color_icon,
    center_window,
    get_chrome_popups,
    title_similarity,
    normalize_path,
    show_notification
)
from config import ICON_DIR
import random

# Regex for parsing --user-data-dir, adopted from Chrome_launcher.py for robustness
USER_DATA_DIR_PATTERN = re.compile(r'--user-data-dir=(?:\"(?P<path>[^"]+)\"|(?P<path_unquoted>[^\s]+(?:\s+[^\s]+)*?(?=\s*--|\s*$)))')
# Regex for parsing --remote-debugging-port
REMOTE_DEBUGGING_PORT_PATTERN = re.compile(r'--remote-debugging-port=(\d+)')

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

class ChromeManager:
    def __init__(self, ui_manager=None):
        self.ui_manager = ui_manager
        
        self.DWMWA_BORDER_COLOR = 34
        
        self.start_time = time.time()
        
        self.settings = load_settings()
        self.shortcut_path = self.settings.get("shortcut_path", "")
        self.cache_dir = self.settings.get("cache_dir", "")
        self.screen_selection = self.settings.get("screen_selection", "")
        self.auto_modify_shortcut_icon = self.settings.get("auto_modify_shortcut_icon", True)
        
        try:
            pythoncom.CoInitialize()
            self.shell = win32com.client.Dispatch("WScript.Shell")
        except Exception as e:
            log_error("初始化COM组件失败", e)
            self.shell = None
        
        self.hook_manager = InputHookManager()
        
        self.windows = {}
        self.master_window = None
        self.is_syncing = False
        
        self.debug_ports = {}
        
        self.shortcut_to_pid = {}
        self.pid_to_number = {}
        
        self.screens = []
        self.screen_names = []
        
        self.temp_files = []
        self.last_activity_time = time.time()
        self.icon_cache = {}
        
        self.memory_monitor_active = True
        self.memory_monitor_thread = threading.Thread(target=self.monitor_memory, daemon=True)
        self.memory_monitor_thread.start()
        
        self.update_screen_info()
    
    def monitor_memory(self):
        try:
            process = psutil.Process(os.getpid())
            light_cleanup_interval = 300
            deep_cleanup_interval = 3600
            last_light_cleanup = time.time()
            last_deep_cleanup = time.time()
            
            while self.memory_monitor_active:
                try:
                    current_time = time.time()
                    mem_usage = process.memory_info().rss / (1024 * 1024)
                    
                    user_idle_time = current_time - self.last_activity_time
                    is_user_idle = user_idle_time > 30
                    
                    if current_time - last_light_cleanup > light_cleanup_interval:
                        self.perform_light_cleanup()
                        last_light_cleanup = current_time
                    
                    if mem_usage > 300:
                        if is_user_idle:
                            self.perform_medium_cleanup()
                    
                    if current_time - last_deep_cleanup > deep_cleanup_interval and is_user_idle and user_idle_time > 120:
                        self.perform_deep_cleanup()
                        last_deep_cleanup = current_time
                        
                    time.sleep(10)
                    
                except Exception as e:
                    log_error("内存监控循环异常", e)
                    time.sleep(60)
        except Exception as e:
            log_error("内存监控线程初始化异常", e)
    
    def perform_light_cleanup(self):
        try:
            if hasattr(self, "debug_ports"):
                active_ports = {}
                for num, port in self.debug_ports.items():
                    is_active = False
                    for window in self.windows.values():
                        if window.get("number") == num:
                            is_active = True
                            break
                    if is_active:
                        active_ports[num] = port
                self.debug_ports = active_ports
            
            gc.collect(0)
            
        except Exception as e:
            log_error("轻量级清理异常", e)
    
    def perform_medium_cleanup(self):
        try:
            gc.collect()
            
            self.clean_temp_files()
            
            if hasattr(self, "icon_cache") and self.icon_cache:
                active_numbers = set()
                for window in self.windows.values():
                    if "number" in window:
                        active_numbers.add(window["number"])
                
                keys_to_remove = [k for k in self.icon_cache.keys() if k not in active_numbers]
                for key in keys_to_remove:
                    del self.icon_cache[key]
            
        except Exception as e:
            log_error("中等强度清理异常", e)
    
    def perform_deep_cleanup(self):
        try:
            gc.collect(2)
            
            if hasattr(self, 'icon_cache'):
                self.icon_cache.clear()
            
            if self.is_syncing and time.time() - self.last_activity_time > 300:
                self.temporarily_release_resources()
            
            self.organize_icon_cache()
            
        except Exception as e:
            log_error("深度清理异常", e)
    
    def update_activity_timestamp(self):
        self.last_activity_time = time.time()
    
    def clean_temp_files(self):
        try:
            removed_count = 0
            for temp_file in list(self.temp_files):
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                        self.temp_files.remove(temp_file)
                        removed_count += 1
                except Exception as e:
                    log_error(f"删除临时文件失败: {temp_file}", e)
            
        except Exception as e:
            log_error("清理临时文件异常", e)
    
    def organize_icon_cache(self):
        try:
            if not self.cache_dir or not os.path.exists(self.cache_dir):
                return
            
            active_numbers = set()
            for window in self.windows.values():
                if "number" in window:
                    active_numbers.add(str(window["number"]))
            
            files_cleaned = 0
            total_size_cleaned = 0
            
            for filename in os.listdir(self.cache_dir):
                if filename.endswith(".ico"):
                    try:
                        number = filename.split(".")[0]
                        if number.isdigit() and number not in active_numbers:
                            file_path = os.path.join(self.cache_dir, filename)
                            file_size = os.path.getsize(file_path)
                            
                            mtime = os.path.getmtime(file_path)
                            if time.time() - mtime > 24 * 3600:
                                os.remove(file_path)
                                files_cleaned += 1
                                total_size_cleaned += file_size
                    except Exception:
                        pass
            
        except Exception as e:
            log_error("整理图标缓存异常", e)
    
    def temporarily_release_resources(self):
        try:
            if hasattr(self, 'hook_manager') and self.is_syncing:
                self.hook_manager.release_unused_hooks()
            
            return True
        except Exception as e:
            log_error("释放资源异常", e)
            return False
    
    def cleanup_on_exit(self):
        try:
            self.memory_monitor_active = False
            
            if hasattr(self, 'memory_monitor_thread') and self.memory_monitor_thread.is_alive():
                self.memory_monitor_thread.join(timeout=0.5)
                
            self.clean_temp_files()
            
            gc.collect()
            
        except Exception as e:
            log_error("退出清理异常", e)
    
    def open_windows(self, numbers_str: str) -> bool:
        self.update_activity_timestamp()
        
        shortcut_dir = self.shortcut_path
        
        if not shortcut_dir or not os.path.exists(shortcut_dir):
            log_error(f"快捷方式目录不存在: {shortcut_dir}")
            return False
        
        try:
            window_numbers = parse_window_numbers(numbers_str)
        except Exception as e:
            log_error("解析窗口编号失败", e)
            return False
        
        self.debug_ports.clear()
        
        temp_files = []
        
        try:
            for num in window_numbers:
                shortcut = os.path.join(shortcut_dir, f"{num}.lnk")
                if not os.path.exists(shortcut):
                    log_error(f"快捷方式不存在: {shortcut}")
                    continue
                
                shortcut_obj = self.shell.CreateShortCut(shortcut)
                target = shortcut_obj.TargetPath
                args = shortcut_obj.Arguments
                working_dir = shortcut_obj.WorkingDirectory
                
                debug_port = 9222 + int(num)
                
                self.debug_ports[num] = debug_port
                
                if "--remote-debugging-port=" in args:
                    new_args = re.sub(
                        r"--remote-debugging-port=\d+",
                        f"--remote-debugging-port={debug_port}",
                        args,
                    )
                else:
                    new_args = f"{args} --remote-debugging-port={debug_port}"
                
                temp_shortcut = os.path.join(shortcut_dir, f"temp_{num}.lnk")
                temp_obj = self.shell.CreateShortCut(temp_shortcut)
                temp_obj.TargetPath = target
                temp_obj.Arguments = new_args
                temp_obj.WorkingDirectory = working_dir
                temp_obj.IconLocation = shortcut_obj.IconLocation
                temp_obj.Save()
                
                temp_files.append(temp_shortcut)
                self.temp_files.append(temp_shortcut)
                
                if os.path.exists(temp_shortcut):
                    try:
                        subprocess.Popen(["start", "", temp_shortcut], shell=True)
                        time.sleep(0.1)
                    except Exception as e:
                        log_error(f"启动窗口 {num} 失败", e)
                else:
                    try:
                        subprocess.Popen(["start", "", shortcut], shell=True)
                        time.sleep(0.1)
                    except Exception as e:
                        log_error(f"启动窗口 {num} 失败", e)
            
            def cleanup_temp_files():
                time.sleep(2)
                for temp_file in temp_files:
                    try:
                        if os.path.exists(temp_file):
                            os.remove(temp_file)
                    except Exception as e:
                        log_error(f"删除临时文件失败: {temp_file}", e)
                
                if hasattr(self, "perform_light_cleanup"):
                    self.perform_light_cleanup()
            
            threading.Thread(target=cleanup_temp_files).start()
            
            self.settings["last_window_numbers"] = numbers_str
            save_settings(self.settings)
            
            return True
        
        except Exception as e:
            log_error("打开窗口失败", e)
            return False
    
    def import_windows(self, update_ui=True) -> List[Dict]:
        if hasattr(self, 'logger'): 
            self.logger.info("进入 import_windows 方法")
        else:
            print("INFO: 进入 import_windows 方法")

        try:
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("WScript.Shell")
            if hasattr(self, 'logger'): self.logger.debug("WScript.Shell COM对象初始化成功")
            else: print("DEBUG: WScript.Shell COM对象初始化成功")
        except Exception as e:
            log_error("导入窗口时初始化COM组件失败", e)
            shell = None
            if hasattr(self, 'logger'): self.logger.error("WScript.Shell COM对象初始化失败")
            else: print("ERROR: WScript.Shell COM对象初始化失败")

        profile_candidate_pids = {}
        processed_proc_count = 0

        if hasattr(self, 'logger'): self.logger.debug("准备开始遍历 psutil.process_iter")
        else: print("DEBUG: 准备开始遍历 psutil.process_iter")

        try:
            for proc in psutil.process_iter(['pid', 'name', 'cmdline', 'exe']):
                processed_proc_count += 1
                try:
                    proc_info = proc.info
                    if not proc_info['name'] or 'chrome.exe' not in proc_info['name'].lower() or not proc_info['cmdline']:
                        continue

                    pid = proc_info['pid']
                    cmdline_list = proc_info['cmdline'] 
                    cmdline_str = " ".join(cmdline_list) 
                    
                    actual_user_data_dir = None
                    profile_number = None

                    match_ud_dir = USER_DATA_DIR_PATTERN.search(cmdline_str)
                    if match_ud_dir:
                        raw_udd = match_ud_dir.group('path') or match_ud_dir.group('path_unquoted')
                        if raw_udd:
                            actual_user_data_dir = os.path.normpath(raw_udd.strip())
                    
                    if actual_user_data_dir:
                        udd_basename = os.path.basename(actual_user_data_dir)
                        if udd_basename.isdigit():
                            profile_number = int(udd_basename)

                    if profile_number is None and actual_user_data_dir and self.shortcut_path and os.path.exists(self.shortcut_path) and shell:
                        try:
                            for f_name in os.listdir(self.shortcut_path):
                                if f_name.lower().endswith(".lnk"):
                                    num_part = f_name[:-4]
                                    if num_part.isdigit():
                                        potential_num = int(num_part)
                                        shortcut_file_path = os.path.join(self.shortcut_path, f_name)
                                        try:
                                            shortcut_obj = shell.CreateShortCut(shortcut_file_path)
                                            shortcut_args_str = shortcut_obj.Arguments
                                            shortcut_match_ud_dir = USER_DATA_DIR_PATTERN.search(shortcut_args_str)
                                            if shortcut_match_ud_dir:
                                                raw_shortcut_udd = shortcut_match_ud_dir.group('path') or shortcut_match_ud_dir.group('path_unquoted')
                                                if raw_shortcut_udd:
                                                    normalized_shortcut_udd = os.path.normpath(raw_shortcut_udd.strip())
                                                    if normalized_shortcut_udd == actual_user_data_dir:
                                                        profile_number = potential_num
                                                        if hasattr(self, 'logger'):
                                                            self.logger.info(f"    策略4: PID {pid} shortcut '{f_name}' -> num {profile_number} (UDD: {actual_user_data_dir})")
                                                        break 
                                        except Exception: pass 
                        except Exception: pass
                    
                    if profile_number is not None:
                        is_likely_main_process = True
                        for arg in cmdline_list: 
                            if arg.startswith("--type="):
                                if arg.split("=")[1].lower() in ["renderer", "gpu-process", "utility", "crashpad-handler", "broker", "service", "extension", "ppapi", "plugin"]:
                                    is_likely_main_process = False
                                    break
                    
                        current_candidate = profile_candidate_pids.get(profile_number)
                        should_update_candidate = False

                        if current_candidate is None: 
                            should_update_candidate = True
                        elif is_likely_main_process and not current_candidate.get('is_likely_main', False):
                            should_update_candidate = True

                        if should_update_candidate:
                            profile_candidate_pids[profile_number] = {
                                'pid': pid,
                                'user_data_dir': actual_user_data_dir,
                                'is_likely_main': is_likely_main_process,
                            }
                            if hasattr(self, 'logger'):
                                self.logger.info(f"更新候选: PID {pid} (主: {is_likely_main_process}) 成为编号 {profile_number} 的候选。旧候选主: {current_candidate.get('is_likely_main') if current_candidate else 'N/A'}.")
                
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
                except AttributeError as e_attr: 
                    if hasattr(self, 'logger') and proc: 
                        self.logger.debug(f"AttributeError accessing proc.info for PID (likely already exited): {proc.pid if hasattr(proc, 'pid') else 'Unknown PID'}. Error: {e_attr}")
                    continue
                except Exception as e_inner_loop: 
                    if hasattr(self, 'logger'): self.logger.error(f"psutil循环内部异常 (PID: {pid if 'pid' in locals() else 'Unknown'}): {e_inner_loop}", exc_info=True)
                    else: print(f"ERROR: psutil循环内部异常 (PID: {pid if 'pid' in locals() else 'Unknown'}): {e_inner_loop}")
                    continue 

        except Exception as e_outer_psutil: 
            if hasattr(self, 'logger'): self.logger.critical(f"psutil.process_iter 遍历时发生严重错误: {e_outer_psutil}", exc_info=True)
            else: print(f"CRITICAL: psutil.process_iter 遍历时发生严重错误: {e_outer_psutil}")
            return []


        if hasattr(self, 'logger'): 
            self.logger.info(f"psutil.process_iter 遍历完成，共检查 {processed_proc_count} 个进程。")
            self.logger.info("所有候选PID处理完成。最终候选者列表 (profile_candidate_pids):")
            if not profile_candidate_pids: self.logger.warning("  profile_candidate_pids 为空!")
            for pn, cand_info in profile_candidate_pids.items():
                self.logger.info(f"  编号 {pn}: PID={cand_info['pid']}, 主进程标志={cand_info['is_likely_main']}")
        else:
            print(f"INFO: psutil.process_iter 遍历完成，共检查 {processed_proc_count} 个进程。")
            print("INFO: 所有候选PID处理完成。最终候选者列表 (profile_candidate_pids):")
            if not profile_candidate_pids: print("WARNING:   profile_candidate_pids 为空!")
            for pn, cand_info in profile_candidate_pids.items():
                print(f"  编号 {pn}: PID={cand_info['pid']}, 主进程标志={cand_info['is_likely_main']}")
        
        if hasattr(self, 'logger'): 
            self.logger.debug(f"准备进行 self.windows 清理和更新。当前 self.windows: {self.windows}")
            self.logger.debug(f"  profile_candidate_pids: {profile_candidate_pids}")
        else: 
            print(f"DEBUG: 准备进行 self.windows 清理和更新。当前 self.windows: {self.windows}")
            print(f"DEBUG:   profile_candidate_pids: {profile_candidate_pids}")

        try: 
            if hasattr(self, 'logger'): self.logger.debug("  尝试计算 current_valid_numbers...")
            else: print("DEBUG:   尝试计算 current_valid_numbers...")
            current_valid_numbers = set(profile_candidate_pids.keys() if profile_candidate_pids else []) 
            if hasattr(self, 'logger'): self.logger.debug(f"    current_valid_numbers: {current_valid_numbers}, type: {type(current_valid_numbers)}")
            else: print(f"DEBUG:     current_valid_numbers: {current_valid_numbers}, type: {type(current_valid_numbers)}")

            if hasattr(self, 'logger'): self.logger.debug("  尝试计算 numbers_to_remove_from_self...")
            else: print("DEBUG:   尝试计算 numbers_to_remove_from_self...")
            if not isinstance(self.windows, dict):
                 if hasattr(self, 'logger'): self.logger.error(f"  self.windows 不是一个字典! 类型是: {type(self.windows)}. 重置为空字典。")
                 else: print(f"ERROR:   self.windows 不是一个字典! 类型是: {type(self.windows)}. 重置为空字典。")
                 self.windows = {} 

            numbers_to_remove_from_self = set(self.windows.keys()) - current_valid_numbers
            if hasattr(self, 'logger'): self.logger.debug(f"    numbers_to_remove_from_self: {numbers_to_remove_from_self}, type: {type(numbers_to_remove_from_self)}")
            else: print(f"DEBUG:     numbers_to_remove_from_self: {numbers_to_remove_from_self}, type: {type(numbers_to_remove_from_self)}")
            
            if hasattr(self, 'logger'): self.logger.debug(f"  将从 self.windows 移除的编号: {numbers_to_remove_from_self}")
            else: print(f"DEBUG:   将从 self.windows 移除的编号: {numbers_to_remove_from_self}")

            for num_to_del in numbers_to_remove_from_self:
                if hasattr(self, 'logger'):
                    self.logger.info(f"清理: 编号 {num_to_del} 不再有候选PID，从 self.windows 中移除。")
                old_pid = self.windows.pop(num_to_del, {}).get('pid')
                if old_pid and old_pid in self.pid_to_number and self.pid_to_number.get(old_pid) == num_to_del: # Safely check pid_to_number
                    del self.pid_to_number[old_pid]
                self.debug_ports.pop(num_to_del, None)

            if hasattr(self, 'logger'): self.logger.debug("完成清理 self.windows。开始更新/添加条目。")
            else: print("DEBUG: 完成清理 self.windows。开始更新/添加条目。")

            for num, candidate_info in profile_candidate_pids.items(): 
                if not isinstance(candidate_info, dict): 
                    if hasattr(self, 'logger'): self.logger.warning(f"  候选信息 for num {num} 不是字典: {candidate_info}。跳过。")
                    else: print(f"WARNING:   候选信息 for num {num} 不是字典: {candidate_info}。跳过。")
                    continue

                pid = candidate_info.get('pid') 
                if pid is None:
                    if hasattr(self, 'logger'): self.logger.warning(f"  候选信息 for num {num} 缺少 PID: {candidate_info}。跳过。")
                    else: print(f"WARNING:   候选信息 for num {num} 缺少 PID: {candidate_info}。跳过。")
                    continue
                
                if not candidate_info.get('is_likely_main'):
                    if num in self.windows and self.windows.get(num, {}).get('pid') != pid : 
                         if hasattr(self, 'logger'):
                            self.logger.warning(f"编号 {num} 的候选PID {pid} 不是主进程，但之前记录的PID是 {self.windows.get(num, {}).get('pid')}。将尝试通过HWND重新关联。")
                    elif num not in self.windows :
                         if hasattr(self, 'logger'):
                            self.logger.debug(f"编号 {num} 的候选PID {pid} 不是主进程，且该编号无现有记录。跳过初步添加。")
                         continue
                
                expected_cdp_port = self.settings.get("BASE_DEBUG_PORT", 9222) + num
                user_data_dir_cand = candidate_info.get('user_data_dir') 

                if num not in self.windows:
                    self.windows[num] = {
                        'number': num,
                        'pid': pid,
                        'user_data_dir': user_data_dir_cand,
                        'debug_port': expected_cdp_port,
                        'status': 'identified_by_pid',
                        'hwnd_debug': None,
                        'title_debug': None,
                        'tabs': []
                    }
                    if hasattr(self, 'logger'):
                        self.logger.info(f"  新识别窗口(基于PID): 编号 {num}, PID {pid}, UDD: {user_data_dir_cand}")
                else:
                    if self.windows.get(num,{}).get('pid') != pid:
                        if hasattr(self, 'logger'):
                            self.logger.info(f"  更新窗口PID(基于PID): 编号 {num}, 旧PID {self.windows.get(num,{}).get('pid')}, 新PID {pid}")
                        self.windows[num]['pid'] = pid
                        self.windows[num]['user_data_dir'] = user_data_dir_cand
                        self.windows[num]['hwnd_debug'] = None
                        self.windows[num]['title_debug'] = None
                    self.windows[num]['debug_port'] = expected_cdp_port
                    self.windows[num]['status'] = 'reconfirmed_by_pid'

                self.debug_ports[num] = expected_cdp_port
                for p_iter, n_map_iter in list(self.pid_to_number.items()): 
                    if n_map_iter == num and p_iter != pid:
                        if hasattr(self, 'logger'): self.logger.debug(f"  清理旧 pid_to_number: del self.pid_to_number[{p_iter}] (原指向编号 {n_map_iter})")
                        del self.pid_to_number[p_iter]
                self.pid_to_number[pid] = num

            if hasattr(self, 'logger'): self.logger.debug("完成更新/添加 self.windows 条目。开始最终清理 stale PIDs。")
            else: print("DEBUG: 完成更新/添加 self.windows 条目。开始最终清理 stale PIDs。")

            final_stale_numbers = []
            for num_key, win_info_val in list(self.windows.items()): 
                pid_val = win_info_val.get('pid')
                if pid_val is not None and not psutil.pid_exists(pid_val):
                    final_stale_numbers.append(num_key)
            
            if hasattr(self, 'logger'): self.logger.debug(f"  将进行最终清理的编号 (PID不存在): {final_stale_numbers}")
            else: print(f"DEBUG:   将进行最终清理的编号 (PID不存在): {final_stale_numbers}")

            for num_to_remove in final_stale_numbers:
                if hasattr(self, 'logger'):
                    self.logger.info(f"  最终清理: PID {self.windows.get(num_to_remove,{}).get('pid')} (编号 {num_to_remove}) 不存在，移除。")
                pid_to_remove_val = self.windows.pop(num_to_remove, {}).get('pid')
                if pid_to_remove_val and pid_to_remove_val in self.pid_to_number:
                    if self.pid_to_number.get(pid_to_remove_val) == num_to_remove: 
                        del self.pid_to_number[pid_to_remove_val]
                if num_to_remove in self.debug_ports:
                    del self.debug_ports[num_to_remove]
        
        except Exception as e_self_windows_update: 
            if hasattr(self, 'logger'):
                self.logger.critical(f"更新 self.windows 或清理PID时发生严重错误: {e_self_windows_update}", exc_info=True)
            else:
                print(f"CRITICAL: 更新 self.windows 或清理PID时发生严重错误: {e_self_windows_update}")
            return [] 

        if hasattr(self, 'logger'):
            self.logger.info(f"PID识别和窗口列表初步构建完成。当前 self.windows 数量: {len(self.windows)}")
            if not self.windows: self.logger.warning("  self.windows 为空!")
            for n,w in self.windows.items(): # 注意这里，确保它在 if hasattr(self, 'logger') 内部
                self.logger.debug(f"  初步列表: 编号 {n}, PID {w.get('pid')}, UDD {w.get('user_data_dir')}, Status {w.get('status')}")
        else: # 这个 else 块需要正确对齐
            print(f"INFO: PID识别和窗口列表初步构建完成。当前 self.windows 数量: {len(self.windows)}")
            if not self.windows: print("WARNING:   self.windows 为空!")
            for n,w in self.windows.items():
                print(f"  初步列表: 编号 {n}, PID {w.get('pid')}, UDD {w.get('user_data_dir')}, Status {w.get('status')}")


        if hasattr(self, 'logger'): self.logger.debug("准备进入 EnumWindows 调用")
        else: print("DEBUG: 准备进入 EnumWindows 调用")

        number_to_hwnd_details = {} 

        def enum_windows_callback(hwnd, callback_data_map):
            try:
                if not (win32gui.IsWindowVisible(hwnd) and win32gui.GetParent(hwnd) == 0):
                    return True

                _, found_pid_for_hwnd = win32process.GetWindowThreadProcessId(hwnd)
                
                if not self.windows:
                     return True

                # 使用 list(self.windows.items()) 来迭代一个副本，以防在迭代过程中 self.windows 被修改（理论上不应在此回调中发生）
                for num_in_loop, window_entry in list(self.windows.items()): 
                    target_pid_from_psutil = window_entry.get('pid')

                    if target_pid_from_psutil == found_pid_for_hwnd:
                        class_name = win32gui.GetClassName(hwnd)
                        title = win32gui.GetWindowText(hwnd)

                        if "Chrome_WidgetWin_1" not in class_name: 
                            continue 

                        if hasattr(self, 'logger'):
                            self.logger.info(f"  [HWND_PID_MATCH] 编号 {num_in_loop} (psutil PID: {target_pid_from_psutil}) MATCHES HWND_PID: {found_pid_for_hwnd} (HWND: {hwnd}, Title: '{title[:50]}', Class: '{class_name}')")
                        
                        current_hwnd_details = callback_data_map.get(num_in_loop)
                        new_is_better = False
                        if current_hwnd_details is None:
                            new_is_better = True
                        else: 
                            if not current_hwnd_details.get('title') and title: 
                                new_is_better = True
                            elif title and len(title) > len(current_hwnd_details.get('title','')) and "Google Chrome" in title :
                                new_is_better = True

                        if new_is_better:
                            callback_data_map[num_in_loop] = {
                                'hwnd': hwnd, 
                                'pid': found_pid_for_hwnd, 
                                'title': title, 
                                'class': class_name
                            }
                            if hasattr(self, 'logger'):
                                self.logger.info(f"    [HWND_STORED] For Number {num_in_loop}: HWND={hwnd}, PID={found_pid_for_hwnd}, Title='{title[:50]}'")
                        break 
            
            except Exception as e_cb: 
                if hasattr(self, 'logger'):
                    self.logger.error(f"  EnumWin CB Exception: {e_cb} for HWND {hwnd if 'hwnd' in locals() else 'Unknown'}. Error: {str(e_cb)}", exc_info=True) 
            return True 
        
        try:
            win32gui.EnumWindows(enum_windows_callback, number_to_hwnd_details)
        except Exception as e_enum:
            if hasattr(self, 'logger'): self.logger.error(f"EnumWindows 调用失败: {e_enum}", exc_info=True)
            else: print(f"ERROR: EnumWindows 调用失败: {e_enum}")

        if hasattr(self, 'logger'): self.logger.info("HWND 获取完成. 更新 self.windows:")
        
        for num, details in number_to_hwnd_details.items():
            if num in self.windows:
                self.windows[num]['hwnd_debug'] = details['hwnd']
                self.windows[num]['title_debug'] = details['title']
                if self.windows[num].get('pid') != details.get('pid'): # 使用 .get 以防万一
                    if hasattr(self, 'logger'):
                        self.logger.warning(f"  PID Mismatch for number {num}: psutil_pid={self.windows[num].get('pid')}, hwnd_pid={details.get('pid')}. Consider updating or trust hwnd_pid.")
                self.windows[num]['status'] = 'imported_with_hwnd'
                if hasattr(self, 'logger'):
                    self.logger.info(f"  编号 {num}: HWND={details['hwnd']}, Title='{details['title'][:50]}' 已关联.")
            else:
                 if hasattr(self, 'logger'): self.logger.warning(f"  编号 {num} 在 number_to_hwnd_details 中但不在 self.windows 中。")


        if hasattr(self, 'logger'):
            self.logger.info(f"最终 self.windows 状态 (共 {len(self.windows)} 条):")
            if not self.windows: self.logger.warning("  最终 self.windows 为空!")
            for n_final, w_final in self.windows.items():
                self.logger.info(f"  编号 {n_final}: PID={w_final.get('pid')}, HWND={w_final.get('hwnd_debug', 'N/A')}, Title='{w_final.get('title_debug', 'N/A')}', Status='{w_final.get('status')}'")
        else:
            print(f"INFO: 最终 self.windows 状态 (共 {len(self.windows)} 条):")
            if not self.windows: print("WARNING:   最终 self.windows 为空!")
            for n_final, w_final in self.windows.items():
                 print(f"  编号 {n_final}: PID={w_final.get('pid')}, HWND={w_final.get('hwnd_debug', 'N/A')}, Title='{w_final.get('title_debug', 'N/A')}', Status='{w_final.get('status')}'")
        
        # 确保 update_ui 和回调存在性检查的日志能够打印
        if hasattr(self, 'logger'):
            self.logger.debug(f"准备UI更新检查: update_ui={update_ui}, has ui_update_callback={hasattr(self, 'ui_update_callback')}")
            if hasattr(self, 'ui_update_callback'):
                 self.logger.debug(f"  ui_update_callback is: {self.ui_update_callback}")
        else:
            print(f"DEBUG: 准备UI更新检查: update_ui={update_ui}, has ui_update_callback={hasattr(self, 'ui_update_callback')}")
            if hasattr(self, 'ui_update_callback'):
                 print(f"  ui_update_callback is: {self.ui_update_callback}")


        if update_ui and hasattr(self, 'ui_update_callback') and self.ui_update_callback:
            windows_for_ui = []
            for num_val, win_info_val in self.windows.items():
                # 确保所有必要的键都存在，即使hwnd没有找到
                ui_entry = {
                    "number": num_val,
                    "pid": win_info_val.get('pid'),
                    "hwnd": win_info_val.get('hwnd_debug'), # UI 使用这个 hwnd
                    "title": win_info_val.get('title_debug') or f"Chrome {num_val}", # 默认标题
                    "status": win_info_val.get('status', 'unknown') 
                }
                windows_for_ui.append(ui_entry)
            
            if hasattr(self, 'logger'):
                self.logger.info(f"准备更新UI，传递 {len(windows_for_ui)} 个窗口条目: {windows_for_ui}") # 打印传递给UI的数据
            else:
                print(f"INFO: 准备更新UI，传递 {len(windows_for_ui)} 个窗口条目: {windows_for_ui}")
            
            try: # 包裹回调调用以捕获异常
                self.ui_update_callback(windows_for_ui)
                if hasattr(self, 'logger'): self.logger.info("ui_update_callback 调用完成。")
                else: print("INFO: ui_update_callback 调用完成。")
            except Exception as e_ui_callback:
                if hasattr(self, 'logger'): self.logger.error(f"ui_update_callback 调用时发生错误: {e_ui_callback}", exc_info=True)
                else: print(f"ERROR: ui_update_callback 调用时发生错误: {e_ui_callback}")
        
        return list(self.windows.values())
    
    def close_windows(self, window_handles: List[int]) -> bool:

        try:
            if self.is_syncing:
                self.stop_sync()
            
            for hwnd in window_handles:
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                except Exception as e:
                    log_error(f"关闭窗口失败: {hwnd}", e)
            
            return True
        
        except Exception as e:
            log_error("关闭窗口失败", e)
            return False
    
    def batch_open_urls(self, url: str, numbers_str: str) -> bool:
        try:
            if not url.startswith(("http://", "https://")):
                url = "https://" + url
            
            window_numbers = parse_window_numbers(numbers_str)
            
            if not self.shortcut_path or not os.path.exists(self.shortcut_path):
                return False
            
            chrome_path = self.find_chrome_path()
            if not chrome_path:
                return False
            
            if not hasattr(self, "shell") or self.shell is None:
                self.shell = win32com.client.Dispatch("WScript.Shell")
            
            success_count = 0
            for window_num in window_numbers:
                try:
                    shortcut_path = os.path.join(self.shortcut_path, f"{window_num}.lnk")
                    
                    if not os.path.exists(shortcut_path):
                        continue
                    
                    try:
                        shortcut_obj = self.shell.CreateShortCut(shortcut_path)
                        cmd_line = shortcut_obj.Arguments
                        
                        if "--user-data-dir=" in cmd_line:
                            user_data_dir = re.search(
                                r'--user-data-dir="?([^"]+)"?', cmd_line
                            )
                            if user_data_dir:
                                user_data_dir = user_data_dir.group(1)
                            else:
                                continue
                        else:
                            user_data_dir = os.path.join(
                                self.cache_dir, str(window_num)
                            )
                            if not os.path.exists(user_data_dir):
                                continue
                    except Exception as e:
                        continue
                    
                    cmd_list = [
                        chrome_path,
                        f"--user-data-dir={user_data_dir}",
                    ]
                    
                    debug_port = 9222 + window_num
                    cmd_list.insert(1, f"--remote-debugging-port={debug_port}")
                    
                    cmd_list.append(url)
                    
                    subprocess.Popen(cmd_list)
                    
                    success_count += 1
                    
                    time.sleep(0.1)
                    
                except Exception as e:
                    log_error(f"打开URL失败 (窗口 {window_num}): {str(e)}")
            
            return success_count > 0
        
        except Exception as e:
            log_error(f"批量打开网页失败: {str(e)}")
            return False
    
    def start_sync(self, window_items) -> bool:

        try:
            if not self.master_window:
                log_error("未设置主控窗口")
                return False
            
            sync_windows = []
            for item in window_items:
                hwnd = int(self.ui_manager.get_window_item_value(item, "hwnd"))
                if hwnd != self.master_window:
                    sync_windows.append(hwnd)
            
            if not sync_windows:
                log_error("没有可同步的窗口")
                return False
            
            self.is_syncing = self.hook_manager.start_sync(self.master_window, sync_windows)
            
            return self.is_syncing
        
        except Exception as e:
            log_error("启动同步失败", e)
            self.is_syncing = False
            return False
    
    def stop_sync(self) -> bool:

        try:
            self.hook_manager.stop_sync()
            self.is_syncing = False
            
            return True
        
        except Exception as e:
            log_error("停止同步失败", e)
            return False
    
    def apply_icons_to_chrome_windows(self, pid_to_number: Dict[int, int]):
        is_auto_modify_shortcut = hasattr(self, "auto_modify_shortcut_icon") and self.auto_modify_shortcut_icon
        
        hwnd_map = {}
        
        icon_paths = {}
        
        def find_chrome_windows():
            def enum_windows_callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        class_name = win32gui.GetClassName(hwnd)
                        if "Chrome_WidgetWin" in class_name:
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            window_number = pid_to_number.get(pid)
                            if window_number:
                                hwnd_map[window_number] = hwnd
                    except Exception as e:
                        log_error(f"枚举Chrome窗口失败: {hwnd}", e)
                return True
            
            try:
                win32gui.EnumWindows(enum_windows_callback, None)
            except Exception as e:
                log_error("查找Chrome窗口失败", e)
        
        def generate_all_icons():
            from utils import generate_color_icon
            
            for number in pid_to_number.values():
                try:
                    icon_path = generate_color_icon(number)
                    if icon_path and os.path.exists(icon_path):
                        icon_paths[number] = icon_path
                except Exception as e:
                    log_error(f"生成图标失败: {number}", e)
        
        def update_shortcut_icons():
            if not is_auto_modify_shortcut:
                return
                
            if not self.shortcut_path or not os.path.exists(self.shortcut_path):
                return
            
            try:
                for number, icon_path in icon_paths.items():
                    shortcut_path = os.path.join(self.shortcut_path, f"{number}.lnk")
                    if os.path.exists(shortcut_path):
                        try:
                            shortcut = self.shell.CreateShortCut(shortcut_path)
                            current_icon = shortcut.IconLocation
                            if icon_path not in current_icon:
                                shortcut.IconLocation = icon_path
                                shortcut.save()
                        except Exception as e:
                            log_error(f"为快捷方式 {number} 设置图标失败: {str(e)}")
            except Exception as e:
                log_error("更新快捷方式图标失败", e)
        
        def set_all_window_icons():
            from utils import set_chrome_icon
            
            valid_windows = []
            for number, hwnd in hwnd_map.items():
                if number in icon_paths and win32gui.IsWindow(hwnd):
                    valid_windows.append((number, hwnd, icon_paths[number]))
            
            if not valid_windows:
                return
            
            for number, hwnd, icon_path in valid_windows:
                try:
                    set_chrome_icon(hwnd, icon_path, retries=2, delay=0.1)
                except Exception as e:
                    log_error(f"设置窗口图标失败: {number}", e)
            
            try:
                ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
            except Exception as e:
                log_error("刷新图标缓存失败", e)
        
        find_chrome_windows()
        
        if not hwnd_map:
            return
        
        def execute_icon_process():
            generate_all_icons()
            
            update_shortcut_icons()
            
            set_all_window_icons()
            
            if hasattr(self, "perform_medium_cleanup"):
                if hasattr(self, "ui_manager") and self.ui_manager:
                    self.ui_manager.root.after(500, self.perform_medium_cleanup)
                else:
                    threading.Timer(0.5, self.perform_medium_cleanup).start()
        
        if hasattr(self, "ui_manager") and self.ui_manager and hasattr(self.ui_manager, "root"):
            self.ui_manager.root.after(100, execute_icon_process)
        else:
            threading.Timer(0.1, execute_icon_process).start()
    
    def arrange_windows(self, window_handles: List[int], x: int, y: int, width: int, height: int, cols: int = None) -> bool:

        try:
            if not window_handles:
                return False
                
            count = len(window_handles)
            
            if not cols:
                screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
                
                if count <= 4:
                    cols = 2
                elif count <= 9:
                    cols = 3
                elif count <= 16:
                    cols = 4
                else:
                    cols = 5
            
            rows = (count + cols - 1) // cols
            
            for i, hwnd in enumerate(window_handles):
                row = i // cols
                col = i % cols
                
                window_x = x + col * width
                window_y = y + row * height
                
                try:
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_TOP,
                        window_x, window_y, width, height,
                        win32con.SWP_SHOWWINDOW
                    )
                except Exception as e:
                    log_error(f"设置窗口位置失败: {hwnd}", e)
            
            return True
        except Exception as e:
            log_error("排列窗口失败", e)
            return False
    
    def set_master_window(self, hwnd: int) -> bool:
        log_error(f"CORE: Attempting to set master window for HWND: {hwnd}")
        try:
            if self.is_syncing:
                log_error("CORE: Stopping sync before setting new master window.")
                self.stop_sync()
            
            self.master_window = hwnd
            log_error(f"CORE: Master window set to HWND: {hwnd}")
            
            title = win32gui.GetWindowText(hwnd)
            if not "[主控]" in title and not "★" in title:
                new_title = f"★ [主控] {title} ★"
                try:
                    win32gui.SetWindowText(hwnd, new_title)
                    log_error(f"CORE: Set new title for master window: '{new_title}'")
                except Exception as e_title:
                    log_error(f"CORE: Failed to set title for master window HWND {hwnd}: {e_title}")

            try:
                dwmapi = ctypes.WinDLL("dwmapi.dll")
                color_ref = ctypes.c_uint(0x0000FF) # BGR Red color
                
                # DWMWA_BORDER_COLOR is 34
                # HRESULT DwmSetWindowAttribute(
                #   HWND    hwnd,
                #   DWORD   dwAttribute,
                #   LPCVOID pvAttribute,
                #   DWORD   cbAttribute
                # );
                # pvAttribute is a pointer to a COLORREF. COLORREF is a DWORD.
                # ctypes.c_uint is suitable here (typically 4 bytes, same as DWORD).
                
                log_error(f"CORE: Attempting to set DWMWA_BORDER_COLOR for HWND {hwnd} to RED (0x0000FF). Attribute size: {ctypes.sizeof(color_ref)}")
                
                hr = dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    self.DWMWA_BORDER_COLOR, # Attribute: DWMWA_BORDER_COLOR (34)
                    ctypes.byref(color_ref), # Value: Pointer to COLORREF (red)
                    ctypes.sizeof(color_ref) # Size: Size of COLORREF
                )
                
                if hr == 0: # S_OK
                    log_error(f"CORE: DwmSetWindowAttribute for border color SUCCEEDED for HWND {hwnd}. HRESULT: {hr}")
                else:
                    log_error(f"CORE: DwmSetWindowAttribute for border color FAILED for HWND {hwnd}. HRESULT: {hr:#010x}")
                    # You might want to use win32api.FormatMessage(hr) to get a human-readable error,
                    # but that's more involved for a quick check.
                    # For now, logging the HRESULT is a good first step.
                
                # Force the window to redraw its frame to reflect the change
                win32gui.SetWindowPos(
                    hwnd,
                    0, # No change in Z-order due to SWP_NOZORDER
                    0, 0, 0, 0, # No move, no size
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | 
                    win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED,
                )
                log_error(f"CORE: Called SetWindowPos with SWP_FRAMECHANGED for HWND {hwnd}")
            except Exception as e_dwm:
                log_error(f"CORE: Exception while setting master window border color or frame for HWND {hwnd}: {e_dwm}")
            
            return True
        
        except Exception as e:
            log_error(f"CORE: Overall failure in set_master_window for HWND {hwnd}: {e}")
            return False

    def reset_master_window(self, hwnd: int) -> bool:
        log_error(f"CORE: Attempting to reset master window style for HWND: {hwnd}")
        try:
            title = win32gui.GetWindowText(hwnd)
            if "[主控]" in title or "★" in title:
                new_title = title.replace("★ [主控] ", "").replace(" ★", "").replace("[主控] ", "")
                try:
                    win32gui.SetWindowText(hwnd, new_title)
                    log_error(f"CORE: Reset title for HWND {hwnd} to '{new_title}'")
                except Exception as e_title:
                    log_error(f"CORE: Failed to reset title for HWND {hwnd}: {e_title}")
            
            try:
                dwmapi = ctypes.WinDLL("dwmapi.dll")
                # To reset border color, you can set it to DWMWA_COLOR_NONE (0xFFFFFFFE)
                # or DWMWA_COLOR_DEFAULT (0xFFFFFFFF).
                # Setting to 0 (black) might also work if the default is no color or if black is acceptable as "reset".
                # Let's try DWMWA_COLOR_DEFAULT for a more explicit reset.
                # However, DWMWA_BORDER_COLOR documentation implies it *sets* a color.
                # A common way to "remove" the custom border is to set it to a transparent color if supported,
                # or simply to a color that matches the default window frame, or rely on the system
                # to not draw it if the attribute is "cleared".
                # For now, let's try setting it to a value that DWM might interpret as "no custom border" or default.
                # The original code used c_int(0), which is black.
                # For robust reset, DWM documentation should be consulted for DWMWA_BORDER_COLOR.
                # If we simply want to revert our *specific* red color, and the default is no color,
                # then perhaps setting it to a color that won't be drawn or is transparent is the way.
                # The API might not have a direct "clear attribute" for this.
                # Using a value like 0 (black) or a specific "none" value like DWMWA_COLOR_NONE (0xFFFFFFFE).
                color_none = ctypes.c_uint(0xFFFFFFFE) # DWMWA_COLOR_NONE

                log_error(f"CORE: Attempting to reset DWMWA_BORDER_COLOR for HWND {hwnd} using DWMWA_COLOR_NONE. Attribute size: {ctypes.sizeof(color_none)}")

                hr = dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    self.DWMWA_BORDER_COLOR,
                    ctypes.byref(color_none), # Try DWMWA_COLOR_NONE
                    ctypes.sizeof(color_none)
                )

                if hr == 0: # S_OK
                    log_error(f"CORE: DwmSetWindowAttribute for resetting border color SUCCEEDED for HWND {hwnd}. HRESULT: {hr}")
                else:
                    log_error(f"CORE: DwmSetWindowAttribute for resetting border color FAILED for HWND {hwnd}. HRESULT: {hr:#010x}")

            except Exception as e_dwm_reset:
                log_error(f"CORE: Exception while resetting DWM border color for HWND {hwnd}: {e_dwm_reset}")

            # Reset TopMost if it was set
            # GWL_EXSTYLE = -20
            # ex_style = win32gui.GetWindowLong(hwnd, GWL_EXSTYLE)
            # if ex_style & win32con.WS_EX_TOPMOST:
            #     win32gui.SetWindowPos(
            #         hwnd,
            #         win32con.HWND_NOTOPMOST,
            #         0, 0, 0, 0,
            #         win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
            #     )
            #     log_error(f"CORE: Reset HWND_NOTOPMOST for HWND {hwnd}")

            # Force frame redraw
            win32gui.SetWindowPos(
                hwnd,
                0, 
                0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | 
                win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED
            )
            log_error(f"CORE: Called SetWindowPos with SWP_FRAMECHANGED for HWND {hwnd} during reset")
            
            return True
        except Exception as e:
            log_error(f"CORE: Overall failure in reset_master_window for HWND {hwnd}: {e}")
            return False
    
    def create_environments(self, numbers_str: str) -> bool:
        try:
            cache_dir = self.cache_dir
            shortcut_dir = self.shortcut_path
            
            if not all([cache_dir, shortcut_dir]):
                return False

            os.makedirs(cache_dir, exist_ok=True)
            os.makedirs(shortcut_dir, exist_ok=True)

            chrome_path = self.find_chrome_path()
            if not chrome_path:
                return False
                
            if not os.path.isfile(chrome_path):
                return False
                
            try:
                with open(chrome_path, 'rb') as f:
                    pass
            except Exception as e:
                return False

            shell = win32com.client.Dispatch("WScript.Shell")

            window_numbers = parse_window_numbers(numbers_str)
            
            if not window_numbers:
                return False

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

            if created_count > 0:
                return True
            else:
                return False

        except Exception as e:
            return False
    
    def find_chrome_path(self) -> str:

        try:
            import winreg
            
            registry_paths = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Google\Chrome\Application"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Google\Chrome\Application"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Google\Chrome\Application"),
            ]
            
            for root_key, sub_key in registry_paths:
                try:
                    with winreg.OpenKey(root_key, sub_key) as key:
                        version, _ = winreg.QueryValueEx(key, "Version")
                        if version:
                            possible_paths = [
                                os.path.join(os.path.dirname(sub_key), "Application", version, "chrome.exe"),
                                os.path.join(os.path.dirname(sub_key), version, "chrome.exe"),
                            ]
                            for path in possible_paths:
                                if os.path.exists(path):
                                    return path
                except Exception:
                    continue
            
            common_paths = [
                os.path.join(os.environ.get('PROGRAMFILES', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
                os.path.join(os.environ.get('PROGRAMFILES(X86)', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
                os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
            ]
            
            for path in common_paths:
                if os.path.exists(path):
                    return path
                    
            try:
                import subprocess
                result = subprocess.run(['where', 'chrome.exe'], 
                                     capture_output=True, 
                                     text=True, 
                                     creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    chrome_path = result.stdout.strip().split('\n')[0]
                    if os.path.exists(chrome_path):
                        return chrome_path
            except Exception:
                pass
            
            return ""
            
        except Exception as e:
            return ""
    
    def keep_only_current_tab(self, window_items) -> bool:
        try:
            selected = []
            for item in window_items:
                hwnd = int(self.ui_manager.get_window_item_value(item, "hwnd"))
                window_num = int(self.ui_manager.get_window_item_value(item, "number"))
                selected.append((window_num, hwnd))
            
            if not selected:
                log_error("没有选中的窗口")
                return False
            
            if not hasattr(self, "debug_ports") or not self.debug_ports:
                log_error("未找到调试端口映射，尝试重建...")
                self.debug_ports = {
                    window_num: 9222 + window_num for window_num, _ in selected
                }
            
            import requests
            import concurrent.futures
            
            def process_tabs():
                try:
                    import traceback
                    
                    port_to_tabs = {}
                    
                    def get_tabs(window_data):
                        window_num, _ = window_data
                        if window_num in self.debug_ports:
                            port = self.debug_ports[window_num]
                            try:
                                response = requests.get(
                                    f"http://localhost:{port}/json", timeout=0.5
                                )
                                if response.status_code == 200:
                                    tabs = response.json()
                                    page_tabs = [
                                        tab for tab in tabs if tab.get("type") == "page"
                                    ]
                                    if len(page_tabs) > 1:
                                        return port, page_tabs, window_num
                            except Exception as e:
                                log_error(f"获取窗口{window_num}的标签页失败: {str(e)}")
                        return None
                    
                    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
                        futures = []
                        for window_data in selected:
                            futures.append(executor.submit(get_tabs, window_data))
                        
                        for future in concurrent.futures.as_completed(futures):
                            result = future.result()
                            if result:
                                port, tabs, window_num = result
                                port_to_tabs[port] = (tabs, window_num)
                    
                    if not port_to_tabs:
                        return
                    
                    close_requests = []
                    
                    for port, (tabs, window_num) in port_to_tabs.items():
                        keep_tab = tabs[0]
                        to_close = []
                        for tab in tabs:
                            if tab.get("id") != keep_tab.get("id"):
                                to_close.append((port, tab.get("id")))
                        close_requests.extend(to_close)
                    
                    def close_tab(request):
                        port, tab_id = request
                        try:
                            requests.get(
                                f"http://localhost:{port}/json/close/{tab_id}", timeout=0.5
                            )
                            return True
                        except Exception as e:
                            log_error(f"关闭标签页失败: {str(e)}")
                            return False
                    
                    with concurrent.futures.ThreadPoolExecutor(max_workers=30) as executor:
                        futures = [
                            executor.submit(close_tab, req) for req in close_requests
                        ]
                        for future in concurrent.futures.as_completed(futures):
                            future.result()
                
                except Exception as e:
                    log_error(f"处理标签页时出错: {str(e)}")
                    traceback.print_exc()
                    return False
                
                return True
            
            thread = threading.Thread(target=process_tabs, daemon=True)
            thread.start()
            return True
            
        except Exception as e:
            log_error(f"仅保留当前标签页失败: {str(e)}")
            return False
    
    def keep_only_new_tab(self, window_items) -> bool:
        try:
            selected = []
            for item in window_items:
                hwnd = int(self.ui_manager.get_window_item_value(item, "hwnd"))
                window_num = int(self.ui_manager.get_window_item_value(item, "number"))
                selected.append((window_num, hwnd))
            
            if not selected:
                log_error("没有选中的窗口")
                return False
            
            if not hasattr(self, "debug_ports") or not self.debug_ports:
                log_error("未找到调试端口映射，尝试重建...")
                self.debug_ports = {
                    window_num: 9222 + window_num for window_num, _ in selected
                }
            
            import requests
            import concurrent.futures
            
            def process_tabs():
                try:
                    import traceback
                    
                    window_tabs = {}
                    
                    def get_tabs(window_data):
                        window_num, _ = window_data
                        if window_num in self.debug_ports:
                            port = self.debug_ports[window_num]
                            try:
                                response = requests.get(
                                    f"http://localhost:{port}/json", timeout=0.5
                                )
                                if response.status_code == 200:
                                    tabs = response.json()
                                    page_tabs = [
                                        tab.get("id")
                                        for tab in tabs
                                        if tab.get("type") == "page"
                                    ]
                                    if page_tabs:
                                        return port, page_tabs, window_num
                            except Exception as e:
                                log_error(f"获取窗口{window_num}的标签页失败: {str(e)}")
                        return None
                    
                    valid_ports = []
                    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
                        futures = []
                        for window_data in selected:
                            futures.append(executor.submit(get_tabs, window_data))
                        
                        for future in concurrent.futures.as_completed(futures):
                            result = future.result()
                            if result:
                                port, tabs, window_num = result
                                window_tabs[port] = (tabs, window_num)
                                valid_ports.append(port)
                    
                    if not valid_ports:
                        return
                    
                    created_tabs = {}
                    
                    def create_new_tab(port_data):
                        port, window_num = port_data
                        try:
                            requests.put(
                                f"http://localhost:{port}/json/new?chrome://newtab/",
                                timeout=0.5,
                            )
                            return port, window_num, True
                        except Exception as e:
                            log_error(f"为窗口 {window_num} 创建新标签页失败: {str(e)}")
                            return port, window_num, False
                    
                    port_to_window = {
                        port: window_num for port, (_, window_num) in window_tabs.items()
                    }
                    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
                        futures = [
                            executor.submit(create_new_tab, (port, port_to_window[port]))
                            for port in valid_ports
                        ]
                        for future in concurrent.futures.as_completed(futures):
                            port, window_num, success = future.result()
                            if success:
                                created_tabs[window_num] = port
                    
                    def close_old_tabs(port_data):
                        port, tabs, window_num = port_data
                        for tab_id in tabs:
                            try:
                                requests.get(
                                    f"http://localhost:{port}/json/close/{tab_id}",
                                    timeout=0.5,
                                )
                            except Exception as e:
                                log_error(f"关闭窗口 {window_num} 的标签页失败: {str(e)}")
                    
                    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
                        futures = []
                        for window_num, port in created_tabs.items():
                            tabs, _ = window_tabs[port]
                            futures.append(
                                executor.submit(close_old_tabs, (port, tabs, window_num))
                            )
                        
                        for future in concurrent.futures.as_completed(futures):
                            future.result()
                    
                    return True
                    
                except Exception as e:
                    log_error(f"处理标签页时出错: {str(e)}")
                    traceback.print_exc()
                    return False
            
            thread = threading.Thread(target=process_tabs, daemon=True)
            thread.start()
            return True
            
        except Exception as e:
            log_error(f"仅保留新标签页失败: {str(e)}")
            return False

    def update_screen_info(self) -> List[Dict]:

        try:
            self.screens = []
            self.screen_names = []

            def callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
                monitor_info = win32api.GetMonitorInfo(hMonitor)
                monitor_rect = monitor_info["Monitor"]
                work_rect = monitor_info["Work"]
                
                device = monitor_info.get("Device", "")
                display_name = f"屏幕 {len(self.screens) + 1}"
                
                primary = monitor_info.get("Flags", 0) & 1 == 1
                if primary:
                    display_name += " (主)"
                
                screen_info = {
                    "name": display_name,
                    "monitor_rect": monitor_rect,
                    "work_rect": work_rect,
                    "primary": primary,
                    "device": device,
                    "monitor": hMonitor
                }
                
                self.screens.append(screen_info)
                self.screen_names.append(display_name)
                
                return True
            
            try:
                MONITORENUMPROC = ctypes.WINFUNCTYPE(
                    ctypes.c_bool,
                    ctypes.c_ulong,
                    ctypes.c_ulong,
                    ctypes.POINTER(wintypes.RECT),
                    ctypes.c_double
                )
                
                if ctypes.windll.user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(callback), 0) == 0:
                    self.screens = []
                    self.screen_names = []
                    
                    virtual_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
                    virtual_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
                    virtual_left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
                    virtual_top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
                    
                    primary_monitor = win32api.MonitorFromPoint((0, 0), win32con.MONITOR_DEFAULTTOPRIMARY)
                    primary_info = win32api.GetMonitorInfo(primary_monitor)
                    
                    self.screens.append({
                        "name": "屏幕 1 (主)",
                        "monitor_rect": primary_info["Monitor"],
                        "work_rect": primary_info["Work"],
                        "primary": True,
                        "device": "",
                        "monitor": primary_monitor
                    })
                    self.screen_names.append("屏幕 1 (主)")
                    
                    try:
                        second_point = (virtual_left + virtual_width - 1, virtual_top + virtual_height // 2)
                        second_monitor = win32api.MonitorFromPoint(second_point, win32con.MONITOR_DEFAULTTONULL)
                        if second_monitor and second_monitor != primary_monitor:
                            second_info = win32api.GetMonitorInfo(second_monitor)
                            self.screens.append({
                                "name": "屏幕 2",
                                "monitor_rect": second_info["Monitor"],
                                "work_rect": second_info["Work"],
                                "primary": False,
                                "device": "",
                                "monitor": second_monitor
                            })
                            self.screen_names.append("屏幕 2")
                    except Exception as e:
                        log_error(f"获取第二屏幕信息失败: {str(e)}")
                    
            except Exception as e:
                log_error("枚举显示器失败", e)
                
                screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
                
                monitor_rect = (0, 0, screen_width, screen_height)
                work_rect = monitor_rect
                
                screen_info = {
                    "name": "主屏幕",
                    "monitor_rect": monitor_rect,
                    "work_rect": work_rect,
                    "primary": True,
                    "device": "",
                    "monitor": None
                }
                
                self.screens.append(screen_info)
                self.screen_names.append("主屏幕")
            
            self.screens.sort(key=lambda x: x["monitor_rect"][0])
            
            self.screen_names = [screen["name"] for screen in self.screens]
            
            return self.screens
        except Exception as e:
            log_error("更新屏幕信息失败", e)
            self.screens = []
            self.screen_names = []
            return []

    def auto_arrange_windows(self, selected_windows: List[Tuple[int, int, Any]]) -> bool:
        try:
            log_error("CORE: 进入 auto_arrange_windows")
            if not selected_windows:
                log_error("CORE: auto_arrange_windows - 没有选中的窗口")
                return False
            log_error(f"CORE: auto_arrange_windows - 选中了 {len(selected_windows)} 个窗口")

            # 更新并检查屏幕信息
            self.update_screen_info() # 确保屏幕信息是最新的
            if not self.screens or self.screen_selection == "":
                log_error("CORE: auto_arrange_windows - 未检测到屏幕或未选择屏幕。")
                return False

            screen_index = 0
            for i, name in enumerate(self.screen_names):
                if name == self.screen_selection:
                    screen_index = i
                    break
            
            if screen_index >= len(self.screens):
                log_error(f"CORE: auto_arrange_windows - 屏幕索引 {screen_index} 无效 (共 {len(self.screens)} 个屏幕)。")
                # 尝试使用第一个屏幕作为回退
                if self.screens:
                    screen_index = 0
                    log_error("CORE: auto_arrange_windows - 回退到使用屏幕索引 0。")
                else:
                    log_error("CORE: auto_arrange_windows - 没有可用的屏幕进行回退。")
                    return False
            
            target_screen = self.screens[screen_index]
            screen_rect = target_screen.get("work_rect")
            if not screen_rect or len(screen_rect) < 4:
                log_error(f"CORE: auto_arrange_windows - 目标屏幕 {target_screen.get('name')} 的工作区信息无效: {screen_rect}")
                return False

            screen_width = screen_rect[2] - screen_rect[0]
            screen_height = screen_rect[3] - screen_rect[1]
            
            log_error(f"CORE: auto_arrange_windows - 使用屏幕: {target_screen.get('name')}, 工作区: {screen_rect}, 屏幕尺寸: {screen_width}x{screen_height}")

            count = len(selected_windows)
            
            # from utils import is_ultrawide_screen # 假设这个工具函数存在
            # if is_ultrawide_screen(screen_width, screen_height):
            #     aspect_ratio = screen_width / screen_height
            #     cols = int(math.sqrt(count * aspect_ratio / 1.77))
            #     cols = max(cols, 2) # 确保至少有2列
            # else:
            #     cols = int(math.sqrt(count))
            #     if cols * cols < count:
            #         cols += 1
            
            # 简化 cols 计算逻辑，确保 cols > 0
            if count == 0: return True # 没有窗口需要排列
            cols = int(math.sqrt(count))
            if cols == 0: cols = 1
            if cols * cols < count and count > 1 : # 避免单窗口时增加列数
                 cols +=1
            
            if cols == 0: # 再次检查，防止除零
                log_error("CORE: auto_arrange_windows - 计算出的列数为0，中止排列。")
                return False

            rows = (count + cols - 1) // cols
            if rows == 0: rows = 1 # 确保行数不为0

            log_error(f"CORE: auto_arrange_windows - 窗口布局: {rows}行 x {cols}列")

            if screen_width <= 0 or screen_height <=0 or cols <= 0 or rows <= 0:
                log_error(f"CORE: auto_arrange_windows - 屏幕尺寸或行列计算无效: screen_width={screen_width}, screen_height={screen_height}, cols={cols}, rows={rows}")
                return False

            width = screen_width // cols
            height = screen_height // rows
            log_error(f"CORE: auto_arrange_windows - 计算出的每个窗口大小: {width}x{height}")


            if width <= 0 or height <= 0:
                log_error(f"CORE: auto_arrange_windows - 计算出的窗口宽或高无效: width={width}, height={height}. 中止排列。")
                return False

            positions = []
            for i in range(count):
                row_idx = i // cols
                col_idx = i % cols
                x_pos = screen_rect[0] + col_idx * width
                y_pos = screen_rect[1] + row_idx * height
                positions.append((x_pos, y_pos))
                log_error(f"CORE: auto_arrange_windows - 窗口 {i} 目标位置: ({x_pos}, {y_pos})")

            for i, (number, hwnd, _) in enumerate(selected_windows):
                try:
                    x, y = positions[i]
                    log_error(f"CORE: auto_arrange_windows - 处理窗口编号 {number} (HWND: {hwnd})")
                    log_error(f"    目标参数: x={x}, y={y}, width={width}, height={height}")

                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    
                    win32gui.SetWindowPos(
                        hwnd, 
                        win32con.HWND_TOP,
                        x, y, width, height,
                        win32con.SWP_SHOWWINDOW | win32con.SWP_NOZORDER
                    )
                    
                    win32gui.UpdateWindow(hwnd)
                    log_error(f"CORE: auto_arrange_windows - 窗口 {number} SetWindowPos 调用完成")

                except Exception as e_inner:
                    log_error(f"CORE: auto_arrange_windows - 移动窗口 {number} (HWND: {hwnd}) 失败: {str(e_inner)}")
                    continue
            
            log_error("CORE: auto_arrange_windows - 所有窗口处理完毕")
            # 后续置顶逻辑...
            self.set_window_priority([hwnd for _, hwnd, _ in selected_windows])
            master_hwnd = None
            # ... (尝试激活主控窗口的逻辑，确保 self.ui_manager.window_list 存在或有其他方式获取主控)
            # 这个部分依赖于 ui_manager 和 window_list，如果 core 直接调用，可能需要调整
            # 暂时注释掉与 self.ui_manager.window_list 相关的部分，因为在 core 中可能没有直接访问权
            # if hasattr(self, 'ui_manager') and self.ui_manager and hasattr(self.ui_manager, 'window_list'):
            #     for item_id_ui in self.ui_manager.window_list.get_children():
            #         if self.ui_manager.window_list.set(item_id_ui, "master") == "√":
            #             values_ui = self.ui_manager.window_list.item(item_id_ui)["values"]
            #             if values_ui and len(values_ui) >= 5:
            #                 master_hwnd = int(values_ui[4]) # hwnd 在第5列
            #                 break
            # if master_hwnd:
            #     self.activate_window(master_hwnd)

            return True
        except Exception as e:
            log_error(f"CORE: auto_arrange_windows - 整体失败: {str(e)}")
            return False

    def custom_arrange_on_single_screen(self, windows: List[Tuple[Any, int]], screen_index_from_ui: int, 
                                       start_x_str: str, start_y_str: str, width_str: str, height_str: str, 
                                       h_spacing_str: str, v_spacing_str: str, windows_per_row_str: str) -> bool:
        try:
            log_error(f"CORE: 进入 custom_arrange_on_single_screen. UI传入屏幕索引: {screen_index_from_ui}")
            
            # 更新并检查屏幕信息
            self.update_screen_info()
            if not self.screens:
                log_error("CORE: custom_arrange - 未检测到屏幕。")
                return False

            # 使用UI传入的索引，但要校验
            screen_index = int(screen_index_from_ui) # 确保是整数
            if not (0 <= screen_index < len(self.screens)):
                log_error(f"CORE: custom_arrange - UI传入屏幕索引 {screen_index} 无效 (共 {len(self.screens)} 个屏幕)。尝试使用屏幕0。")
                if self.screens: # 确保至少有一个屏幕
                    screen_index = 0
                else:
                    return False # 没有屏幕可选

            target_screen = self.screens[screen_index]
            screen_rect = target_screen.get("work_rect")
            if not screen_rect or len(screen_rect) < 4:
                log_error(f"CORE: custom_arrange - 目标屏幕 {target_screen.get('name')} 的工作区信息无效: {screen_rect}")
                return False

            screen_x_val = screen_rect[0]
            screen_y_val = screen_rect[1]
            # screen_width_val = screen_rect[2] - screen_rect[0] # 用于超宽屏判断，暂时不用
            # screen_height_val = screen_rect[3] - screen_rect[1]

            log_error(f"CORE: custom_arrange - 使用屏幕: {target_screen.get('name')}, 工作区: {screen_rect}")

            # 参数转换和校验
            try:
                start_x = int(start_x_str)
                start_y = int(start_y_str)
                width = int(width_str)
                height = int(height_str)
                h_spacing = int(h_spacing_str)
                v_spacing = int(v_spacing_str)
                windows_per_row = int(windows_per_row_str)
                log_error(f"    转换后参数: start_x={start_x}, start_y={start_y}, width={width}, height={height}, h_spacing={h_spacing}, v_spacing={v_spacing}, windows_per_row={windows_per_row}")
            except ValueError as ve:
                log_error(f"CORE: custom_arrange - 参数转换失败: {ve}")
                return False

            if width <= 0 or height <= 0 or windows_per_row <= 0:
                log_error(f"CORE: custom_arrange - 宽度、高度或每行窗口数无效: width={width}, height={height}, per_row={windows_per_row}")
                return False

            current_width_to_use = width 

            abs_start_x = screen_x_val + start_x
            abs_start_y = screen_y_val + start_y
            log_error(f"    绝对起始位置: ({abs_start_x}, {abs_start_y})")
            
            for i, (item_data, hwnd) in enumerate(windows):
                row_idx = i // windows_per_row
                col_idx = i % windows_per_row
                
                x_pos = abs_start_x + col_idx * (current_width_to_use + h_spacing)
                y_pos = abs_start_y + row_idx * (height + v_spacing)
                
                log_error(f"CORE: custom_arrange - 处理窗口 (HWND: {hwnd})")
                log_error(f"    目标参数: x={x_pos}, y={y_pos}, width={current_width_to_use}, height={height}")
                
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                
                win32gui.SetWindowPos(
                    hwnd, 
                    0, 
                    x_pos, y_pos, current_width_to_use, height,
                    win32con.SWP_SHOWWINDOW | win32con.SWP_NOZORDER
                )
                
                win32gui.UpdateWindow(hwnd)
                log_error(f"CORE: custom_arrange - 窗口 {hwnd} SetWindowPos 调用完成")

            # 后续置顶逻辑
            self.set_window_priority([hwnd for _, hwnd in windows])
            # 激活主控窗口的逻辑与 auto_arrange_windows 中类似，可能需要调整或暂时移除
            # master_hwnd = None
            # ...
            # if master_hwnd:
            #     self.activate_window(master_hwnd)

            return True
        except Exception as e:
            log_error(f"CORE: custom_arrange_on_single_screen - 整体失败: {type(e).__name__} - {str(e)}")
            return False
    
    def set_window_priority(self, window_handles: List[int]) -> bool:
        try:
            for hwnd in window_handles:
                try:
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_TOPMOST,
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
                    )
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_NOTOPMOST,
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
                    )
                except Exception as e:
                    log_error(f"设置窗口 {hwnd} 优先级失败", e)
                    
            return True
        except Exception as e:
            log_error("设置窗口优先级失败", e)
            return False
    
    def activate_window(self, hwnd: int) -> bool:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            return True
        except Exception as e:
            log_error(f"激活窗口失败: {hwnd}", e)
            return False

    def get_active_screens(self) -> List[Dict]:
        try:
            self.update_screen_info()
            
            if hasattr(self, 'ui_manager') and self.ui_manager is not None:
                if hasattr(self.ui_manager, 'screen_selection'):
                    self.screen_selection = self.ui_manager.screen_selection
                    log_error(f"从UI同步了屏幕选择: {self.screen_selection}")
            
            if not self.screens:
                log_error("未检测到任何屏幕")
                return []
                
            screen_selection = self.screen_selection
            log_error(f"当前屏幕选择: {screen_selection}")
            
            if not screen_selection or screen_selection == "自动" or screen_selection == "所有屏幕":
                log_error("使用所有可用屏幕")
                return self.screens
                
            selected_screen = None
            for screen in self.screens:
                if screen["name"] == screen_selection:
                    selected_screen = screen
                    break
                    
            if selected_screen:
                log_error(f"使用指定屏幕: {selected_screen['name']}")
                return [selected_screen]
                
            for screen in self.screens:
                if "(主)" in screen["name"]:
                    log_error(f"找不到指定屏幕 '{screen_selection}'，使用主屏幕")
                    return [screen]
                    
            log_error(f"找不到指定屏幕 '{screen_selection}' 和主屏幕，使用第一个可用屏幕")
            return [self.screens[0]] if self.screens else []
            
        except Exception as e:
            log_error("获取活跃屏幕失败", e)
            return []
    
    def custom_arrange_on_multiple_screens(self, windows: List[Tuple[Any, int]], active_screens: List[Dict],
                                          start_x: int, start_y: int, width: int, height: int,
                                          h_spacing: int, v_spacing: int, windows_per_row: int) -> bool:
        try:
            if not windows:
                log_error("没有窗口需要排列")
                return False
                
            if not active_screens:
                log_error("没有可用的活跃屏幕")
                return False
                
            log_error(f"多屏幕排列: {len(windows)} 个窗口, {len(active_screens)} 个活跃屏幕")
            
            total_windows = len(windows)
            screens_count = len(active_screens)
            
            screen_weights = []
            total_width = sum(s["work_rect"][2] - s["work_rect"][0] for s in active_screens)
            
            for screen in active_screens:
                screen_width = screen["work_rect"][2] - screen["work_rect"][0]
                weight = screen_width / total_width
                screen_weights.append(weight)
            
            windows_per_screen = []
            remainder = total_windows
            
            for i in range(screens_count - 1):
                count = int(total_windows * screen_weights[i])
                windows_per_screen.append(count)
                remainder -= count
            
            windows_per_screen.append(remainder)
            
            allocation = []
            start_index = 0
            
            for i in range(screens_count):
                count = windows_per_screen[i]
                
                if count > 0:
                    end_index = start_index + count
                    screen_windows = windows[start_index:end_index]
                    
                    allocation.append({
                        "screen": active_screens[i],
                        "windows": screen_windows
                    })
                    
                    start_index = end_index
            
            for alloc in allocation:
                screen = alloc["screen"]
                screen_windows = alloc["windows"]
                
                screen_rect = screen["work_rect"]  
                screen_x = screen_rect[0]  
                screen_y = screen_rect[1]  
                screen_width = screen_rect[2] - screen_rect[0]
                screen_height = screen_rect[3] - screen_rect[1]
                
                log_error(f"在屏幕 {screen['name']} 上排列 {len(screen_windows)} 个窗口")
                log_error(f"屏幕工作区: {screen_rect}")
                
                from utils import is_ultrawide_screen
                
                current_windows_per_row = windows_per_row
                current_width = width
                
                if is_ultrawide_screen(screen_width, screen_height) and windows_per_row == 5:
                    windows_count = len(screen_windows)
                    current_windows_per_row = min(int(windows_count * 0.5), 8)
                    if current_windows_per_row < 2:
                        current_windows_per_row = 2
                    adjusted_width = min(screen_width // current_windows_per_row - 20, width)
                    current_width = adjusted_width
                
                abs_start_x = screen_x + start_x
                abs_start_y = screen_y + start_y
                log_error(f"调整后的起始位置: ({abs_start_x}, {abs_start_y})")
                
                for i, (_, hwnd) in enumerate(screen_windows):
                    row = i // current_windows_per_row
                    col = i % current_windows_per_row
                    
                    x = abs_start_x + col * (current_width + h_spacing)
                    y = abs_start_y + row * (height + v_spacing)
                    
                    log_error(f"窗口在屏幕 {screen['name']} 上的位置: ({x}, {y})")
                    
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    
                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                    style |= win32con.WS_SIZEBOX | win32con.WS_SYSMENU
                    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
                    win32gui.SetWindowPos(
                        hwnd, 
                        win32con.HWND_TOP,
                        x, y, current_width, height,
                        win32con.SWP_SHOWWINDOW | win32con.SWP_FRAMECHANGED
                    )
                    
                    win32gui.MoveWindow(hwnd, x, y, current_width, height, True)
                    win32gui.UpdateWindow(hwnd)
                    win32gui.RedrawWindow(
                        hwnd, 
                        None, 
                        None, 
                        win32con.RDW_INVALIDATE | win32con.RDW_ERASE | 
                        win32con.RDW_FRAME | win32con.RDW_ALLCHILDREN
                    )

            for _, hwnd in windows:
                try:
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_TOPMOST,
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                    )
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_NOTOPMOST,
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                    )
                except Exception as e:
                    log_error(f"设置窗口置顶失败: {str(e)}")
            
            return True
        except Exception as e:
            log_error("多屏幕自定义排列失败", e)
            return False
            
    def auto_arrange_multi_screens(self, selected_windows: List[Tuple[int, int, Any]]) -> bool:

        try:
            if not selected_windows:
                log_error("没有选中的窗口")
                return False
                
            log_error(f"多屏幕自动排列: 选中了 {len(selected_windows)} 个窗口")
            
            active_screens = self.get_active_screens()
            
            if not active_screens:
                log_error("没有找到可用的屏幕")
                return False
            
            log_error(f"找到 {len(active_screens)} 个活跃屏幕")
            
            selected_windows.sort(key=lambda x: x[0])
            
            total_windows = len(selected_windows)
            screens_count = len(active_screens)
            
            windows_per_screen = total_windows // screens_count
            remainder = total_windows % screens_count
            
            allocation = []
            start_index = 0
            
            for i in range(screens_count):
                count = windows_per_screen + (1 if i < remainder else 0)
                
                if count > 0:
                    end_index = start_index + count
                    screen_windows = selected_windows[start_index:end_index]
                    
                    allocation.append({
                        "screen": active_screens[i],
                        "windows": screen_windows
                    })
                    
                    start_index = end_index
            
            for alloc in allocation:
                screen = alloc["screen"]
                screen_windows = alloc["windows"]
                
                screen_rect = screen["work_rect"]  
                screen_x = screen_rect[0]  
                screen_y = screen_rect[1]  
                
                screen_width = screen_rect[2] - screen_rect[0]
                screen_height = screen_rect[3] - screen_rect[1]
                
                log_error(f"在屏幕 {screen['name']} 上排列 {len(screen_windows)} 个窗口")
                log_error(f"屏幕工作区: {screen_rect}")
                
                count = len(screen_windows)
                cols = int(math.sqrt(count))
                if cols * cols < count:
                    cols += 1
                rows = (count + cols - 1) // cols
                
                width = screen_width // cols
                height = screen_height // rows
                
                positions = []
                for i in range(count):
                    row = i // cols
                    col = i % cols
                    x = screen_x + col * width
                    y = screen_y + row * height
                    positions.append((x, y))
                
                for i, (number, hwnd, _) in enumerate(screen_windows):
                    try:
                        if i < len(positions):
                            x, y = positions[i]
                            log_error(f"移动窗口 {number} (句柄: {hwnd}) 到屏幕 {screen['name']} 的位置 ({x}, {y})")
                            
                            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                            
                            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                            style |= win32con.WS_SIZEBOX | win32con.WS_SYSMENU
                            win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
                            win32gui.SetWindowPos(
                                hwnd, 
                                win32con.HWND_TOP,
                                x, y, width, height,
                                win32con.SWP_SHOWWINDOW | win32con.SWP_FRAMECHANGED
                            )
                            
                            win32gui.MoveWindow(hwnd, x, y, width, height, True)                           
                            win32gui.UpdateWindow(hwnd)
                            win32gui.RedrawWindow(
                                hwnd, 
                                None, 
                                None, 
                                win32con.RDW_INVALIDATE | win32con.RDW_ERASE | 
                                win32con.RDW_FRAME | win32con.RDW_ALLCHILDREN
                            )
                            
                            log_error(f"窗口 {number} 移动成功")
                            
                    except Exception as e:
                        log_error(f"移动窗口 {number} (句柄: {hwnd}) 失败: {str(e)}")
                        continue
            
            for number, hwnd, _ in selected_windows:
                try:
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_TOPMOST,
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                    )
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_NOTOPMOST,
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                    )
                except Exception as e:
                    log_error(f"设置窗口 {hwnd} 置顶失败: {str(e)}")
            
            log_error("多屏幕窗口排列完成")
            return True
            
        except Exception as e:
            log_error(f"多屏幕自动排列失败: {str(e)}")
            return False

    def is_profile_running(self, profile_number: int) -> bool:
        """检查具有给定编号的配置文件当前是否被认为正在运行 (基于 self.windows)。"""
        # self.windows 的键是 profile_number (int)
        win_info = self.windows.get(profile_number)
        if win_info and win_info.get('hwnd_debug'): # 或者 win_info.get('pid') 存在且psutil.pid_exists
            # 更严格的检查可以是hwnd_debug存在且有效
            try:
                if win32gui.IsWindow(win_info['hwnd_debug']):
                    return True
            except: # win32gui.IsWindow 可能会在句柄无效时抛出pywintypes.error
                return False
        elif win_info and win_info.get('pid') and psutil.pid_exists(win_info.get('pid')):
            return True # 如果只有PID，也认为它在某种程度上是"活动的"
        return False

    def get_valid_profiles_for_sequential_launch(self, numbers_range_str: str) -> List[int]:
        """
        根据范围字符串，获取所有有效的、存在的、且当前未运行的快捷方式对应的配置文件编号列表。
        用于初始化"依次启动"序列。
        返回排序后的数字列表。
        """
        if not self.shortcut_path or not os.path.exists(self.shortcut_path):
            log_error("CORE: get_valid_profiles_for_sequential_launch - 快捷方式目录未设置或不存在。")
            return []

        try:
            target_numbers = parse_window_numbers(numbers_range_str) # utils.py中的函数
        except ValueError as e:
            log_error(f"CORE: get_valid_profiles_for_sequential_launch - 解析范围字符串失败: {e}")
            raise # 重新抛出，让UI层捕获并显示给用户

        valid_and_not_running = []
        for num in sorted(list(set(target_numbers))): # 排序并去重
            shortcut_file = os.path.join(self.shortcut_path, f"{num}.lnk")
            if os.path.exists(shortcut_file):
                if not self.is_profile_running(num):
                    valid_and_not_running.append(num)
                else:
                    log_error(f"CORE: get_valid_profiles_for_sequential_launch - 编号 {num} 已在运行，从序列中排除。")
            else:
                log_error(f"CORE: get_valid_profiles_for_sequential_launch - 快捷方式 {num}.lnk 不存在，从序列中排除。")
        
        log_error(f"CORE: get_valid_profiles_for_sequential_launch - 范围 '{numbers_range_str}', 有效且未运行的编号: {valid_and_not_running}")
        return valid_and_not_running


    def launch_random_profiles(self, numbers_range_str: str, count: int) -> List[int]:
        """
        从指定数字范围内的有效且未运行的快捷方式中，随机选择指定数量的进行启动。
        返回成功发送启动命令的编号列表。
        """
        log_error(f"CORE: launch_random_profiles - Range: '{numbers_range_str}', Count: {count}")
        if not self.shortcut_path or not os.path.exists(self.shortcut_path):
            log_error("CORE: launch_random_profiles - 快捷方式目录未设置或不存在。")
            return []

        try:
            all_possible_numbers_in_range = parse_window_numbers(numbers_range_str)
        except ValueError as e:
            log_error(f"CORE: launch_random_profiles - 解析范围字符串失败: {e}")
            raise # 让UI层处理

        available_for_launch = []
        for num in sorted(list(set(all_possible_numbers_in_range))): # 排序并去重
            shortcut_file = os.path.join(self.shortcut_path, f"{num}.lnk")
            if os.path.exists(shortcut_file):
                if not self.is_profile_running(num): # 假设 is_profile_running 检查 self.windows
                    available_for_launch.append(num)
        
        log_error(f"CORE: launch_random_profiles - 可供随机选择的未运行编号 (在范围 '{numbers_range_str}' 内且快捷方式存在): {available_for_launch}")

        if not available_for_launch:
            log_error("CORE: launch_random_profiles - 没有可供随机启动的编号。")
            return []

        actual_count_to_launch = min(count, len(available_for_launch))
        if actual_count_to_launch <= 0:
            log_error("CORE: launch_random_profiles - 计算出的实际启动数量为0。")
            return []

        selected_numbers = random.sample(available_for_launch, actual_count_to_launch)
        log_error(f"CORE: launch_random_profiles - 随机选中的编号: {selected_numbers}")

        # 使用现有的 open_windows 逻辑来启动这些选中的编号
        # open_windows 接受一个逗号分隔的字符串
        # 我们需要确保 open_windows 能正确处理并返回一些状态，或者我们信任它会启动
        numbers_to_launch_str = ",".join(map(str, selected_numbers))
        
        # 当前 self.open_windows 返回 bool，我们假设它如果返回True就是命令已发送
        # 为了返回成功启动的列表，open_windows 可能需要修改，或者我们在这里做最佳猜测
        success_overall = self.open_windows(numbers_to_launch_str) # 调用现有的

        if success_overall:
            log_error(f"CORE: launch_random_profiles - open_windows 调用成功 for: {numbers_to_launch_str}")
            # 由于 open_windows 不返回具体列表，我们返回我们尝试启动的列表
            return selected_numbers 
        else:
            log_error(f"CORE: launch_random_profiles - open_windows 调用失败 for: {numbers_to_launch_str}")
            return [] # 或者根据 open_windows 的具体行为调整