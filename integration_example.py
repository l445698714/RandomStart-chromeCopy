"""
Chrome图标管理系统集成示例

这个示例展示了如何将chrome_icon_manager.py集成到现有的Chrome多窗口管理程序中。
主要包含以下集成方式：

1. 基础集成 - 直接使用图标管理器
2. UI集成 - 与现有UI界面集成
3. 高级集成 - 自定义配置和扩展功能
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time

# 导入我们创建的Chrome图标管理系统
from chrome_icon_manager import ChromeIconManager, create_chrome_icon_manager, quick_apply_icons_to_chrome_windows


class ChromeIconIntegrationExample:
    """Chrome图标管理系统集成示例类"""
    
    def __init__(self):
        """初始化集成示例"""
        self.icon_manager = create_chrome_icon_manager()
        self.shortcut_path = ""  # 快捷方式目录
        self.chrome_exe_path = ""  # Chrome可执行文件路径
        
        # 创建简单的测试UI
        self.create_test_ui()
    
    def create_test_ui(self):
        """创建测试用的UI界面"""
        self.root = tk.Tk()
        self.root.title("Chrome图标管理系统集成示例")
        self.root.geometry("600x500")
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text="Chrome图标管理系统集成示例",
            font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # 配置区域
        config_frame = ttk.LabelFrame(main_frame, text="配置设置", padding=10)
        config_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 快捷方式路径
        shortcut_frame = ttk.Frame(config_frame)
        shortcut_frame.pack(fill=tk.X, pady=2)
        ttk.Label(shortcut_frame, text="快捷方式目录:").pack(side=tk.LEFT)
        self.shortcut_var = tk.StringVar()
        ttk.Entry(shortcut_frame, textvariable=self.shortcut_var, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(shortcut_frame, text="浏览", command=self.browse_shortcut_dir).pack(side=tk.LEFT)
        
        # Chrome路径
        chrome_frame = ttk.Frame(config_frame)
        chrome_frame.pack(fill=tk.X, pady=2)
        ttk.Label(chrome_frame, text="Chrome路径:").pack(side=tk.LEFT)
        self.chrome_var = tk.StringVar()
        ttk.Entry(chrome_frame, textvariable=self.chrome_var, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(chrome_frame, text="浏览", command=self.browse_chrome_exe).pack(side=tk.LEFT)
        
        # 基础功能区域
        basic_frame = ttk.LabelFrame(main_frame, text="基础功能", padding=10)
        basic_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 按钮行1
        btn_row1 = ttk.Frame(basic_frame)
        btn_row1.pack(fill=tk.X, pady=2)
        
        ttk.Button(btn_row1, text="查找Chrome窗口", command=self.find_chrome_windows).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row1, text="生成测试图标", command=self.generate_test_icons).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row1, text="一键应用图标", command=self.quick_apply_icons).pack(side=tk.LEFT, padx=5)
        
        # 按钮行2
        btn_row2 = ttk.Frame(basic_frame)
        btn_row2.pack(fill=tk.X, pady=2)
        
        ttk.Button(btn_row2, text="更新快捷方式图标", command=self.update_shortcut_icons).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2, text="还原默认图标", command=self.restore_default_icons).pack(side=tk.LEFT, padx=5)
        
        # 高级功能区域
        advanced_frame = ttk.LabelFrame(main_frame, text="高级功能", padding=10)
        advanced_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 按钮行3
        btn_row3 = ttk.Frame(advanced_frame)
        btn_row3.pack(fill=tk.X, pady=2)
        
        ttk.Button(btn_row3, text="清理系统图标缓存", command=self.clean_system_cache).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row3, text="清理旧图标文件", command=self.cleanup_old_icons).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row3, text="查看缓存信息", command=self.show_cache_info).pack(side=tk.LEFT, padx=5)
        
        # 进度显示区域
        progress_frame = ttk.LabelFrame(main_frame, text="操作进度", padding=10)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_var = tk.StringVar(value="准备就绪")
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_var)
        self.progress_label.pack(anchor=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初始化路径
        self.auto_detect_paths()
    
    def browse_shortcut_dir(self):
        """浏览选择快捷方式目录"""
        from tkinter import filedialog
        directory = filedialog.askdirectory(title="选择快捷方式目录")
        if directory:
            self.shortcut_var.set(directory)
            self.shortcut_path = directory
    
    def browse_chrome_exe(self):
        """浏览选择Chrome可执行文件"""
        from tkinter import filedialog
        filepath = filedialog.askopenfilename(
            title="选择Chrome可执行文件",
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )
        if filepath:
            self.chrome_var.set(filepath)
            self.chrome_exe_path = filepath
    
    def auto_detect_paths(self):
        """自动检测Chrome和快捷方式路径"""
        # 检测Chrome路径
        possible_chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
        ]
        
        for path in possible_chrome_paths:
            if os.path.exists(path):
                self.chrome_var.set(path)
                self.chrome_exe_path = path
                break
        
        # 设置默认快捷方式目录
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "Chrome快捷方式")
        self.shortcut_var.set(desktop_path)
        self.shortcut_path = desktop_path
        
        self.log_message("自动检测到Chrome路径和默认快捷方式目录")
    
    def log_message(self, message):
        """添加日志消息"""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_progress(self, current, total, message):
        """更新进度显示"""
        if total > 0:
            progress_percent = (current / total) * 100
            self.progress_bar['value'] = progress_percent
        
        self.progress_var.set(f"{message} ({current}/{total})")
        self.root.update_idletasks()
    
    def find_chrome_windows(self):
        """查找Chrome窗口"""
        self.log_message("开始查找Chrome窗口...")
        
        def find_task():
            try:
                chrome_windows = self.icon_manager.find_chrome_windows()
                
                self.root.after(0, lambda: self.log_message(f"找到 {len(chrome_windows)} 个Chrome窗口"))
                
                if chrome_windows:
                    details = []
                    for hwnd, number in chrome_windows.items():
                        details.append(f"  窗口句柄 {hwnd} -> 编号 {number}")
                    
                    for detail in details:
                        self.root.after(0, lambda msg=detail: self.log_message(msg))
                else:
                    self.root.after(0, lambda: self.log_message("未找到Chrome窗口，请确保Chrome正在运行"))
                
            except Exception as e:
                error_msg = f"查找Chrome窗口失败: {str(e)}"
                self.root.after(0, lambda: self.log_message(error_msg))
        
        threading.Thread(target=find_task, daemon=True).start()
    
    def generate_test_icons(self):
        """生成测试图标"""
        self.log_message("开始生成测试图标...")
        
        def generate_task():
            try:
                test_numbers = [1, 2, 3, 4, 5]
                
                for i, number in enumerate(test_numbers):
                    icon_path = self.icon_manager.generate_numbered_icon(number)
                    
                    if icon_path:
                        self.root.after(0, lambda num=number, path=icon_path: 
                                      self.log_message(f"生成图标 {num}: {path}"))
                    else:
                        error = self.icon_manager.get_last_error()
                        self.root.after(0, lambda num=number, err=error: 
                                      self.log_message(f"生成图标 {num} 失败: {err}"))
                    
                    self.root.after(0, lambda current=i+1, total=len(test_numbers): 
                                  self.update_progress(current, total, "生成测试图标"))
                
                self.root.after(0, lambda: self.log_message("测试图标生成完成"))
                
            except Exception as e:
                error_msg = f"生成测试图标失败: {str(e)}"
                self.root.after(0, lambda: self.log_message(error_msg))
        
        threading.Thread(target=generate_task, daemon=True).start()
    
    def quick_apply_icons(self):
        """一键应用图标到所有Chrome窗口"""
        self.log_message("开始一键应用图标...")
        
        def apply_task():
            try:
                def progress_callback(current, total, message):
                    self.root.after(0, lambda: self.update_progress(current, total, message))
                    self.root.after(0, lambda: self.log_message(f"进度: {message}"))
                
                success = quick_apply_icons_to_chrome_windows(progress_callback)
                
                if success:
                    self.root.after(0, lambda: self.log_message("一键应用图标完成！"))
                else:
                    self.root.after(0, lambda: self.log_message("一键应用图标失败，请检查是否有Chrome窗口运行"))
                
            except Exception as e:
                error_msg = f"一键应用图标失败: {str(e)}"
                self.root.after(0, lambda: self.log_message(error_msg))
        
        threading.Thread(target=apply_task, daemon=True).start()
    
    def update_shortcut_icons(self):
        """更新快捷方式图标"""
        if not self.shortcut_path:
            messagebox.showwarning("警告", "请先设置快捷方式目录！")
            return
        
        self.log_message("开始更新快捷方式图标...")
        
        def update_task():
            try:
                # 生成一些测试图标
                test_numbers = [1, 2, 3, 4, 5]
                number_icon_map = {}
                
                for number in test_numbers:
                    icon_path = self.icon_manager.generate_numbered_icon(number)
                    if icon_path:
                        number_icon_map[number] = icon_path
                
                if number_icon_map:
                    results = self.icon_manager.update_shortcut_icons(self.shortcut_path, number_icon_map)
                    
                    success_count = sum(1 for success in results.values() if success)
                    total_count = len(results)
                    
                    self.root.after(0, lambda: self.log_message(f"更新快捷方式图标完成: {success_count}/{total_count} 成功"))
                else:
                    self.root.after(0, lambda: self.log_message("没有可用的图标文件"))
                
            except Exception as e:
                error_msg = f"更新快捷方式图标失败: {str(e)}"
                self.root.after(0, lambda: self.log_message(error_msg))
        
        threading.Thread(target=update_task, daemon=True).start()
    
    def restore_default_icons(self):
        """还原默认Chrome图标"""
        if not self.shortcut_path or not self.chrome_exe_path:
            messagebox.showwarning("警告", "请先设置快捷方式目录和Chrome路径！")
            return
        
        result = messagebox.askyesno("确认", "确定要还原所有快捷方式为默认Chrome图标吗？")
        if not result:
            return
        
        self.log_message("开始还原默认图标...")
        
        def restore_task():
            try:
                success = self.icon_manager.restore_default_chrome_icons(
                    self.shortcut_path, self.chrome_exe_path
                )
                
                if success:
                    self.root.after(0, lambda: self.log_message("还原默认图标完成！"))
                else:
                    error = self.icon_manager.get_last_error()
                    self.root.after(0, lambda: self.log_message(f"还原默认图标失败: {error}"))
                
            except Exception as e:
                error_msg = f"还原默认图标失败: {str(e)}"
                self.root.after(0, lambda: self.log_message(error_msg))
        
        threading.Thread(target=restore_task, daemon=True).start()
    
    def clean_system_cache(self):
        """清理系统图标缓存"""
        result = messagebox.askyesno(
            "确认清理", 
            "清理系统图标缓存将：\n"
            "1. 关闭资源管理器\n"
            "2. 删除图标缓存文件\n"
            "3. 重启资源管理器\n\n"
            "这可能导致桌面短暂消失，确定继续吗？"
        )
        
        if not result:
            return
        
        self.log_message("开始清理系统图标缓存...")
        
        def clean_task():
            try:
                success = self.icon_manager.clean_system_icon_cache()
                
                if success:
                    self.root.after(0, lambda: self.log_message("系统图标缓存清理完成！"))
                else:
                    error = self.icon_manager.get_last_error()
                    self.root.after(0, lambda: self.log_message(f"清理系统缓存失败: {error}"))
                
            except Exception as e:
                error_msg = f"清理系统缓存失败: {str(e)}"
                self.root.after(0, lambda: self.log_message(error_msg))
        
        threading.Thread(target=clean_task, daemon=True).start()
    
    def cleanup_old_icons(self):
        """清理旧的图标文件"""
        # 保留编号1-10的图标
        keep_numbers = list(range(1, 11))
        
        self.log_message("开始清理旧图标文件...")
        
        def cleanup_task():
            try:
                deleted_count = self.icon_manager.cleanup_old_icons(keep_numbers)
                self.root.after(0, lambda: self.log_message(f"清理完成，删除了 {deleted_count} 个旧图标文件"))
                
            except Exception as e:
                error_msg = f"清理旧图标失败: {str(e)}"
                self.root.after(0, lambda: self.log_message(error_msg))
        
        threading.Thread(target=cleanup_task, daemon=True).start()
    
    def show_cache_info(self):
        """显示缓存信息"""
        try:
            cache_info = self.icon_manager.get_icon_cache_info()
            
            info_text = f"""
缓存信息:
- 内存缓存大小: {cache_info['cache_size']} 项
- 图标目录: {cache_info['icon_directory']}
- 目录文件数: {len(cache_info['directory_files'])} 个

目录文件列表:
{chr(10).join(f"  {f}" for f in cache_info['directory_files'][:10])}
{'  ...' if len(cache_info['directory_files']) > 10 else ''}
            """
            
            messagebox.showinfo("缓存信息", info_text)
            
        except Exception as e:
            messagebox.showerror("错误", f"获取缓存信息失败: {str(e)}")
    
    def run(self):
        """运行示例程序"""
        self.root.mainloop()


# 集成到现有项目的简化示例
class SimpleIconIntegration:
    """简化的图标集成示例，适合集成到现有项目"""
    
    def __init__(self, existing_ui_manager=None):
        """
        初始化简化集成
        
        Args:
            existing_ui_manager: 现有的UI管理器实例
        """
        self.ui_manager = existing_ui_manager
        self.icon_manager = create_chrome_icon_manager()
    
    def apply_icons_to_imported_windows(self, windows_data):
        """
        为导入的窗口应用图标
        
        Args:
            windows_data: 窗口数据列表，格式如 [{'hwnd': 12345, 'number': 1}, ...]
        """
        try:
            # 构建窗口到编号的映射
            window_icon_map = {}
            for window in windows_data:
                hwnd = window.get('hwnd')
                number = window.get('number')
                if hwnd and number:
                    window_icon_map[hwnd] = number
            
            if not window_icon_map:
                return False
            
            # 定义进度回调
            def progress_callback(current, total, message):
                if hasattr(self.ui_manager, 'show_status'):
                    self.ui_manager.show_status(f"{message} ({current}/{total})")
            
            # 批量应用图标
            results = self.icon_manager.batch_apply_icons_to_windows(
                window_icon_map, progress_callback
            )
            
            # 统计结果
            success_count = sum(1 for success in results.values() if success)
            
            if hasattr(self.ui_manager, 'show_notification'):
                self.ui_manager.show_notification(
                    "图标应用完成", 
                    f"成功为 {success_count}/{len(results)} 个窗口应用了图标"
                )
            
            return success_count > 0
            
        except Exception as e:
            if hasattr(self.ui_manager, 'show_error'):
                self.ui_manager.show_error(f"应用图标失败: {str(e)}")
            return False
    
    def update_shortcut_icons_for_numbers(self, shortcut_dir, numbers):
        """
        为指定编号更新快捷方式图标
        
        Args:
            shortcut_dir: 快捷方式目录
            numbers: 编号列表
        """
        try:
            # 生成图标并构建映射
            number_icon_map = {}
            for number in numbers:
                icon_path = self.icon_manager.generate_numbered_icon(number)
                if icon_path:
                    number_icon_map[number] = icon_path
            
            if not number_icon_map:
                return False
            
            # 更新快捷方式图标
            results = self.icon_manager.update_shortcut_icons(shortcut_dir, number_icon_map)
            
            success_count = sum(1 for success in results.values() if success)
            
            if hasattr(self.ui_manager, 'show_notification'):
                self.ui_manager.show_notification(
                    "快捷方式更新完成", 
                    f"成功更新 {success_count}/{len(results)} 个快捷方式图标"
                )
            
            return success_count > 0
            
        except Exception as e:
            if hasattr(self.ui_manager, 'show_error'):
                self.ui_manager.show_error(f"更新快捷方式图标失败: {str(e)}")
            return False


# 示例：如何在现有项目中集成
def integrate_with_existing_chrome_manager(chrome_manager_instance):
    """
    示例：如何与现有的ChromeManager集成
    
    Args:
        chrome_manager_instance: 现有的ChromeManager实例
    """
    
    # 创建图标集成实例
    icon_integration = SimpleIconIntegration(chrome_manager_instance)
    
    # 假设现有的import_windows方法返回窗口数据
    def enhanced_import_windows():
        """增强的导入窗口方法，包含图标应用"""
        
        # 调用原有的导入窗口逻辑
        windows = chrome_manager_instance.import_windows()
        
        if windows:
            # 应用图标到导入的窗口
            icon_integration.apply_icons_to_imported_windows(windows)
        
        return windows
    
    # 替换原有方法
    chrome_manager_instance.enhanced_import_windows = enhanced_import_windows
    
    return icon_integration


if __name__ == "__main__":
    """运行集成示例"""
    
    print("Chrome图标管理系统集成示例")
    print("=" * 40)
    
    # 运行完整的UI示例
    try:
        app = ChromeIconIntegrationExample()
        app.run()
    except Exception as e:
        print(f"运行示例失败: {str(e)}")
        
        # 如果UI示例失败，运行命令行示例
        print("\n切换到命令行模式...")
        
        manager = create_chrome_icon_manager()
        
        print("1. 查找Chrome窗口...")
        windows = manager.find_chrome_windows()
        print(f"找到 {len(windows)} 个Chrome窗口")
        
        print("2. 生成测试图标...")
        for i in range(1, 4):
            icon_path = manager.generate_numbered_icon(i)
            print(f"图标 {i}: {icon_path or '生成失败'}")
        
        print("3. 获取缓存信息...")
        cache_info = manager.get_icon_cache_info()
        print(f"缓存大小: {cache_info['cache_size']}")
        print(f"图标目录: {cache_info['icon_directory']}")
        
        print("\n示例运行完成！") 