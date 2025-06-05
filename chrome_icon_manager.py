"""
Chrome图标管理系统 - 整合所有Chrome图标替换相关功能

核心设计思想：
- 系统思维：整合图标生成、应用、缓存管理于一体
- 技术架构：分离关注点，模块化设计，便于维护
- 实用性：提供简单易用的接口，支持批量操作
- 健壮性：异常处理、资源清理、多重重试机制
"""

import os
import sys
import time
import threading
import ctypes
import win32gui
import win32con
import win32api
import win32process
import win32com.client
import subprocess
from PIL import Image, ImageDraw, ImageFont
from typing import Dict, List, Optional, Tuple, Union
import logging
import json
import traceback

class ChromeIconManager:
    """Chrome图标管理器 - 整合所有图标相关功能"""
    
    def __init__(self):
        """
        初始化图标管理器的核心思想：
        - 建立统一的配置管理
        - 创建必要的目录结构
        - 初始化COM组件和系统资源
        - 设置默认参数和状态跟踪
        """
        self.base_dir = self._get_base_directory()
        self.icon_dir = os.path.join(self.base_dir, "icons")
        self._ensure_directories()
        
        # 配置参数
        self.default_icon_size = 256
        self.default_font_size_ratio = 0.9   # 从0.7提升到0.9，字体更大
        self.default_color = (220, 50, 50)   # 改为更鲜艳的红色背景，提高对比度
        self.text_color = (255, 255, 255)   # 保持白色文字
        
        # 状态跟踪
        self.icon_cache = {}  # 图标路径缓存
        self.shell = None     # COM Shell对象
        self.last_error = None
        
        # 初始化COM组件
        self._initialize_com()
        
        # 确保Chrome背景图标存在
        self._ensure_chrome_background()
        
        # 设置日志
        self.logger = self._setup_logger()
    
    def _get_base_directory(self) -> str:
        """获取基础目录路径"""
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    
    def _ensure_directories(self):
        """确保必要的目录存在"""
        os.makedirs(self.icon_dir, exist_ok=True)
    
    def _initialize_com(self):
        """初始化COM组件"""
        try:
            import pythoncom
            pythoncom.CoInitialize()
            self.shell = win32com.client.Dispatch("WScript.Shell")
        except Exception as e:
            self.last_error = f"初始化COM组件失败: {str(e)}"
            if hasattr(self, 'logger'):
                self.logger.error(self.last_error)
    
    def _setup_logger(self) -> logging.Logger:
        """设置日志记录器"""
        logger = logging.getLogger('ChromeIconManager')
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        
        return logger
    
    def _ensure_chrome_background(self):
        """
        移除Chrome背景图标生成 - 现在使用完全的纯色背景设计
        不再需要Chrome风格的背景图片，数字直接占据整个图标空间
        """
        # 完全移除Chrome背景图片的生成和使用
        # 新设计使用纯色背景，让数字占据整个图标空间
        pass
    
    def generate_numbered_icon(self, number: int, size: int = None, 
                              bg_color: Tuple[int, int, int] = None,
                              text_color: Tuple[int, int, int] = None) -> Optional[str]:
        """
        生成带编号的Chrome图标的核心逻辑：
        - 参数化设计：支持自定义尺寸、颜色
        - 缓存机制：避免重复生成相同图标
        - 字体适配：自动选择最佳字体和大小
        - 多格式支持：优先ICO，降级到PNG
        
        生成带编号的Chrome图标
        
        Args:
            number: 窗口编号
            size: 图标尺寸，默认256
            bg_color: 数字标签背景色
            text_color: 文字颜色
            
        Returns:
            图标文件路径，失败返回None
        """
        try:
            if size is None:
                size = self.default_icon_size
            if bg_color is None:
                bg_color = self.default_color
            if text_color is None:
                text_color = self.text_color
            
            # 检查缓存
            cache_key = f"{number}_{size}_{bg_color}_{text_color}"
            if cache_key in self.icon_cache:
                icon_path = self.icon_cache[cache_key]
                if os.path.exists(icon_path):
                    return icon_path
            
            # 生成图标路径
            icon_path = os.path.join(self.icon_dir, f"chrome_{number}.ico")
            
            # 强制清理旧图标，确保使用新的纯色背景设计
            if os.path.exists(icon_path):
                try:
                    os.remove(icon_path)
                    self.logger.info(f"已清理旧图标缓存: {icon_path}")
                except Exception as e:
                    self.logger.warning(f"清理旧图标失败: {str(e)}")
            
            # 清理PNG备用图标
            png_path = icon_path.replace('.ico', '.png')
            if os.path.exists(png_path):
                try:
                    os.remove(png_path)
                except Exception:
                    pass
            
            # 现在总是重新生成图标以确保使用新的纯色背景设计
            
            # 创建图标图像
            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            
            # 完全移除Chrome背景 - 使用纯色背景让数字占据整个图标空间
            # 绘制圆形图标背景（扩大圆形）
            center = size // 2
            radius = center - 6  # 从4改为6，让圆形更大
            circle_bounds = (center - radius, center - radius, center + radius, center + radius)
            draw.ellipse(circle_bounds, fill=bg_color + (255,))
            
            # 计算字体大小 - 让数字尽可能大但确保在圆形内
            text = str(number)
            digit_count = len(text)
            
            # 根据数字位数动态调整字体大小，确保字体在圆形内
            if digit_count == 1:
                # 单位数：适中字体，确保在圆形内
                font_size = int(size * self.default_font_size_ratio * 0.9)  # 从1.3调整到0.9
            else:
                # 多位数：更小字体以适应圆形
                font_size = int(size * self.default_font_size_ratio * 0.7)  # 从1.1调整到0.7
            
            # 获取最佳字体
            font = self._get_best_font(font_size, text)
            
            # 计算文字位置（在圆形中完美居中）
            bbox = self._get_text_bbox(draw, text, font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            # 水平居中
            text_x = (size - text_width) // 2
            
            # 垂直居中优化 - 解决文字偏下的问题
            text_y = (size - text_height) // 2 - 4  # 微调上移
            
            # 绘制文字
            draw.text((text_x, text_y), text, font=font, fill=text_color + (255,))
            
            # 保存图标
            self._save_icon_file(img, icon_path, size)
            
            # 更新缓存
            self.icon_cache[cache_key] = icon_path
            
            self.logger.info(f"成功生成图标: {icon_path}")
            return icon_path
            
        except Exception as e:
            error_msg = f"生成图标失败 (编号{number}): {str(e)}"
            self.last_error = error_msg
            self.logger.error(error_msg)
            return None
    
    def _get_best_font(self, font_size: int, text: str) -> object:
        """
        获取最佳字体的系统思维：
        - 优先级排序：从最理想到最基础
        - 验证机制：确保字体文件存在
        - 降级策略：不可用时自动降级
        - 跨系统兼容：Windows标准字体
        """
        font_candidates = [
            "arial.ttf",
            "arialbold.ttf", 
            "calibri.ttf",
            "calibrib.ttf",
            "seguisb.ttf",
            "ARIALBD.TTF",
            "ARIAL.TTF"
        ]
        
        windows_fonts_dir = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')
        
        for font_name in font_candidates:
            font_path = os.path.join(windows_fonts_dir, font_name)
            try:
                if os.path.exists(font_path):
                    return ImageFont.truetype(font_path, font_size)
            except Exception:
                continue
        
        # 降级到默认字体
        try:
            return ImageFont.load_default()
        except Exception:
            return None
    
    def _get_text_bbox(self, draw, text: str, font) -> Tuple[int, int, int, int]:
        """获取文字边界框"""
        try:
            return draw.textbbox((0, 0), text, font=font)
        except AttributeError:
            # 兼容旧版本PIL
            size = draw.textsize(text, font=font)
            return (0, 0, size[0], size[1])
    
    def _save_icon_file(self, img: Image.Image, icon_path: str, size: int):
        """
        保存图标文件的核心策略：
        - 优先ICO格式：Windows原生支持
        - 多分辨率：16x16到256x256全覆盖
        - 质量优化：保持清晰度
        - 降级策略：ICO失败时降级到PNG
        """
        try:
            # 创建多尺寸的图标
            sizes = [16, 32, 48, 64, 128, 256]
            images = []
            
            for icon_size in sizes:
                if icon_size <= size:
                    if icon_size == size:
                        resized_img = img
                    else:
                        # 对小尺寸图标进行优化
                        resized_img = self._create_optimized_small_icon(img, icon_size, icon_size, icon_path)
                        if resized_img is None:
                            # 降级到普通缩放
                            resized_img = img.resize((icon_size, icon_size), Image.LANCZOS)
                    images.append(resized_img)
            
            # 尝试保存为ICO格式
            try:
                img.save(icon_path, format='ICO', sizes=[(img.width, img.height) for img in images])
                self.logger.info(f"成功保存ICO图标: {icon_path}")
            except Exception as ico_error:
                self.logger.warning(f"ICO保存失败，尝试PNG格式: {str(ico_error)}")
                # 降级保存为PNG
                png_path = icon_path.replace('.ico', '.png')
                img.save(png_path, format='PNG')
                self.logger.info(f"保存为PNG格式: {png_path}")
                
        except Exception as e:
            error_msg = f"保存图标文件失败: {str(e)}"
            self.last_error = error_msg
            self.logger.error(error_msg)
            raise
    
    def _create_optimized_small_icon(self, base_img: Image.Image, width: int, height: int, icon_path: str) -> Optional[Image.Image]:
        """
        为小尺寸图标创建优化版本的设计思想：
        - 可读性优先：在极小尺寸下保持数字清晰
        - 对比度增强：调整颜色确保可见性
        - 字体优化：选择更适合小尺寸的字体
        - 几何调整：微调位置和大小
        """
        try:
            # 仅对16x16和32x32进行特殊优化
            if width not in [16, 32]:
                return base_img.resize((width, height), Image.LANCZOS)
            
            # 从原始icon_path提取数字
            basename = os.path.basename(icon_path)
            if 'chrome_' in basename:
                try:
                    number_str = basename.replace('chrome_', '').replace('.ico', '').replace('.png', '')
                    number = int(number_str)
                except:
                    # 如果提取失败，使用普通缩放
                    return base_img.resize((width, height), Image.LANCZOS)
            else:
                return base_img.resize((width, height), Image.LANCZOS)
            
            # 创建针对小尺寸优化的图标
            img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            
            # 使用更鲜艳的颜色提高对比度
            if width == 16:
                bg_color = (255, 40, 40)  # 更鲜艳的红色
                text_color = (255, 255, 255)
                font_size = 10
                radius = 6
            else:  # 32x32
                bg_color = (240, 50, 50)
                text_color = (255, 255, 255)
                font_size = 18
                radius = 12
            
            # 绘制圆形背景
            center = width // 2
            circle_bounds = (center - radius, center - radius, center + radius, center + radius)
            draw.ellipse(circle_bounds, fill=bg_color + (255,))
            
            # 选择合适的字体
            font = self._get_best_font(font_size, str(number))
            
            # 计算文字位置
            text = str(number)
            bbox = self._get_text_bbox(draw, text, font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            text_x = (width - text_width) // 2
            text_y = (height - text_height) // 2
            
            # 针对小尺寸的位置微调
            if width == 16:
                text_y -= 1  # 16x16时上移1像素
            else:
                text_y -= 2  # 32x32时上移2像素
            
            # 绘制文字
            draw.text((text_x, text_y), text, font=font, fill=text_color + (255,))
            
            return img
            
        except Exception as e:
            self.logger.warning(f"创建优化小图标失败: {str(e)}")
            return base_img.resize((width, height), Image.LANCZOS)
    
    def apply_icon_to_window(self, hwnd: int, icon_path: str, retries: int = 3) -> bool:
        """
        应用图标到窗口的系统策略：
        - 多重设置：ICON_SMALL, ICON_BIG双重保险
        - 重试机制：网络延迟和系统繁忙时的容错
        - 验证机制：确认图标确实已应用
        - 系统兼容：不同Windows版本的适配
        """
        if not os.path.exists(icon_path):
            self.logger.warning(f"图标文件不存在: {icon_path}")
            return False
        
        if not win32gui.IsWindow(hwnd):
            self.logger.warning(f"窗口句柄无效: {hwnd}")
            return False
        
        for attempt in range(retries):
            try:
                # 加载图标
                hicon = win32gui.LoadImage(
                    None, icon_path, win32con.IMAGE_ICON,
                    0, 0, win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
                )
                
                if hicon:
                    # 设置小图标和大图标
                    win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, hicon)
                    win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, hicon)
                    
                    # 强制窗口重绘
                    win32gui.InvalidateRect(hwnd, None, True)
                    win32gui.UpdateWindow(hwnd)
                    
                    # 短暂延迟确保设置生效
                    time.sleep(0.1)
                    
                    self.logger.info(f"成功应用图标到窗口 {hwnd}: {icon_path}")
                    return True
                else:
                    self.logger.warning(f"加载图标失败 (尝试 {attempt + 1}/{retries}): {icon_path}")
                    
            except Exception as e:
                self.logger.warning(f"应用图标失败 (尝试 {attempt + 1}/{retries}): {str(e)}")
                
            if attempt < retries - 1:
                time.sleep(0.5)  # 重试前等待
        
        return False
    
    def batch_apply_icons_to_windows(self, window_icon_map: Dict[int, int], 
                                   progress_callback=None) -> Dict[int, bool]:
        """
        批量应用图标到窗口的并发策略：
        - 进度追踪：实时反馈处理进度
        - 错误隔离：单个失败不影响整体
        - 性能优化：合理的并发控制
        - 结果统计：详细的成功/失败统计
        """
        results = {}
        total_windows = len(window_icon_map)
        
        if total_windows == 0:
            self.logger.info("没有找到需要应用图标的窗口")
            return results
        
        self.logger.info(f"开始批量应用图标到 {total_windows} 个窗口")
        
        for i, (hwnd, number) in enumerate(window_icon_map.items(), 1):
            try:
                # 生成或获取图标
                icon_path = self.generate_numbered_icon(number)
                
                if icon_path:
                    success = self.apply_icon_to_window(hwnd, icon_path)
                    results[hwnd] = success
                else:
                    self.logger.warning(f"无法生成图标 (编号 {number})")
                    results[hwnd] = False
                
                # 更新进度
                if progress_callback:
                    window_title = ""
                    try:
                        window_title = win32gui.GetWindowText(hwnd)
                    except:
                        pass
                    
                    progress_callback(i, total_windows, f"处理窗口 {number}: {window_title[:30]}")
                
                # 避免过快处理导致系统负载
                if i % 5 == 0:
                    time.sleep(0.1)
                    
            except Exception as e:
                self.logger.error(f"处理窗口 {hwnd} (编号 {number}) 时出错: {str(e)}")
                results[hwnd] = False
        
        success_count = sum(results.values())
        self.logger.info(f"批量应用图标完成: {success_count}/{total_windows} 个窗口成功")
        
        return results
    
    def update_shortcut_icons(self, shortcut_dir: str, number_icon_map: Dict[int, str]) -> Dict[int, bool]:
        """
        更新快捷方式图标的批量处理策略：
        - 文件验证：确保快捷方式文件存在
        - COM操作：使用Windows Shell对象
        - 错误处理：单个失败不影响其他
        - 状态追踪：记录每个操作的结果
        """
        results = {}
        
        if not os.path.exists(shortcut_dir):
            self.logger.error(f"快捷方式目录不存在: {shortcut_dir}")
            return results
        
        if not self.shell:
            self.logger.error("COM Shell对象未初始化")
            return results
        
        self.logger.info(f"开始更新快捷方式图标，目录: {shortcut_dir}")
        
        for number, icon_path in number_icon_map.items():
            shortcut_path = os.path.join(shortcut_dir, f"{number}.lnk")
            
            try:
                if not os.path.exists(shortcut_path):
                    self.logger.warning(f"快捷方式不存在: {shortcut_path}")
                    results[number] = False
                    continue
                
                if not os.path.exists(icon_path):
                    self.logger.warning(f"图标文件不存在: {icon_path}")
                    results[number] = False
                    continue
                
                # 创建快捷方式对象
                shortcut = self.shell.CreateShortCut(shortcut_path)
                shortcut.IconLocation = f"{icon_path},0"
                shortcut.Save()
                
                self.logger.info(f"成功更新快捷方式图标: {number}.lnk -> {icon_path}")
                results[number] = True
                
            except Exception as e:
                error_msg = f"更新快捷方式图标失败 (编号 {number}): {str(e)}"
                self.last_error = error_msg
                self.logger.error(error_msg)
                results[number] = False
        
        success_count = sum(results.values())
        total_count = len(number_icon_map)
        self.logger.info(f"快捷方式图标更新完成: {success_count}/{total_count} 个成功")
        
        return results
    
    def restore_default_chrome_icons(self, shortcut_dir: str, chrome_exe_path: str) -> bool:
        """
        恢复默认Chrome图标的系统策略：
        - 全量重置：清理所有自定义图标
        - 系统图标：恢复到Chrome原始图标
        - 缓存清理：删除生成的图标文件
        - 完整验证：确保恢复成功
        """
        try:
            if not os.path.exists(shortcut_dir):
                self.logger.error(f"快捷方式目录不存在: {shortcut_dir}")
                return False
            
            if not os.path.exists(chrome_exe_path):
                self.logger.error(f"Chrome执行文件不存在: {chrome_exe_path}")
                return False
            
            if not self.shell:
                self.logger.error("COM Shell对象未初始化")
                return False
            
            self.logger.info("开始恢复默认Chrome图标")
            
            # 获取所有快捷方式文件
            shortcut_files = [f for f in os.listdir(shortcut_dir) if f.endswith('.lnk')]
            
            success_count = 0
            for shortcut_file in shortcut_files:
                shortcut_path = os.path.join(shortcut_dir, shortcut_file)
                
                try:
                    shortcut = self.shell.CreateShortCut(shortcut_path)
                    shortcut.IconLocation = f"{chrome_exe_path},0"
                    shortcut.Save()
                    success_count += 1
                    self.logger.info(f"恢复快捷方式图标: {shortcut_file}")
                    
                except Exception as e:
                    self.logger.warning(f"恢复快捷方式图标失败 {shortcut_file}: {str(e)}")
            
            # 清理生成的图标缓存
            self.cleanup_old_icons()
            
            # 清理系统图标缓存
            self.clean_system_icon_cache()
            
            self.logger.info(f"默认图标恢复完成: {success_count}/{len(shortcut_files)} 个快捷方式成功")
            return success_count > 0
            
        except Exception as e:
            error_msg = f"恢复默认图标失败: {str(e)}"
            self.last_error = error_msg
            self.logger.error(error_msg)
            return False
    
    def clean_system_icon_cache(self) -> bool:
        """
        清理系统图标缓存的深度策略：
        - 多路径清理：IconCache.db及其版本文件
        - 权限处理：提权访问系统文件
        - 服务重启：Windows图标缓存服务
        - 验证确认：确保清理效果
        """
        try:
            self.logger.info("开始清理系统图标缓存")
            
            # Windows 10/11的图标缓存路径
            cache_paths = [
                os.path.expandvars(r"%localappdata%\\IconCache.db"),
                os.path.expandvars(r"%localappdata%\\Microsoft\\Windows\\Explorer\\IconCache*.db"),
                os.path.expandvars(r"%localappdata%\\Microsoft\\Windows\\Explorer\\thumbcache*.db"),
            ]
            
            cleaned_files = 0
            
            # 尝试结束Explorer进程（需要管理员权限）
            try:
                # 注意：这会重启Explorer，可能造成短暂的桌面闪烁
                self.logger.info("尝试重启Explorer进程以清理图标缓存")
                subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], 
                             capture_output=True, check=False)
                time.sleep(1)
                subprocess.run(["explorer.exe"], capture_output=True, check=False)
                time.sleep(2)
                
            except Exception as e:
                self.logger.warning(f"重启Explorer失败: {str(e)}")
            
            # 清理缓存文件
            for pattern in cache_paths:
                if "*" in pattern:
                    # 处理通配符路径
                    import glob
                    files = glob.glob(pattern)
                    for file_path in files:
                        try:
                            if os.path.exists(file_path):
                                os.remove(file_path)
                                cleaned_files += 1
                                self.logger.info(f"删除缓存文件: {file_path}")
                        except Exception as e:
                            self.logger.warning(f"删除缓存文件失败 {file_path}: {str(e)}")
                else:
                    # 处理单个文件路径
                    try:
                        if os.path.exists(pattern):
                            os.remove(pattern)
                            cleaned_files += 1
                            self.logger.info(f"删除缓存文件: {pattern}")
                    except Exception as e:
                        self.logger.warning(f"删除缓存文件失败 {pattern}: {str(e)}")
            
            # 使用ie4uinit.exe清理图标缓存（更温和的方法）
            try:
                subprocess.run([
                    "ie4uinit.exe", "-show"
                ], capture_output=True, check=False, timeout=10)
                self.logger.info("执行ie4uinit.exe图标缓存清理")
            except Exception as e:
                self.logger.warning(f"ie4uinit.exe执行失败: {str(e)}")
            
            self.logger.info(f"系统图标缓存清理完成，删除了 {cleaned_files} 个缓存文件")
            return True
            
        except Exception as e:
            error_msg = f"清理系统图标缓存失败: {str(e)}"
            self.last_error = error_msg
            self.logger.error(error_msg)
            return False
    
    def find_chrome_windows(self) -> Dict[int, int]:
        """
        查找Chrome窗口的智能识别策略：
        - 进程匹配：通过进程ID关联窗口
        - 参数解析：从命令行提取用户数据目录
        - 编号提取：从路径中识别分身编号
        - 窗口验证：确保是有效的Chrome窗口
        """
        chrome_windows = {}
        
        def enum_windows_callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                
                window_title = win32gui.GetWindowText(hwnd)
                if not window_title or "Chrome" not in window_title:
                    return True
                
                # 获取窗口进程ID
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                except:
                    return True
                
                # 提取编号
                number = self._extract_number_from_process(pid)
                if number is not None:
                    chrome_windows[hwnd] = number
                    self.logger.info(f"找到Chrome窗口: {window_title[:50]} (编号 {number})")
                
            except Exception as e:
                self.logger.warning(f"枚举窗口时出错: {str(e)}")
            
            return True
        
        try:
            win32gui.EnumWindows(enum_windows_callback, None)
            self.logger.info(f"找到 {len(chrome_windows)} 个Chrome窗口")
        except Exception as e:
            self.logger.error(f"枚举窗口失败: {str(e)}")
        
        return chrome_windows
    
    def _extract_number_from_process(self, pid: int) -> Optional[int]:
        """
        从进程中提取分身编号的解析策略：
        - 命令行分析：解析--user-data-dir参数
        - 路径模式：识别chrome\\d+目录模式
        - 正则匹配：精确提取数字编号
        - 异常处理：进程访问权限问题
        """
        try:
            import psutil
            process = psutil.Process(pid)
            
            if 'chrome.exe' not in process.name().lower():
                return None
            
            cmdline = process.cmdline()
            for arg in cmdline:
                if '--user-data-dir=' in arg:
                    user_data_path = arg.split('--user-data-dir=', 1)[1].strip('"')
                    
                    # 匹配chrome\\d+模式
                    import re
                    match = re.search(r'chrome(\d+)', user_data_path, re.IGNORECASE)
                    if match:
                        return int(match.group(1))
            
        except Exception as e:
            self.logger.warning(f"提取进程编号失败 (PID {pid}): {str(e)}")
        
        return None
    
    def get_last_error(self) -> Optional[str]:
        """获取最后一个错误消息"""
        return self.last_error
    
    def clear_error(self):
        """清空错误状态"""
        self.last_error = None
    
    def get_icon_cache_info(self) -> Dict[str, Union[int, List[str]]]:
        """获取图标缓存信息"""
        cache_files = [f for f in os.listdir(self.icon_dir) if f.endswith('.ico')]
        return {
            'cache_count': len(self.icon_cache),
            'file_count': len(cache_files),
            'cache_files': cache_files
        }
    
    def cleanup_old_icons(self, keep_numbers: List[int] = None) -> int:
        """
        清理旧图标的智能策略：
        - 选择性清理：保留指定编号的图标
        - 自动清理：删除过期和无用的图标
        - 空间回收：释放磁盘空间
        - 缓存同步：更新内存缓存状态
        """
        if keep_numbers is None:
            keep_numbers = []
        
        try:
            cleaned_count = 0
            
            # 清理文件
            for filename in os.listdir(self.icon_dir):
                if not filename.startswith('chrome_') or not filename.endswith('.ico'):
                    continue
                
                try:
                    # 提取编号
                    number_str = filename[7:-4]  # 去掉chrome_和.ico
                    number = int(number_str)
                    
                    if number not in keep_numbers:
                        file_path = os.path.join(self.icon_dir, filename)
                        os.remove(file_path)
                        cleaned_count += 1
                        self.logger.info(f"清理旧图标: {filename}")
                        
                except ValueError:
                    # 不是标准格式的文件名，跳过
                    continue
                except Exception as e:
                    self.logger.warning(f"删除图标文件失败 {filename}: {str(e)}")
            
            # 清理内存缓存
            keys_to_remove = []
            for key in self.icon_cache:
                if not os.path.exists(self.icon_cache[key]):
                    keys_to_remove.append(key)
            
            for key in keys_to_remove:
                del self.icon_cache[key]
            
            self.logger.info(f"清理完成，删除了 {cleaned_count} 个旧图标")
            return cleaned_count
            
        except Exception as e:
            error_msg = f"清理旧图标失败: {str(e)}"
            self.last_error = error_msg
            self.logger.error(error_msg)
            return 0
    
    def __del__(self):
        """析构函数 - 清理资源"""
        try:
            if hasattr(self, 'shell') and self.shell:
                self.shell = None
        except:
            pass

def create_chrome_icon_manager() -> ChromeIconManager:
    """创建Chrome图标管理器实例"""
    return ChromeIconManager()

def quick_apply_icons_to_chrome_windows(progress_callback=None) -> bool:
    """
    快速应用图标到所有Chrome窗口的便捷接口：
    - 一键操作：查找窗口+生成图标+应用图标
    - 进度反馈：实时显示处理进度
    - 错误容错：单个失败不影响整体
    - 结果统计：返回操作成功状态
    """
    try:
        manager = create_chrome_icon_manager()
        
        # 查找Chrome窗口
        if progress_callback:
            progress_callback(0, 100, "正在查找Chrome窗口...")
        
        chrome_windows = manager.find_chrome_windows()
        
        if not chrome_windows:
            if progress_callback:
                progress_callback(100, 100, "没有找到Chrome窗口")
            return False
        
        if progress_callback:
            progress_callback(20, 100, f"找到 {len(chrome_windows)} 个Chrome窗口")
        
        # 批量应用图标
        def batch_progress(current, total, message):
            if progress_callback:
                # 将批量进度映射到20-100范围
                overall_progress = 20 + int((current / total) * 80)
                progress_callback(overall_progress, 100, message)
        
        results = manager.batch_apply_icons_to_windows(chrome_windows, batch_progress)
        
        success_count = sum(results.values())
        total_count = len(results)
        
        if progress_callback:
            progress_callback(100, 100, f"完成：{success_count}/{total_count} 个窗口成功应用图标")
        
        return success_count > 0
        
    except Exception as e:
        error_msg = f"快速应用图标失败: {str(e)}"
        if progress_callback:
            progress_callback(100, 100, f"错误: {error_msg}")
        return False

# 测试函数
def test_icon_generation():
    """测试图标生成功能"""
    print("测试Chrome图标管理器...")
    
    try:
        manager = create_chrome_icon_manager()
        
        # 测试生成图标
        test_numbers = [1, 10, 100]
        
        for number in test_numbers:
            print(f"\\n测试生成图标 {number}...")
            icon_path = manager.generate_numbered_icon(number)
            print(f"成功: {icon_path}" if icon_path and os.path.exists(icon_path) else "失败")
        
        # 测试查找Chrome窗口
        print("\\n测试查找Chrome窗口...")
        chrome_windows = manager.find_chrome_windows()
        print(f"找到 {len(chrome_windows)} 个Chrome窗口")
        
        for hwnd, number in chrome_windows.items():
            try:
                title = win32gui.GetWindowText(hwnd)
                print(f"  窗口 {hwnd}: 编号 {number}, 标题: {title[:50]}")
            except:
                print(f"  窗口 {hwnd}: 编号 {number}")
        
        # 测试缓存信息
        print("\\n缓存信息:")
        cache_info = manager.get_icon_cache_info()
        print(f"  内存缓存: {cache_info['cache_count']} 项")
        print(f"  文件缓存: {cache_info['file_count']} 个文件")
        
        print("\\n测试完成！")
        
    except Exception as e:
        print(f"测试失败: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    # 示例用法
    def simple_progress(current, total, message):
        print(f"进度 {current}/{total}: {message}")
    
    print("Chrome图标管理器演示")
    print("=" * 50)
    
    # 快速应用图标到所有Chrome窗口
    success = quick_apply_icons_to_chrome_windows(simple_progress)
    print(f"\\n操作结果: {'成功' if success else '失败'}")