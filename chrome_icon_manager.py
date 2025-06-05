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
            if hasattr(font, 'getmetrics'):
                ascent, descent = font.getmetrics()
                text_y = (size - ascent) // 2
            else:
                # 手动调整Y位置让数字视觉上居中
                text_y = (size - text_height) // 2 - int(text_height * 0.1)  # 向上调整10%
            
            # 绘制文字阴影（提高可读性）
            shadow_offset = max(2, int(size * 0.012))
            draw.text((text_x + shadow_offset, text_y + shadow_offset), text, 
                     fill=(0, 0, 0, 180), font=font)  # 半透明黑色阴影
            
            # 绘制主要文字
            draw.text((text_x, text_y), text, fill=text_color + (255,), font=font)
            
            # 保存图标文件
            self._save_icon_file(img, icon_path, size)
            
            # 更新缓存
            self.icon_cache[cache_key] = icon_path
            
            self.logger.info(f"成功生成图标: {icon_path}")
            return icon_path
            
        except Exception as e:
            self.last_error = f"生成编号图标失败 (编号: {number}): {str(e)}"
            self.logger.error(self.last_error)
            return None
    
    def _get_best_font(self, font_size: int, text: str) -> object:
        """选择最佳字体 - 使用指定的字体大小
        
        Args:
            font_size: 直接指定的字体大小
            text: 要显示的文本
        """
        # 确保最小字体大小
        font_size = max(font_size, 12)
        
        # 尝试系统字体 - 优先使用粗体字体提高可读性
        font_paths = [
            os.path.join(os.environ.get("WINDIR", ""), "Fonts", "arialbd.ttf"),  # Arial Bold
            os.path.join(os.environ.get("WINDIR", ""), "Fonts", "calibrib.ttf"), # Calibri Bold
            os.path.join(os.environ.get("WINDIR", ""), "Fonts", "arial.ttf"),    # Arial Regular
            os.path.join(os.environ.get("WINDIR", ""), "Fonts", "calibri.ttf"), # Calibri Regular
        ]
        
        for font_path in font_paths:
            if os.path.exists(font_path):
                try:
                    return ImageFont.truetype(font_path, font_size)
                except Exception:
                    continue
        
        # 后备方案
        try:
            return ImageFont.load_default()
        except Exception:
            return None
    
    def _get_text_bbox(self, draw, text: str, font) -> Tuple[int, int, int, int]:
        """获取文字边界框"""
        try:
            if hasattr(font, "getbbox"):
                return font.getbbox(text)
            elif hasattr(draw, "textsize"):
                w, h = draw.textsize(text, font=font)
                return (0, 0, w, h)
            else:
                # 估算大小
                font_size = getattr(font, 'size', 24)
                estimated_width = font_size * len(text) * 0.6
                estimated_height = font_size
                return (0, 0, int(estimated_width), int(estimated_height))
        except Exception:
            # 安全降级
            return (0, 0, 50, 30)
    
    def _save_icon_file(self, img: Image.Image, icon_path: str, size: int):
        """保存图标文件 - 优化任务栏显示效果"""
        try:
            # 生成多种尺寸的图标，确保任务栏显示清晰
            # 特别关注16x16和32x32，这是任务栏最常用的尺寸
            sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
            
            # 只保留不超过当前尺寸的sizes
            valid_sizes = [(w, h) for w, h in sizes if w <= size and h <= size]
            if not valid_sizes:
                valid_sizes = [(size, size)]
            
            # 对于小尺寸图标，进行特殊优化
            optimized_images = []
            for w, h in valid_sizes:
                if w <= 32:
                    # 为小图标重新渲染，使用更粗的线条和更大的文字
                    small_img = self._create_optimized_small_icon(img, w, h, icon_path)
                    if small_img:
                        optimized_images.append((small_img, (w, h)))
                    else:
                        # 降级使用普通缩放
                        resized = img.resize((w, h), Image.LANCZOS)
                        optimized_images.append((resized, (w, h)))
                else:
                    # 大图标直接缩放
                    resized = img.resize((w, h), Image.LANCZOS)
                    optimized_images.append((resized, (w, h)))
            
            # 保存为ICO格式（包含多个尺寸）
            if optimized_images:
                img.save(icon_path, format="ICO", 
                        sizes=[size for _, size in optimized_images])
            else:
                img.save(icon_path, format="ICO", sizes=valid_sizes)
                
        except Exception as e:
            # 降级保存为PNG
            png_path = icon_path.replace('.ico', '.png')
            img.save(png_path, format="PNG")
            # 更新返回路径
            return png_path
    
    def _create_optimized_small_icon(self, base_img: Image.Image, width: int, height: int, icon_path: str) -> Optional[Image.Image]:
        """为小尺寸图标创建优化版本 - 全图标数字设计"""
        try:
            # 从icon_path提取数字
            import re
            match = re.search(r'chrome_(\d+)\.ico', icon_path)
            if not match:
                return None
            
            number = int(match.group(1))
            
            # 重新生成小图标，数字占据整个图标空间
            img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            
            # 绘制圆形图标背景（扩大圆形）
            center_x = width // 2
            center_y = height // 2
            radius = min(center_x, center_y) - 3  # 从2改为3，让小图标圆形也更大
            circle_bounds = (center_x - radius, center_y - radius, center_x + radius, center_y + radius)
            draw.ellipse(circle_bounds, fill=self.default_color + (255,))
            
            text = str(number)
            digit_count = len(text)
            
            # 根据图标尺寸和数字位数计算字体大小，确保在圆形内
            if width <= 16:
                # 超小图标：适中字体，确保在圆形内
                if digit_count == 1:
                    font_size = max(10, int(width * 0.6))  # 从0.9调整到0.6
                else:
                    font_size = max(8, int(width * 0.45))  # 从0.75调整到0.45
            elif width <= 32:
                # 小图标：
                if digit_count == 1:
                    font_size = max(14, int(width * 0.6))  # 从0.85调整到0.6
                else:
                    font_size = max(12, int(width * 0.45)) # 从0.7调整到0.45
            else:
                # 较大图标：
                if digit_count == 1:
                    font_size = max(18, int(width * 0.6))  # 从0.8调整到0.6
                else:
                    font_size = max(16, int(width * 0.45)) # 从0.65调整到0.45
            
            # 获取字体
            try:
                font = ImageFont.truetype(
                    os.path.join(os.environ.get("WINDIR", ""), "Fonts", "arialbd.ttf"), 
                    font_size
                )
            except:
                font = ImageFont.load_default()
            
            # 计算文字位置（在圆形中完美居中）
            bbox = self._get_text_bbox(draw, text, font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            # 水平居中
            text_x = (width - text_width) // 2
            
            # 垂直居中优化 - 解决文字偏下的问题
            if hasattr(font, 'getmetrics'):
                ascent, descent = font.getmetrics()
                text_y = (height - ascent) // 2
            else:
                # 手动调整Y位置让数字视觉上居中
                text_y = (height - text_height) // 2 - int(text_height * 0.1)  # 向上调整10%
            
            # 绘制文字阴影
            shadow_offset = max(1, width // 16)
            draw.text((text_x + shadow_offset, text_y + shadow_offset), text, 
                     fill=(0, 0, 0, 120), font=font)  # 半透明黑色阴影
            
            # 绘制主要文字
            draw.text((text_x, text_y), text, fill=(255, 255, 255, 255), font=font)
            
            return img
            
        except Exception as e:
            return None
    
    def apply_icon_to_window(self, hwnd: int, icon_path: str, retries: int = 3) -> bool:
        """
        应用图标到窗口的设计思路：
        - 多重验证：确认窗口有效性和图标文件存在
        - 分级设置：同时设置大小图标确保兼容性
        - 重试机制：处理系统繁忙或临时失败
        - 系统通知：强制刷新图标缓存
        
        将图标应用到指定窗口
        
        Args:
            hwnd: 窗口句柄
            icon_path: 图标文件路径
            retries: 重试次数
            
        Returns:
            成功返回True，失败返回False
        """
        if not icon_path or not os.path.exists(icon_path):
            self.last_error = f"图标文件不存在: {icon_path}"
            return False
        
        for attempt in range(retries):
            try:
                # 验证窗口有效性
                if not win32gui.IsWindow(hwnd):
                    self.last_error = f"无效窗口句柄: {hwnd}"
                    return False
                
                # 加载图标资源
                large_icon = win32gui.LoadImage(
                    0, icon_path, win32con.IMAGE_ICON, 32, 32, 
                    win32con.LR_LOADFROMFILE
                )
                small_icon = win32gui.LoadImage(
                    0, icon_path, win32con.IMAGE_ICON, 16, 16, 
                    win32con.LR_LOADFROMFILE
                )
                
                if not large_icon or not small_icon:
                    self.last_error = f"加载图标资源失败: {icon_path}"
                    time.sleep(0.1 * (attempt + 1))
                    continue
                
                # 再次验证窗口（防止在加载图标期间窗口被关闭）
                if not win32gui.IsWindow(hwnd):
                    self.last_error = f"窗口在设置图标过程中被关闭: {hwnd}"
                    return False
                
                # 设置窗口图标
                win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, large_icon)
                time.sleep(0.05)  # 短暂延迟确保消息处理
                
                win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, small_icon)
                
                # 强制刷新系统图标缓存
                ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
                
                self.logger.info(f"成功设置窗口图标: HWND={hwnd}, 图标={icon_path}")
                return True
                
            except Exception as e:
                self.last_error = f"设置窗口图标失败 (尝试 {attempt+1}/{retries}): {str(e)}"
                self.logger.warning(self.last_error)
                time.sleep(0.2 * (attempt + 1))  # 递增延迟
        
        return False
    
    def batch_apply_icons_to_windows(self, window_icon_map: Dict[int, int], 
                                   progress_callback=None) -> Dict[int, bool]:
        """
        批量应用图标的策略设计：
        - 并发优化：在生成阶段使用多线程
        - 进度跟踪：提供回调机制便于UI更新
        - 错误隔离：单个失败不影响整体处理
        - 资源管理：及时清理和释放资源
        
        批量为窗口应用编号图标
        
        Args:
            window_icon_map: {窗口句柄: 编号} 映射
            progress_callback: 进度回调函数 callback(current, total, message)
            
        Returns:
            {窗口句柄: 是否成功} 结果映射
        """
        results = {}
        total_count = len(window_icon_map)
        
        # 第一阶段：生成所有需要的图标
        if progress_callback:
            progress_callback(0, total_count, "正在生成图标...")
        
        icon_paths = {}
        unique_numbers = set(window_icon_map.values())
        
        for i, number in enumerate(unique_numbers):
            icon_path = self.generate_numbered_icon(number)
            if icon_path:
                icon_paths[number] = icon_path
            
            if progress_callback:
                progress_callback(i + 1, len(unique_numbers), f"生成图标 {number}")
        
        # 第二阶段：应用图标到窗口
        if progress_callback:
            progress_callback(0, total_count, "正在应用图标...")
        
        processed_count = 0
        for hwnd, number in window_icon_map.items():
            icon_path = icon_paths.get(number)
            if icon_path:
                success = self.apply_icon_to_window(hwnd, icon_path)
                results[hwnd] = success
            else:
                results[hwnd] = False
                self.logger.error(f"窗口 {hwnd} 的图标 {number} 生成失败")
            
            processed_count += 1
            if progress_callback:
                progress_callback(processed_count, total_count, 
                                f"应用图标到窗口 {hwnd}")
        
        success_count = sum(1 for success in results.values() if success)
        self.logger.info(f"批量应用图标完成: {success_count}/{total_count} 成功")
        
        return results
    
    def update_shortcut_icons(self, shortcut_dir: str, number_icon_map: Dict[int, str]) -> Dict[int, bool]:
        """
        更新快捷方式图标的设计理念：
        - 安全检查：验证目录和快捷方式文件存在性
        - COM接口：使用Shell COM组件确保兼容性
        - 批量处理：支持多个快捷方式同时更新
        - 错误恢复：单个失败不影响其他快捷方式
        
        更新快捷方式图标
        
        Args:
            shortcut_dir: 快捷方式目录
            number_icon_map: {编号: 图标路径} 映射
            
        Returns:
            {编号: 是否成功} 结果映射
        """
        if not self.shell:
            self.last_error = "Shell COM组件未初始化"
            return {}
        
        if not os.path.exists(shortcut_dir):
            self.last_error = f"快捷方式目录不存在: {shortcut_dir}"
            return {}
        
        results = {}
        
        for number, icon_path in number_icon_map.items():
            try:
                shortcut_path = os.path.join(shortcut_dir, f"{number}.lnk")
                
                if not os.path.exists(shortcut_path):
                    self.logger.warning(f"快捷方式不存在: {shortcut_path}")
                    results[number] = False
                    continue
                
                if not os.path.exists(icon_path):
                    self.logger.warning(f"图标文件不存在: {icon_path}")
                    results[number] = False
                    continue
                
                # 更新快捷方式图标
                shortcut = self.shell.CreateShortCut(shortcut_path)
                
                # 检查当前图标是否已经是目标图标
                current_icon = getattr(shortcut, 'IconLocation', '')
                if icon_path in current_icon:
                    results[number] = True
                    continue
                
                shortcut.IconLocation = icon_path
                shortcut.save()
                
                results[number] = True
                self.logger.info(f"更新快捷方式图标成功: {shortcut_path} -> {icon_path}")
                
            except Exception as e:
                self.last_error = f"更新快捷方式 {number} 图标失败: {str(e)}"
                self.logger.error(self.last_error)
                results[number] = False
        
        return results
    
    def restore_default_chrome_icons(self, shortcut_dir: str, chrome_exe_path: str) -> bool:
        """
        还原默认Chrome图标的策略：
        - 批量还原：一次性处理所有快捷方式
        - 系统集成：调用系统命令刷新图标缓存
        - 错误处理：记录失败项目但不中断整体流程
        - 验证机制：确认Chrome可执行文件存在
        
        将所有快捷方式图标还原为Chrome默认图标
        
        Args:
            shortcut_dir: 快捷方式目录
            chrome_exe_path: Chrome可执行文件路径
            
        Returns:
            操作是否成功
        """
        if not self.shell:
            self.last_error = "Shell COM组件未初始化"
            return False
        
        if not os.path.exists(shortcut_dir):
            self.last_error = f"快捷方式目录不存在: {shortcut_dir}"
            return False
        
        if not os.path.exists(chrome_exe_path):
            self.last_error = f"Chrome可执行文件不存在: {chrome_exe_path}"
            return False
        
        try:
            success_count = 0
            total_count = 0
            
            # 遍历快捷方式文件
            for filename in os.listdir(shortcut_dir):
                if filename.endswith('.lnk'):
                    total_count += 1
                    shortcut_path = os.path.join(shortcut_dir, filename)
                    
                    try:
                        shortcut = self.shell.CreateShortCut(shortcut_path)
                        shortcut.IconLocation = f"{chrome_exe_path},0"
                        shortcut.save()
                        success_count += 1
                        self.logger.info(f"还原默认图标: {shortcut_path}")
                    except Exception as e:
                        self.logger.error(f"还原快捷方式 {shortcut_path} 失败: {str(e)}")
            
            # 刷新系统图标缓存
            try:
                os.system("ie4uinit.exe -show")
                ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
            except Exception as e:
                self.logger.warning(f"刷新图标缓存失败: {str(e)}")
            
            self.logger.info(f"还原默认图标完成: {success_count}/{total_count} 成功")
            return success_count == total_count
            
        except Exception as e:
            self.last_error = f"还原默认图标失败: {str(e)}"
            self.logger.error(self.last_error)
            return False
    
    def clean_system_icon_cache(self) -> bool:
        """
        清理系统图标缓存的全面策略：
        - 多层清理：图标缓存、缩略图缓存、托盘记忆
        - 服务管理：重启相关系统服务
        - 进程安全：确保Explorer正常重启
        - 用户友好：提供详细的操作反馈
        
        清理系统图标缓存
        
        Returns:
            操作是否成功
        """
        try:
            self.logger.info("开始清理系统图标缓存...")
            
            # 1. 关闭Explorer
            self.logger.info("正在关闭Explorer...")
            subprocess.run(
                ["taskkill", "/f", "/im", "explorer.exe"], 
                creationflags=subprocess.CREATE_NO_WINDOW,
                timeout=10
            )
            time.sleep(1)
            
            # 2. 清理图标缓存文件
            self.logger.info("正在清理图标缓存文件...")
            cache_commands = [
                'attrib -h -s -r "%userprofile%\\AppData\\Local\\IconCache.db"',
                'del /f "%userprofile%\\AppData\\Local\\IconCache.db"',
                'del /f "%userprofile%\\AppData\\Local\\Microsoft\\Windows\\Explorer\\IconCache_*.db"',
            ]
            
            for cmd in cache_commands:
                try:
                    subprocess.run(
                        cmd, shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW,
                        timeout=30
                    )
                except subprocess.TimeoutExpired:
                    self.logger.warning(f"命令执行超时: {cmd}")
                except Exception as e:
                    self.logger.warning(f"执行命令失败: {cmd}, 错误: {str(e)}")
            
            # 3. 清理缩略图缓存
            self.logger.info("正在清理缩略图缓存...")
            thumbnail_commands = [
                'attrib /s /d -h -s -r "%userprofile%\\AppData\\Local\\Microsoft\\Windows\\Explorer\\*"',
                'del /f "%userprofile%\\AppData\\Local\\Microsoft\\Windows\\Explorer\\thumbcache_*.db"',
            ]
            
            for cmd in thumbnail_commands:
                try:
                    subprocess.run(
                        cmd, shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW,
                        timeout=30
                    )
                except Exception as e:
                    self.logger.warning(f"清理缩略图缓存失败: {str(e)}")
            
            # 4. 清理系统托盘图标记忆
            self.logger.info("正在清理系统托盘图标记忆...")
            tray_commands = [
                'reg delete "HKEY_CLASSES_ROOT\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\TrayNotify" /v IconStreams /f',
                'reg delete "HKEY_CLASSES_ROOT\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\TrayNotify" /v PastIconsStream /f',
            ]
            
            for cmd in tray_commands:
                try:
                    subprocess.run(
                        cmd, shell=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW,
                        timeout=10
                    )
                except Exception as e:
                    self.logger.warning(f"清理托盘记忆失败: {str(e)}")
            
            # 5. 刷新系统图标缓存
            self.logger.info("正在刷新系统图标缓存...")
            try:
                subprocess.run(
                    "ie4uinit.exe -show", 
                    shell=True, 
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    timeout=30
                )
            except Exception as e:
                self.logger.warning(f"刷新图标缓存失败: {str(e)}")
            
            # 6. 重启Explorer
            self.logger.info("正在重启Explorer...")
            subprocess.Popen(
                ["start", "explorer.exe"],
                shell=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            time.sleep(2)
            
            # 7. 验证Explorer是否启动成功
            for _ in range(10):
                try:
                    result = subprocess.run(
                        ["tasklist", "/FI", "IMAGENAME eq explorer.exe"], 
                        capture_output=True, 
                        text=True, 
                        creationflags=subprocess.CREATE_NO_WINDOW,
                        timeout=5
                    )
                    if "explorer.exe" in result.stdout.lower():
                        self.logger.info("Explorer重启成功")
                        break
                except Exception:
                    pass
                time.sleep(0.5)
            else:
                # 强制启动Explorer
                subprocess.Popen("explorer.exe")
            
            self.logger.info("系统图标缓存清理完成")
            return True
            
        except Exception as e:
            self.last_error = f"清理系统图标缓存失败: {str(e)}"
            self.logger.error(self.last_error)
            
            # 确保Explorer运行
            try:
                subprocess.Popen("explorer.exe")
            except Exception:
                pass
            
            return False
    
    def find_chrome_windows(self) -> Dict[int, int]:
        """
        查找Chrome窗口的智能策略：
        - 进程关联：通过PID关联窗口和进程
        - 类名过滤：识别Chrome特有的窗口类
        - 可见性检查：只处理用户可见的窗口
        - 数据提取：从命令行参数解析用户配置目录
        
        查找所有Chrome窗口并返回{窗口句柄: 编号}映射
        
        Returns:
            {窗口句柄: 窗口编号} 映射字典
        """
        chrome_windows = {}
        
        def enum_windows_callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            
            try:
                class_name = win32gui.GetClassName(hwnd)
                if "Chrome_WidgetWin" not in class_name:
                    return True
                
                # 获取窗口进程ID
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                
                # 尝试从进程命令行提取编号
                number = self._extract_number_from_process(pid)
                if number:
                    chrome_windows[hwnd] = number
                
            except Exception as e:
                self.logger.warning(f"处理窗口 {hwnd} 失败: {str(e)}")
            
            return True
        
        try:
            win32gui.EnumWindows(enum_windows_callback, None)
            self.logger.info(f"找到 {len(chrome_windows)} 个Chrome窗口")
        except Exception as e:
            self.logger.error(f"枚举Chrome窗口失败: {str(e)}")
        
        return chrome_windows
    
    def _extract_number_from_process(self, pid: int) -> Optional[int]:
        """从进程命令行参数中提取编号"""
        try:
            import psutil
            process = psutil.Process(pid)
            cmdline = ' '.join(process.cmdline())
            
            # 查找 --user-data-dir 参数
            import re
            pattern = r'--user-data-dir=(?:\"([^"]+)\"|([^\s]+))'
            match = re.search(pattern, cmdline)
            
            if match:
                data_dir = match.group(1) or match.group(2)
                # 从路径中提取数字
                path_parts = data_dir.replace('\\', '/').split('/')
                for part in reversed(path_parts):
                    if part.isdigit():
                        return int(part)
            
        except Exception as e:
            self.logger.debug(f"提取进程 {pid} 编号失败: {str(e)}")
        
        return None
    
    def get_last_error(self) -> Optional[str]:
        """获取最后一个错误信息"""
        return self.last_error
    
    def clear_error(self):
        """清除错误状态"""
        self.last_error = None
    
    def get_icon_cache_info(self) -> Dict[str, Union[int, List[str]]]:
        """获取图标缓存信息"""
        return {
            'cache_size': len(self.icon_cache),
            'cached_icons': list(self.icon_cache.keys()),
            'icon_directory': self.icon_dir,
            'directory_files': os.listdir(self.icon_dir) if os.path.exists(self.icon_dir) else []
        }
    
    def cleanup_old_icons(self, keep_numbers: List[int] = None) -> int:
        """
        清理旧的图标文件
        
        Args:
            keep_numbers: 要保留的编号列表，None表示不删除任何文件
            
        Returns:
            删除的文件数量
        """
        if not os.path.exists(self.icon_dir):
            return 0
        
        if keep_numbers is None:
            return 0
        
        deleted_count = 0
        keep_set = set(keep_numbers)
        
        try:
            for filename in os.listdir(self.icon_dir):
                if filename.startswith('chrome_') and (filename.endswith('.ico') or filename.endswith('.png')):
                    # 提取编号
                    name_part = filename.replace('chrome_', '').replace('.ico', '').replace('.png', '')
                    try:
                        number = int(name_part)
                        if number not in keep_set:
                            file_path = os.path.join(self.icon_dir, filename)
                            os.remove(file_path)
                            deleted_count += 1
                            self.logger.info(f"删除旧图标文件: {filename}")
                    except ValueError:
                        continue
            
            # 清理缓存
            self.icon_cache = {k: v for k, v in self.icon_cache.items() if any(str(num) in k for num in keep_numbers)}
            
        except Exception as e:
            self.logger.error(f"清理旧图标文件失败: {str(e)}")
        
        return deleted_count
    
    def __del__(self):
        """析构函数，清理资源"""
        try:
            if self.shell:
                self.shell = None
        except Exception:
            pass


# 便捷函数接口
def create_chrome_icon_manager() -> ChromeIconManager:
    """创建Chrome图标管理器实例"""
    return ChromeIconManager()


def quick_apply_icons_to_chrome_windows(progress_callback=None) -> bool:
    """
    快速应用图标的一站式解决方案：
    - 自动发现：无需手动指定窗口和编号
    - 一键操作：简化复杂的多步骤流程
    - 进度反馈：支持UI集成和用户体验
    - 错误处理：提供友好的错误信息
    
    快速为所有Chrome窗口应用编号图标
    
    Args:
        progress_callback: 进度回调函数
        
    Returns:
        操作是否成功
    """
    try:
        manager = create_chrome_icon_manager()
        
        # 查找Chrome窗口
        if progress_callback:
            progress_callback(0, 100, "正在查找Chrome窗口...")
        
        chrome_windows = manager.find_chrome_windows()
        if not chrome_windows:
            if progress_callback:
                progress_callback(100, 100, "未找到Chrome窗口")
            return False
        
        # 批量应用图标
        results = manager.batch_apply_icons_to_windows(chrome_windows, progress_callback)
        
        success_count = sum(1 for success in results.values() if success)
        total_count = len(results)
        
        if progress_callback:
            progress_callback(100, 100, f"完成: {success_count}/{total_count} 成功")
        
        return success_count > 0
        
    except Exception as e:
        if progress_callback:
            progress_callback(100, 100, f"操作失败: {str(e)}")
        return False


# 示例使用代码
if __name__ == "__main__":
    """
    主程序演示的教学价值：
    - 实用示例：展示各个功能的具体使用方法
    - 错误处理：演示异常情况的处理方式
    - 最佳实践：提供推荐的使用模式
    - 测试验证：验证系统功能的正确性
    """
    
    # 创建图标管理器
    manager = create_chrome_icon_manager()
    
    # 示例1：为单个窗口生成并应用图标
    print("示例1：生成编号图标")
    icon_path = manager.generate_numbered_icon(1)
    if icon_path:
        print(f"成功生成图标: {icon_path}")
    else:
        print(f"生成图标失败: {manager.get_last_error()}")
    
    # 示例2：查找Chrome窗口
    print("\n示例2：查找Chrome窗口")
    chrome_windows = manager.find_chrome_windows()
    print(f"找到 {len(chrome_windows)} 个Chrome窗口")
    for hwnd, number in chrome_windows.items():
        print(f"  窗口 {hwnd} -> 编号 {number}")
    
    # 示例3：批量应用图标
    print("\n示例3：批量应用图标")
    if chrome_windows:
        def progress_callback(current, total, message):
            print(f"进度: {current}/{total} - {message}")
        
        results = manager.batch_apply_icons_to_windows(chrome_windows, progress_callback)
        success_count = sum(1 for success in results.values() if success)
        print(f"批量应用结果: {success_count}/{len(results)} 成功")
    
    # 示例4：使用快速接口
    print("\n示例4：使用快速接口")
    def simple_progress(current, total, message):
        if current % 10 == 0 or current == total:  # 只显示关键进度
            print(f"快速模式: {current}/{total} - {message}")
    
    success = quick_apply_icons_to_chrome_windows(simple_progress)
    print(f"快速应用图标: {'成功' if success else '失败'}")
    
    # 示例5：获取缓存信息
    print("\n示例5：缓存信息")
    cache_info = manager.get_icon_cache_info()
    print(f"缓存大小: {cache_info['cache_size']}")
    print(f"图标目录: {cache_info['icon_directory']}")
    print(f"目录文件数: {len(cache_info['directory_files'])}") 