#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Chrome分身启动器打包脚本
使用PyInstaller打包应用为独立的exe文件
"""

import os
import sys
import shutil

def clean_build_folders():
    """清理旧的构建文件夹"""
    print("清理旧的构建文件和文件夹...")
    folders_to_remove = ['build', 'dist', '__pycache__']
    files_to_remove = ['Chrome_launcher.spec', 'Chrome分身启动器.spec', 'ChromeLauncher.spec']
    
    for folder in folders_to_remove:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"已删除文件夹: {folder}")
            
    for file in files_to_remove:
        if os.path.exists(file):
            os.remove(file)
            print(f"已删除文件: {file}")

def build_application():
    """使用PyInstaller构建应用"""
    print("开始构建应用...")
    
    # 检查必要文件是否存在
    required_files = [
        "Chrome_launcher.py",
        "chrome_icon_manager.py", 
        "utils.py",
        "ico.ico"
    ]
    
    for file in required_files:
        if not os.path.exists(file):
            print(f"错误: 找不到必需文件 {file}")
            return False
    
    # 图标文件参数
    icon_param = []
    if os.path.exists("ico.ico"):
        icon_param = ["--icon=ico.ico"]
        print("找到应用图标文件: ico.ico")
    
    # 构建参数 - 直接调用PyInstaller的Python API
    try:
        # 导入PyInstaller模块
        from PyInstaller.__main__ import run
        
        # 构建参数列表
        args = [
            "--name=Chrome分身启动器",
            "--onefile",  # 打包为单个exe文件
            "--windowed",  # 不显示控制台窗口 (替代--noconsole)
            "--clean",  # 构建前清理
            "--noconfirm",  # 不确认覆盖
            *icon_param,  # 解包可能的图标参数
            
            # 添加数据文件 - 图标资源文件夹
            "--add-data=icons;icons",
            "--add-data=ico.ico;.",
            
            # 隐藏导入 - 确保PyQt5模块被正确打包
            "--hidden-import=PyQt5.QtCore",
            "--hidden-import=PyQt5.QtGui", 
            "--hidden-import=PyQt5.QtWidgets",
            "--hidden-import=win32com.client",
            "--hidden-import=win32gui",
            "--hidden-import=win32process",
            "--hidden-import=win32con",
            "--hidden-import=win32api",
            "--hidden-import=psutil",
            
            # 包含本地模块
            "--hidden-import=chrome_icon_manager",
            "--hidden-import=utils",
            
            # 排除不需要的模块以减小文件大小
            "--exclude-module=matplotlib",
            "--exclude-module=numpy",
            "--exclude-module=pandas",
            "--exclude-module=scipy",
            
            "Chrome_launcher.py"
        ]
        
        print("PyInstaller 参数:")
        for arg in args:
            print(f"  {arg}")
        print()
        
        # 调用PyInstaller
        run(args)
        print("\n应用构建成功!")
        print(f"可执行文件位置: {os.path.join('dist', 'Chrome分身启动器.exe')}")
        return True
    except ImportError:
        print("错误: 找不到PyInstaller模块。请确保已经安装PyInstaller")
        print("安装命令: pip install pyinstaller")
        return False
    except Exception as e:
        print(f"构建失败: {str(e)}")
        return False

def post_build_operations():
    """构建后的附加操作"""
    print("\n执行构建后操作...")
    
    try:
        # 复制文档文件到dist文件夹
        docs_to_copy = ["README.md", "README.txt", "version.txt"]
        
        for doc in docs_to_copy:
            if os.path.exists(doc):
                shutil.copy(doc, os.path.join("dist", doc))
                print(f"已复制 {doc} 到dist文件夹")
        
        # 检查最终exe文件
        exe_path = os.path.join("dist", "Chrome分身启动器.exe")
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"生成的exe文件大小: {size_mb:.1f} MB")
        
    except Exception as e:
        print(f"构建后操作失败: {str(e)}")

def check_dependencies():
    """检查构建依赖"""
    print("检查构建依赖...")
    
    try:
        import PyInstaller
        print(f"✅ PyInstaller 版本: {PyInstaller.__version__}")
    except ImportError:
        print("❌ PyInstaller 未安装")
        return False
    
    try:
        import PyQt5
        print(f"✅ PyQt5 已安装")
    except ImportError:
        print("❌ PyQt5 未安装")
        return False
    
    try:
        import psutil
        print(f"✅ psutil 已安装")
    except ImportError:
        print("❌ psutil 未安装")
        return False
    
    try:
        import win32com.client
        print(f"✅ pywin32 已安装")
    except ImportError:
        print("❌ pywin32 未安装")
        return False
    
    return True

def main():
    """主函数"""
    print("=" * 50)
    print("Chrome分身启动器 - 构建脚本")
    print("=" * 50)
    
    # 检查依赖
    if not check_dependencies():
        print("\n请先安装缺失的依赖包:")
        print("pip install -r requirements.txt")
        print("pip install pyinstaller")
        print("pip install pywin32")
        return
    
    # 询问是否要清理旧的构建文件
    clean = input("\n是否清理旧的构建文件和文件夹? (y/n): ").lower()
    if clean == 'y':
        clean_build_folders()
    
    # 构建应用
    if build_application():
        post_build_operations()
        
        # 询问是否要运行应用
        run_app = input("\n是否要立即运行应用? (y/n): ").lower()
        if run_app == 'y':
            exe_path = os.path.join("dist", "Chrome分身启动器.exe")
            if os.path.exists(exe_path):
                print("启动应用...")
                os.startfile(exe_path)
            else:
                print(f"错误: 找不到执行文件 {exe_path}")
    
    print("\n构建过程完成!")

if __name__ == "__main__":
    main() 