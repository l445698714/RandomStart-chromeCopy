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
    files_to_remove = ['Chrome_launcher.spec']
    
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
    
    # 判断是否存在图标文件
    icon_param = []
    if os.path.exists("chrome_icon.ico"):
        icon_param = ["--icon=chrome_icon.ico"]
    
    # 构建参数 - 直接调用PyInstaller的Python API
    try:
        # 导入PyInstaller模块
        from PyInstaller.__main__ import run
        
        # 构建参数列表
        args = [
            "--name=Chrome分身启动器",
            "--onefile",  # 打包为单个exe文件
            "--noconsole",  # 不显示控制台窗口
            "--clean",  # 构建前清理
            "--noconfirm",  # 不确认覆盖
            *icon_param,  # 解包可能的图标参数
            "Chrome_launcher.py"
        ]
        
        # 调用PyInstaller
        run(args)
        print("\n应用构建成功!")
        print(f"可执行文件位置: {os.path.join('dist', 'Chrome分身启动器.exe')}")
        return True
    except ImportError:
        print("错误: 找不到PyInstaller模块。请确保已经安装PyInstaller (pip install pyinstaller)")
        return False
    except Exception as e:
        print(f"构建失败: {str(e)}")
        return False

def post_build_operations():
    """构建后的附加操作"""
    print("\n执行构建后操作...")
    
    # 如果需要复制任何附加文件到dist文件夹
    try:
        # 如果存在README.md，复制到dist文件夹
        if os.path.exists("README.md"):
            shutil.copy("README.md", os.path.join("dist", "README.md"))
            print("已复制README.md到dist文件夹")
    except Exception as e:
        print(f"构建后操作失败: {str(e)}")

def main():
    """主函数"""
    print("=" * 50)
    print("Chrome分身启动器 - 构建脚本")
    print("=" * 50)
    
    # 询问是否要清理旧的构建文件
    clean = input("是否清理旧的构建文件和文件夹? (y/n): ").lower()
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
