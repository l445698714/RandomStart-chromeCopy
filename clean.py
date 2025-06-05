#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清理脚本 - 清理生成的图标缓存和临时文件
作者：l445698714
"""

import os
import shutil

def clean_icons():
    """清理生成的图标文件"""
    icons_dir = "icons"
    
    if not os.path.exists(icons_dir):
        print("icons目录不存在")
        return
    
    # 只清理动态生成的图标（编号>45的图标）
    cleaned_count = 0
    for filename in os.listdir(icons_dir):
        if filename.startswith("chrome_") and filename.endswith(".ico"):
            try:
                number = int(filename[7:-4])  # 提取编号
                if number > 45:  # 只清理动态生成的图标
                    file_path = os.path.join(icons_dir, filename)
                    os.remove(file_path)
                    print(f"已删除动态图标: {filename}")
                    cleaned_count += 1
            except ValueError:
                continue
    
    print(f"共清理了 {cleaned_count} 个动态生成的图标")

def clean_cache():
    """清理Python缓存"""
    cache_dirs = ['__pycache__', 'build', 'dist']
    
    for cache_dir in cache_dirs:
        if os.path.exists(cache_dir):
            shutil.rmtree(cache_dir)
            print(f"已删除缓存目录: {cache_dir}")
    
    # 清理.pyc文件
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.endswith('.pyc'):
                file_path = os.path.join(root, file)
                os.remove(file_path)
                print(f"已删除缓存文件: {file_path}")

def clean_spec_files():
    """清理PyInstaller规范文件"""
    spec_files = ['Chrome_launcher.spec', 'Chrome分身启动器.spec', 'ChromeLauncher.spec']
    
    for spec_file in spec_files:
        if os.path.exists(spec_file):
            os.remove(spec_file)
            print(f"已删除规范文件: {spec_file}")

def main():
    print("Chrome分身启动器 - 清理脚本")
    print("=" * 40)
    
    choice = input("请选择清理类型:\n1. 清理动态图标\n2. 清理Python缓存\n3. 清理规范文件\n4. 全部清理\n请输入选择 (1-4): ")
    
    if choice == '1':
        clean_icons()
    elif choice == '2':
        clean_cache()
    elif choice == '3':
        clean_spec_files()
    elif choice == '4':
        clean_icons()
        clean_cache()
        clean_spec_files()
    else:
        print("无效选择")
    
    print("\n清理完成！")

if __name__ == "__main__":
    main()
