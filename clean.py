#!/usr/bin/env python3
"""
清理项目中的临时文件和编译文件
"""

import os
import shutil
import sys

def clean_project():
    """清理项目中的临时文件和编译文件"""
    print("开始清理项目...")
    
    # 要删除的文件扩展名
    extensions_to_remove = [
        '.pyc',      # Python编译文件
        '.pyo',      # Python优化编译文件
        '.pyd',      # Python动态链接库
        '.~',        # 临时文件
        '.bak',      # 备份文件
        '.swp',      # Vim临时文件
        '.log',      # 日志文件
    ]
    
    # 要删除的目录
    dirs_to_remove = [
        '__pycache__',   # Python缓存目录
        '.pytest_cache',  # pytest缓存
        'build',         # 构建目录
        'dist',          # 分发目录
        '.eggs',         # eggs目录
        '*.egg-info',    # egg信息
        '.coverage',     # 覆盖率报告
    ]
    
    # 获取当前目录
    current_dir = os.getcwd()
    
    # 删除文件
    files_removed = 0
    for root, dirs, files in os.walk(current_dir):
        # 删除指定扩展名的文件
        for file in files:
            for ext in extensions_to_remove:
                if file.endswith(ext):
                    try:
                        os.remove(os.path.join(root, file))
                        print(f"已删除文件: {os.path.join(root, file)}")
                        files_removed += 1
                    except Exception as e:
                        print(f"删除文件失败: {os.path.join(root, file)} - {str(e)}")
    
    # 删除目录
    dirs_removed = 0
    for root, dirs, files in os.walk(current_dir, topdown=False):
        for dir_name in dirs:
            for pattern in dirs_to_remove:
                if dir_name == pattern or (pattern.endswith('*') and dir_name.startswith(pattern[:-1])):
                    try:
                        dir_path = os.path.join(root, dir_name)
                        shutil.rmtree(dir_path)
                        print(f"已删除目录: {dir_path}")
                        dirs_removed += 1
                    except Exception as e:
                        print(f"删除目录失败: {dir_path} - {str(e)}")
    
    print(f"\n清理完成! 已删除 {files_removed} 个文件和 {dirs_removed} 个目录。")

if __name__ == "__main__":
    # 确认是否继续
    if len(sys.argv) > 1 and sys.argv[1] == "--force":
        clean_project()
    else:
        confirm = input("此操作将删除项目中的临时文件和编译文件。确定要继续吗? (y/n): ")
        if confirm.lower() in ['y', 'yes']:
            clean_project()
        else:
            print("操作已取消。") 