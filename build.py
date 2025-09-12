"""
构建可执行文件的脚本
支持Windows和macOS
"""

import subprocess
import sys
import os
import platform
from pathlib import Path

def build_executable():
    """构建可执行文件"""
    
    system = platform.system()
    print(f"在{system}系统上构建可执行文件...")
    
    # 确保在项目根目录
    project_root = Path(__file__).parent
    os.chdir(project_root)
    
    # 基础PyInstaller命令
    cmd = [
        "pyinstaller",
        "--onefile",  # 单文件
        "--name=DataToPDF",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl", 
        "--hidden-import=reportlab",
        "--hidden-import=click",
        "src/cli/main.py"
    ]
    
    # 根据系统设置数据路径
    if system == "Windows":
        cmd.insert(-1, "--add-data=src;src")
    else:
        cmd.insert(-1, "--add-data=src:src")
    
    try:
        # 清理旧文件
        if os.path.exists("dist"):
            import shutil
            shutil.rmtree("dist")
        if os.path.exists("build"):
            import shutil
            shutil.rmtree("build")
        
        # 运行PyInstaller
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ 构建成功!")
            
            if system == "Windows":
                exe_file = "dist/DataToPDF.exe"
                print(f"生成文件: {exe_file}")
                print("\n📋 Windows使用:")
                print("  - 双击 DataToPDF.exe 启动")
                print("  - 拖拽Excel文件到exe图标上")
            else:
                exe_file = "dist/DataToPDF"
                print(f"生成文件: {exe_file}")
                print("\n📋 macOS使用:")
                print("  - 双击 DataToPDF 启动")
                print("  - 拖拽Excel文件到图标上")
            
            print("\n🎯 命令行用法:")
            print("  DataToPDF file.xlsx")
            print("  DataToPDF --input file.xlsx --output /path/to/save/")
            
        else:
            print("❌ 构建失败:")
            print(result.stderr)
            
    except Exception as e:
        print(f"构建过程出错: {e}")

if __name__ == "__main__":
    build_executable()