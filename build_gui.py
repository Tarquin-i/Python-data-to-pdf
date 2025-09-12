"""
构建GUI可视化界面应用
支持Windows和macOS
"""

import subprocess
import sys
import os
import platform
from pathlib import Path

def build_gui_app():
    """构建GUI应用"""
    
    system = platform.system()
    print(f"在{system}系统上构建GUI应用...")
    
    # 确保在项目根目录
    project_root = Path(__file__).parent
    os.chdir(project_root)
    
    # GUI应用名称
    if system == "Windows":
        app_name = "DataToPDF_GUI"
    else:
        app_name = "DataToPDF_GUI"
    
    # PyInstaller命令 - 针对不同系统优化
    cmd = [
        "pyinstaller",
        "--onefile",  # 单文件
        "--noconsole" if system == "Windows" else "--console",  # Windows隐藏控制台
        f"--name={app_name}",
    ]
    
    # Windows特殊配置
    if system == "Windows":
        cmd.extend([
            "--icon=icon.ico",  # 如果有图标文件
            "--version-file=version.txt",  # 如果有版本信息
        ])
    
    # 根据系统设置数据路径
    if system == "Windows":
        cmd.append("--add-data=src;src")  # Windows使用;分隔
    else:
        cmd.append("--add-data=src:src")  # macOS/Linux使用:分隔
    
    # 添加隐藏导入
    cmd.extend([
        "--hidden-import=pandas",
        "--hidden-import=openpyxl", 
        "--hidden-import=reportlab",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",
        "src/gui_app.py"
    ])
    
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
            print("✅ GUI应用构建成功!")
            if system == "Windows":
                print(f"生成文件: dist/{app_name}.exe")
                print("📋 Windows使用说明:")
                print(f"  - 双击 {app_name}.exe 启动可视化界面")
                print("  - 支持Windows 10和Windows 11")
            else:
                print(f"生成文件: dist/{app_name}")
                print("📋 macOS使用说明:")
                print(f"  - 双击 {app_name} 文件启动可视化界面")
                print("  - 不生成.app应用包，只生成可执行文件")
            print("\n🎯 操作流程:")
            print("1. 双击运行应用")
            print("2. 点击'选择Excel文件'按钮选择xlsx文件")
            print("3. 点击'生成PDF'按钮，选择保存位置")
            print("4. 自动提取总张数并生成标签PDF")
            print("\n💡 特色功能:")
            print("  - 可视化界面，操作简单")
            print("  - 支持自定义PDF保存位置")
            print("  - 支持文件拖拽到应用图标启动")
        else:
            print("❌ 构建失败:")
            print(result.stderr)
            
    except Exception as e:
        print(f"构建过程出错: {e}")

if __name__ == "__main__":
    build_gui_app()