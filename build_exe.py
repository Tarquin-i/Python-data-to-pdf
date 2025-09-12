"""
PyInstaller打包脚本

用于将data-to-pdf项目打包成可执行文件
"""

import PyInstaller.__main__
import os
import sys
from pathlib import Path

def build_exe():
    """构建可执行文件"""
    
    # 获取项目根目录
    project_root = Path(__file__).parent
    main_script = project_root / "src" / "cli" / "main.py"
    
    # PyInstaller参数
    args = [
        str(main_script),                    # 主脚本路径
        "--onefile",                         # 打包成单个文件
        "--name=data-to-pdf",               # 可执行文件名
        "--console",                         # 控制台应用
        "--clean",                          # 清理临时文件
        f"--distpath={project_root}/dist",  # 输出目录
        f"--workpath={project_root}/build", # 工作目录
        f"--specpath={project_root}",       # spec文件目录
        
        # 包含必要的模块
        "--hidden-import=click",
        "--hidden-import=pathlib",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=reportlab",
        
        # 排除不必要的模块以减小文件大小
        "--exclude-module=tkinter",
        "--exclude-module=matplotlib",
        "--exclude-module=scipy",
        "--exclude-module=numpy.testing",
    ]
    
    print("🚀 开始打包...")
    print(f"主脚本: {main_script}")
    print(f"输出目录: {project_root}/dist")
    
    # 执行打包
    PyInstaller.__main__.run(args)
    
    print("✅ 打包完成!")
    print(f"可执行文件位置: {project_root}/dist/data-to-pdf")

if __name__ == "__main__":
    build_exe()