#!/usr/bin/env python3
"""
Data to PDF Print - 统一构建脚本
支持 macOS 和 Windows 平台的 GUI 应用构建

使用方法:
    python scripts/build.py              # 自动检测当前系统
    python scripts/build.py --platform macOS     # 强制构建 macOS 版本
    python scripts/build.py --platform Windows   # 强制构建 Windows 版本
    python scripts/build.py --help               # 显示帮助信息
"""

import subprocess
import sys
import os
import platform
import argparse
import shutil
from pathlib import Path


class AppBuilder:
    """应用程序构建器"""
    
    def __init__(self):
        # 确保在项目根目录
        self.project_root = Path(__file__).parent.parent
        os.chdir(self.project_root)
        
        # 配置
        self.app_name = "DataToPDF_GUI"
        self.entry_point = "src/gui_app.py"
        
    def check_requirements(self, target_platform):
        """检查构建要求"""
        print(f"检查 {target_platform} 构建要求...")
        
        # 检查入口文件
        if not os.path.exists(self.entry_point):
            print(f"ERROR: 找不到入口文件 {self.entry_point}")
            return False
            
        # 检查 Windows 特殊要求
        if target_platform == "Windows":
            font_file = "src/fonts/msyh.ttf"
            if os.path.exists(font_file):
                file_size = os.path.getsize(font_file) / (1024 * 1024)  # MB
                print(f"[OK] 字体文件: {font_file} ({file_size:.1f}MB)")
            else:
                print(f"[WARNING] 字体文件未找到: {font_file}")
                print("这可能导致 Windows 版本中文显示问题")
                
        return True
        
    def clean_build_files(self):
        """清理旧的构建文件"""
        print("清理旧的构建文件...")
        
        for directory in ["dist", "build"]:
            if os.path.exists(directory):
                shutil.rmtree(directory)
                print(f"  已删除: {directory}/")
                
        # 清理自动生成的 spec 文件
        for spec_file in Path(".").glob("*.spec"):
            spec_file.unlink()
            print(f"  已删除: {spec_file}")
    
    def build_app(self, target_platform):
        """构建应用程序"""
        print(f"开始构建 {target_platform} 版本...")
        
        # 基础 PyInstaller 命令
        cmd = [
            "pyinstaller",
            "--onefile",
            f"--name={self.app_name}",
            "--clean",
            "--noconfirm",
        ]
        
        # 平台特定配置
        if target_platform == "Windows":
            cmd.extend([
                "--windowed",  # Windows 隐藏控制台
                "--add-data=src;src",  # Windows 路径分隔符
            ])
            
            # 添加字体文件（如果存在）
            font_file = "src/fonts/msyh.ttf"
            if os.path.exists(font_file):
                cmd.extend(["--add-data", f"{font_file};fonts"])
                
        else:  # macOS/Linux
            cmd.extend([
                "--console",  # macOS 显示控制台
                "--add-data=src:src",  # Unix 路径分隔符
            ])
        
        # 添加隐藏导入
        hidden_imports = [
            "pandas", "openpyxl", "reportlab",
            "tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox"
        ]
        for import_name in hidden_imports:
            cmd.extend(["--hidden-import", import_name])
            
        # 添加入口点
        cmd.append(self.entry_point)
        
        # 执行构建
        try:
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                self._build_success(target_platform)
                return True
            else:
                self._build_failed(result)
                return False
                
        except FileNotFoundError:
            print("ERROR: PyInstaller 未找到，请先安装:")
            print("  pip install pyinstaller")
            return False
        except Exception as e:
            print(f"ERROR: 构建过程出错: {e}")
            return False
    
    def _build_success(self, target_platform):
        """构建成功处理"""
        print("✅ 构建成功!")
        
        if target_platform == "Windows":
            exe_path = f"dist/{self.app_name}.exe"
            print(f"生成文件: {exe_path}")
            
            # 检查文件大小
            if os.path.exists(exe_path):
                file_size_mb = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"文件大小: {file_size_mb:.1f}MB")
                
            print("\n📋 Windows 使用说明:")
            print(f"  - 双击 {self.app_name}.exe 启动应用")
            print("  - 支持 Windows 7/8/10/11")
            print("  - 无需安装 Python 环境")
            
        else:  # macOS
            app_path = f"dist/{self.app_name}"
            print(f"生成文件: {app_path}")
            
            print("\n📋 macOS 使用说明:")
            print(f"  - 双击 {self.app_name} 启动应用")
            print("  - 适用于 M1/M2 Mac")
            
        print("\n🎯 操作流程:")
        print("1. 启动应用程序")
        print("2. 选择 Excel 文件")
        print("3. 选择模板类型和参数")
        print("4. 生成 PDF 标签")
        
    def _build_failed(self, result):
        """构建失败处理"""
        print("❌ 构建失败:")
        print("STDOUT:", result.stdout)
        print("STDERR:", result.stderr)
        
        # 常见错误解决方案
        if "No module named" in result.stderr:
            print("\n💡 解决方案:")
            print("请安装依赖: pip install -r requirements.txt")
        elif "command not found" in result.stderr:
            print("\n💡 解决方案:")
            print("请安装 PyInstaller: pip install pyinstaller")
    
    def create_distribution_package(self, target_platform):
        """创建分发包"""
        if target_platform != "Windows":
            return
            
        exe_path = f"dist/{self.app_name}.exe"
        if not os.path.exists(exe_path):
            return
            
        print("\n创建 Windows 分发包...")
        
        # 创建分发目录
        dist_dir = Path("windows_distribution")
        if dist_dir.exists():
            shutil.rmtree(dist_dir)
        dist_dir.mkdir()
        
        # 复制可执行文件
        shutil.copy2(exe_path, dist_dir / f"{self.app_name}.exe")
        
        # 创建使用说明
        readme_content = f"""# Data to PDF Print - Windows 版本

## 使用说明
1. 双击运行 {self.app_name}.exe
2. 选择 Excel 文件 (.xlsx 格式)
3. 选择模板类型和参数
4. 点击生成 PDF 按钮
5. 选择保存位置

## 系统要求
- Windows 7/8/10/11 (64位)
- 无需安装 Python 或其他依赖

## 注意事项
- 首次运行可能被 Windows Defender 检测，选择"允许"
- 建议将程序放在非系统盘（如 D: 盘）
- 支持的 Excel 格式: .xlsx, .xls

## 问题反馈
如遇到问题请联系开发者
"""
        
        with open(dist_dir / "使用说明.txt", "w", encoding="utf-8") as f:
            f.write(readme_content)
        
        print(f"✅ 分发包已创建: {dist_dir.absolute()}")
        print("包含文件:")
        for file in dist_dir.iterdir():
            print(f"  - {file.name}")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="Data to PDF Print 构建脚本",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python scripts/build.py                    # 自动检测系统
  python scripts/build.py --platform macOS   # 构建 macOS 版本
  python scripts/build.py --platform Windows # 构建 Windows 版本
        """
    )
    
    parser.add_argument(
        "--platform",
        choices=["macOS", "Windows"],
        help="目标平台 (默认: 自动检测)"
    )
    
    args = parser.parse_args()
    
    # 确定目标平台
    if args.platform:
        target_platform = args.platform
    else:
        current_system = platform.system()
        if current_system == "Darwin":
            target_platform = "macOS"
        elif current_system == "Windows":
            target_platform = "Windows"
        else:
            print(f"ERROR: 不支持的系统: {current_system}")
            print("请使用 --platform 参数指定目标平台")
            sys.exit(1)
    
    print(f"目标平台: {target_platform}")
    print(f"当前系统: {platform.system()}")
    
    if target_platform != platform.system().replace("Darwin", "macOS"):
        print("⚠️  跨平台构建: 生成的可执行文件可能无法在目标平台正常运行")
    
    # 开始构建
    builder = AppBuilder()
    
    if not builder.check_requirements(target_platform):
        sys.exit(1)
    
    builder.clean_build_files()
    
    if builder.build_app(target_platform):
        builder.create_distribution_package(target_platform)
        print("\n🎉 构建完成!")
    else:
        print("\n💥 构建失败!")
        sys.exit(1)


if __name__ == "__main__":
    main()