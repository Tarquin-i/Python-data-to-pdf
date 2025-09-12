"""
Windows GUI build script
Ensures the generated executable runs properly on Windows
"""

import subprocess
import sys
import os
import platform
from pathlib import Path

def build_windows_gui():
    """Build Windows GUI application"""
    
    if platform.system() != "Windows":
        print("WARNING: Not running on Windows, the generated executable may not work on Windows")
        print("Recommended to run this script on Windows system")
        
    print("Building Windows GUI application...")
    
    # Ensure we're in project root directory
    project_root = Path(__file__).parent
    os.chdir(project_root)
    
    # Check required files
    if not os.path.exists("src/gui_app.py"):
        print("ERROR: Cannot find src/gui_app.py")
        return
    
    # Check font file
    font_file = "src/fonts/msyh.ttf"
    if os.path.exists(font_file):
        file_size = os.path.getsize(font_file) / (1024 * 1024)  # MB
        print(f"[OK] Font file found: {font_file} ({file_size:.1f}MB)")
    else:
        print(f"[ERROR] Font file not found: {font_file}")
        print("Available files in src/fonts:")
        if os.path.exists("src/fonts"):
            for f in os.listdir("src/fonts"):
                print(f"  - {f}")
        else:
            print("  fonts directory does not exist!")
        return
    
    # Use Windows-specific configuration file if available, otherwise direct build
    if os.path.exists("DataToPDF_GUI_Windows.spec"):
        cmd = [
            "pyinstaller",
            "DataToPDF_GUI_Windows.spec",
            "--clean",  # Clear cache
            "--noconfirm",  # Don't ask for overwrite confirmation
        ]
    else:
        print("WARNING: DataToPDF_GUI_Windows.spec not found, using direct build")
        cmd = [
            "pyinstaller",
            "--onefile",
            "--windowed",  # No console window
            "--name=DataToPDF_GUI",
            "--clean",
            "--noconfirm",
            "--add-data", f"{font_file};fonts",  # 直接添加字体文件
            "src/gui_app.py"
        ]
    
    try:
        # Clean old files
        if os.path.exists("dist"):
            import shutil
            shutil.rmtree("dist")
        if os.path.exists("build"):
            import shutil
            shutil.rmtree("build")
        
        print("Running PyInstaller...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("SUCCESS: Windows GUI application built successfully!")
            exe_path = "dist/DataToPDF_GUI.exe"
            print(f"Generated file: {exe_path}")
            
            # 检查文件大小
            if os.path.exists(exe_path):
                file_size_mb = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"File size: {file_size_mb:.1f}MB")
                
                # 如果文件大小明显增加，说明字体被打包了
                if file_size_mb > 60:  # 预期包含字体后会超过60MB
                    print("[OK] Font file appears to be included (large file size)")
                else:
                    print("[WARNING] Font file may not be included (small file size)")
            
            print()
            print("Windows Usage Instructions:")
            print("  - Double-click DataToPDF_GUI.exe to launch GUI")
            print("  - Supports Windows 7/8/10/11")
            print("  - No Python installation required")
            print("  - Can be shared with other Windows users")
            print()
            print("Operation Steps:")
            print("1. Double-click to run DataToPDF_GUI.exe")
            print("2. Click 'Select Excel File' button to choose xlsx file")
            print("3. Click 'Generate PDF' button and choose save location")
            print("4. Automatically extract total pages and generate label PDF")
            print()
            print("Distribution Notes:")
            print("  - Can directly copy DataToPDF_GUI.exe to other users")
            print("  - Recommend packaging as ZIP file for distribution")
            print(f"  - File size: {file_size_mb:.1f}MB (includes Microsoft YaHei font)")
            
        else:
            print("ERROR: Build failed:")
            print("stdout:", result.stdout)
            print("stderr:", result.stderr)
            
            # Provide common error solutions
            if "No module named" in result.stderr:
                print()
                print("Solution:")
                print("Please install dependencies: pip install -r requirements.txt")
            elif "PyInstaller" not in result.stderr:
                print()
                print("Solution:")
                print("Please install PyInstaller: pip install pyinstaller")
                
    except FileNotFoundError:
        print("ERROR: PyInstaller not found, please install first:")
        print("pip install pyinstaller")
    except Exception as e:
        print(f"ERROR: Build process failed: {e}")

def create_distribution_package():
    """Create Windows distribution package"""
    if os.path.exists("dist/DataToPDF_GUI.exe"):
        print()
        print("Creating distribution package...")
        
        # Create distribution directory
        dist_dir = Path("windows_distribution")
        if dist_dir.exists():
            import shutil
            shutil.rmtree(dist_dir)
        dist_dir.mkdir()
        
        # Copy executable file
        import shutil
        shutil.copy2("dist/DataToPDF_GUI.exe", dist_dir / "DataToPDF_GUI.exe")
        
        # Create usage instructions
        readme_content = """# Data to PDF Print - Windows Version

## Usage Instructions
1. Double-click to run DataToPDF_GUI.exe
2. Select Excel file (.xlsx format)
3. Click Generate PDF button
4. Choose save location

## System Requirements
- Windows 7/8/10/11
- No Python installation or other dependencies required

## Notes
- First run may be detected by Windows Defender, select "Allow"
- Recommend placing program on non-system drive (e.g., D: drive)
- Supported Excel formats: .xlsx, .xls

## Feedback
Contact developer if you encounter any issues
"""
        
        with open(dist_dir / "README.txt", "w", encoding="utf-8") as f:
            f.write(readme_content)
        
        print(f"SUCCESS: Distribution package created: {dist_dir.absolute()}")
        print("Included files:")
        for file in dist_dir.iterdir():
            print(f"  - {file.name}")

if __name__ == "__main__":
    build_windows_gui()
    create_distribution_package()