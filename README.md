# Data to PDF Print

一个基于Python的Excel数据到PDF标签生成工具，采用Tkinter GUI界面，支持多种专业标签模板。

## 🚀 快速开始

### 系统要求
- Python 3.8+ 
- Git
- 500MB+ 磁盘空间

### 一键运行（推荐新手）

```bash
# 1. 克隆项目
git clone <your-repository-url>
cd data-to-pdfprint

# 2. 创建虚拟环境
python3 -m venv venv

# 3. 激活虚拟环境
# macOS/Linux:
source venv/bin/activate
# Windows:
# venv\Scripts\activate

# 4. 安装依赖
pip3 install -r requirements.txt

# 5. 运行程序
python src/gui_app.py
```

如果一切正常，GUI应用将启动并显示主界面。

## 📦 详细安装指南

### 1. 环境准备

**Python版本检查**
```bash
python3 --version  # 确保 >= 3.8
```

**克隆仓库**
```bash
git clone <your-repository-url>
cd data-to-pdfprint
```

### 2. 虚拟环境设置

**创建虚拟环境**
```bash
python3 -m venv venv
```

**激活虚拟环境**
```bash
# macOS/Linux
source venv/bin/activate

# Windows
venv\Scripts\activate

# 验证激活成功（提示符应显示 (venv)）
which python  # 应指向 venv 目录
```

### 3. 依赖安装

```bash
# 安装所有依赖
pip3 install -r requirements.txt

# 验证关键依赖
python -c "import tkinter; import reportlab; import pandas; print('依赖安装成功')"
```

### 4. 字体文件（Windows构建必需）

项目已包含中文字体文件：
- `src/fonts/msyh.ttf` - 微软雅黑（Windows构建必需）
- `src/fonts/msyhbd.ttc` - 微软雅黑粗体

## 🎮 运行程序

### 开发模式运行

```bash
# 确保虚拟环境已激活
python src/gui_app.py
```

### 验证功能

1. GUI界面应正常启动
2. 点击"选择Excel文件"按钮测试文件选择
3. 尝试选择不同模板类型
4. 检查是否有错误提示

### 常见运行问题

**ImportError: No module named 'xxx'**
```bash
# 重新安装依赖
pip3 install -r requirements.txt
```

**tkinter相关错误**
```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# macOS (如果使用系统Python)
# 建议使用 Homebrew 安装的 Python
```

## 🛠️ 构建打包

### 使用统一构建脚本

```bash
# 自动检测当前系统并构建
python scripts/build.py

# 指定目标平台
python scripts/build.py --platform macOS     # macOS版本
python scripts/build.py --platform Windows   # Windows版本

# 查看所有选项
python scripts/build.py --help
```

### 构建输出

**macOS构建**
```bash
python scripts/build.py --platform macOS
# 输出: dist/DataToPDF_GUI
# 使用: 双击运行或 ./dist/DataToPDF_GUI
```

**Windows构建**
```bash
python scripts/build.py --platform Windows
# 输出: dist/DataToPDF_GUI.exe
#      windows_distribution/ (分发包)
```

### 构建要求

- **PyInstaller**: `pip3 install pyinstaller`（已包含在requirements.txt）
- **磁盘空间**: 至少1GB可用空间
- **内存**: 建议4GB+

### 测试构建结果

```bash
# 测试可执行文件
./dist/DataToPDF_GUI  # macOS
# 或
dist\DataToPDF_GUI.exe  # Windows

# 验证功能完整性
```

## 📁 项目结构

```
data-to-pdfprint/
├── src/                          # 源代码目录
│   ├── __init__.py              # Python包标识
│   ├── gui_app.py               # 🚀 GUI应用主入口
│   ├── fonts/                   # 字体文件目录
│   │   ├── msyh.ttf            # 微软雅黑字体
│   │   └── msyhbd.ttc          # 微软雅黑粗体字体
│   ├── pdf/                     # PDF生成模块
│   │   ├── __init__.py          # Python包标识
│   │   ├── generator.py         # PDF生成器主类
│   │   ├── regular_box/         # 常规盒标模板
│   │   │   ├── __init__.py      # Python包标识
│   │   │   ├── data_processor.py # 数据处理器
│   │   │   ├── renderer.py      # 渲染器
│   │   │   ├── template.py      # 模板主类
│   │   │   └── ui_dialog.py     # UI配置对话框
│   │   ├── split_box/           # 分盒模板
│   │   │   ├── __init__.py      # Python包标识
│   │   │   ├── data_processor.py # 数据处理器
│   │   │   ├── renderer.py      # 渲染器
│   │   │   ├── template.py      # 模板主类
│   │   │   └── ui_dialog.py     # UI配置对话框
│   │   └── nested_box/          # 套盒模板
│   │       ├── __init__.py      # Python包标识
│   │       ├── data_processor.py # 数据处理器
│   │       ├── renderer.py      # 渲染器
│   │       ├── template.py      # 模板主类
│   │       └── ui_dialog.py     # UI配置对话框
│   └── utils/                   # 工具类目录
│       ├── base_data_processor.py # 基础数据处理器
│       ├── base_renderer.py    # 基础渲染器
│       ├── base_ui_dialog.py   # 基础UI对话框
│       ├── data_input_dialog.py # 数据输入对话框
│       ├── excel_data_extractor.py # Excel数据提取器
│       ├── font_manager.py     # 字体管理器
│       ├── pdf_base.py         # PDF基础类
│       └── text_processor.py   # 文本处理器
├── scripts/                     # 构建脚本目录
│   └── build.py                # 🔧 统一构建脚本
├── CLAUDE.md                   # Claude Code项目指导文件
├── README.md                   # 📖 项目说明文档
├── requirements.txt             # Python依赖列表
├── venv/                       # 虚拟环境（自动创建）
├── build/                      # 构建临时文件（自动创建）
├── dist/                       # 构建输出目录（自动创建）
└── output/                     # PDF输出目录
```

### 核心模块说明

- **GUI层**: `src/gui_app.py` - Tkinter主界面
- **模板系统**: `src/pdf/*/` - 三种标签模板的完整实现
- **渲染引擎**: `src/utils/base_renderer.py` - PDF渲染基础类
- **数据处理**: `src/utils/excel_data_extractor.py` - Excel读取和处理

## 🔧 开发指南

### 代码风格

```bash
# 代码格式化
black src/

# 代码检查
flake8 src/
```

## 🆘 问题排查

### 安装问题

**Python版本过低**
```bash
# 检查版本
python3 --version

# 升级Python（macOS使用Homebrew）
brew install python@3.11
```

**虚拟环境问题**
```bash
# 删除并重新创建
rm -rf venv
python3 -m venv venv
source venv/bin/activate
pip3 install -r requirements.txt
```

### 运行时错误

**字体问题（中文显示异常）**
- 确认 `src/fonts/msyh.ttf` 文件存在
- Windows系统确保字体文件完整

**Excel读取失败**
```python
# 测试Excel读取
python -c "
import pandas as pd
df = pd.read_excel('your_file.xlsx')
print(df.head())
"
```

**PDF生成失败**
- 检查输出目录权限
- 确认磁盘空间充足
- 查看控制台错误信息

### 构建问题

**PyInstaller失败**
```bash
# 清理并重新构建
rm -rf build/ dist/
python scripts/build.py
```

**文件过大**
- 正常情况：50-100MB
- 异常情况：检查是否包含不必要的文件

**缺少依赖**
```bash
# 重新安装构建依赖
pip3 install pyinstaller
```

## 🤝 贡献指南

### 代码规范

- 使用 `black` 格式化代码
- 遵循 PEP 8 规范
- 添加必要的注释和文档字符串
- 确保新功能有对应的测试

### 报告问题

提交Issue时请包含：
- 操作系统和Python版本
- 完整的错误信息
- 复现步骤
- 预期行为

## 📋 技术栈

- **GUI**: Tkinter (Python标准库)
- **PDF生成**: ReportLab
- **数据处理**: Pandas, OpenPyXL  
- **打包**: PyInstaller
- **开发工具**: Black, Flake8, Pytest