# Data to PDF Print

一个用于读取Excel数据并生成PDF标签的Python命令行工具。支持自定义模板，可以根据相同的数据生成不同风格的PDF标签。

## 功能特性

- 📊 **Excel数据读取**: 支持读取Excel文件并提取指定字段
- 🎨 **模板系统**: 支持自定义PDF模板，实现不同风格的标签输出
- 📄 **PDF生成**: 使用ReportLab库生成高质量PDF文档
- ⚡ **批量处理**: 支持批量生成多个PDF标签
- 🖥️ **命令行界面**: 简单易用的命令行操作

## 项目结构

```
data-to-pdfprint/
├── src/                        # 源代码目录
│   ├── __init__.py             # 包初始化文件
│   ├── cli/                    # 命令行界面模块
│   │   ├── __init__.py
│   │   └── main.py             # CLI主程序入口
│   ├── data/                   # 数据处理模块
│   │   ├── __init__.py
│   │   ├── excel_reader.py     # Excel文件读取器
│   │   └── data_processor.py   # 数据处理器
│   ├── pdf/                    # PDF生成模块
│   │   ├── __init__.py
│   │   └── generator.py        # PDF生成器
│   ├── config/                 # 配置管理模块
│   │   ├── __init__.py
│   │   └── settings.py         # 项目配置管理
│   └── utils/                  # 工具函数模块
│       ├── __init__.py
│       └── helpers.py          # 通用工具函数
├── requirements.txt            # 项目依赖
├── setup.py                    # 安装配置
├── .gitignore                  # Git忽略文件
└── README.md                   # 项目说明文档
```

## 模块说明

### CLI模块 (`src/cli/`)
- **main.py**: 命令行主程序入口，负责处理命令行参数解析和主要流程控制

### 数据模块 (`src/data/`)
- **excel_reader.py**: Excel文件读取器，负责读取Excel文件并解析数据
- **data_processor.py**: 数据处理器，负责处理和转换Excel中的原始数据

### PDF模块 (`src/pdf/`)
- **generator.py**: PDF生成器，使用ReportLab生成PDF文档

### 配置模块 (`src/config/`)
- **settings.py**: 项目配置管理，管理项目的各种配置选项

### 工具模块 (`src/utils/`)
- **helpers.py**: 通用工具函数，提供项目中常用的工具函数

## 设计原则

1. **单一职责**: 每个模块和类都只负责一个具体的功能
2. **模块化设计**: 功能模块之间松耦合，易于维护和扩展
3. **可扩展性**: 支持自定义模板，可以轻松添加新的PDF样式
4. **可测试性**: 每个模块都有对应的测试文件
5. **配置管理**: 统一的配置管理，方便部署和维护

## 快速开始

### 1. 环境准备

首先确保你的系统已安装 Python 3.8 或更高版本。

### 2. 创建虚拟环境
```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
# macOS/Linux:
source venv/bin/activate
# Windows:
# venv\Scripts\activate
```

### 3. 安装依赖
```bash
# 安装项目依赖
pip install -r requirements.txt
```

### 4. 安装项目
```bash
# 开发模式安装 (推荐，修改代码无需重新安装)
pip install -e .
```

### 5. 验证安装
```bash
# 查看命令帮助
data-to-pdf --help

# 查看版本信息
data-to-pdf --version

# 运行基本功能测试
data-to-pdf
```

### 6. 基本使用
```bash
# 使用默认模板处理Excel文件
data-to-pdf --input data.xlsx --template basic

# 指定输出目录
data-to-pdf --input data.xlsx --template basic --output output/

# 快速测试（无参数运行查看使用提示）
data-to-pdf

# 选择1：在Windows系统上构建（推荐）
# 这会生成 dist/DataToPDF_GUI.exe，可以直接在 Windows 上运行。

python build_windows.py

```

## setup.py 说明

`setup.py` 是Python项目的配置文件，具有以下作用：

- **项目元数据**: 定义项目名称、版本、作者等信息
- **依赖管理**: 自动读取requirements.txt中的依赖包
- **命令行工具注册**: 创建 `data-to-pdf` 命令行工具
- **包安装**: 支持 `pip install` 安装项目
- **开发模式**: 支持 `pip install -e .` 可编辑安装

## 打包成可执行文件

### 使用 PyInstaller 打包

项目已配置好 PyInstaller 支持，可以将Python项目打包成独立的可执行文件：

```bash
# 1. 激活虚拟环境
source venv/bin/activate

# 2. 安装打包工具 (如果尚未安装)
pip install pyinstaller

# 3. 使用预配置的spec文件打包
pyinstaller data-to-pdf.spec
```

打包完成后，可执行文件位于 `dist/data-to-pdf`

### 测试可执行文件

```bash
# 查看帮助
./dist/data-to-pdf --help

# 查看版本
./dist/data-to-pdf --version

# 运行程序
./dist/data-to-pdf
```

### 打包配置说明

- **`data-to-pdf.spec`**: PyInstaller配置文件，包含所有打包设置
- **`build_exe.py`**: 可选的Python脚本，提供编程方式的打包
- **单文件输出**: 生成单个可执行文件，无需额外依赖
- **文件大小**: 约32MB (包含所有Python运行环境和依赖)

### 跨平台支持

- **macOS**: 生成 `data-to-pdf` 可执行文件
- **Windows**: 需要在Windows环境下打包生成 `data-to-pdf.exe`
- **Linux**: 需要在Linux环境下打包生成相应的可执行文件
