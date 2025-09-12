# CLAUDE.md

此文件为 Claude Code (claude.ai/code) 在此代码库中工作时提供指导。

## 项目概述

Data to PDF Print (data-to-pdfprint) 是一个用于读取Excel数据并生成PDF标签的Python CLI工具。它支持自定义模板，允许从同一数据生成不同样式的PDF标签。

## 开发命令

### 安装和设置
```bash
# 安装依赖
pip install -r requirements.txt

# 以开发模式安装项目
pip install -e .
```

### 测试
```bash
# 运行所有测试
python -m pytest tests/

# 运行特定测试文件（当tests目录存在时）
python -m pytest tests/test_excel_reader.py
```

### 代码质量
```bash
# 格式化代码
black src/

# 代码检查
flake8 src/
```

### 运行CLI
```bash
# 使用默认模板的基本用法
data-to-pdf --input data.xlsx --template basic

# 自定义模板用法
data-to-pdf --input data.xlsx --template custom_template

# 批量生成
data-to-pdf --input data.xlsx --template basic --output output_dir/
```

### 构建GUI应用程序

#### macOS版本（当前系统）
```bash
# 构建macOS版本
python build_gui.py
```

#### Windows版本（跨平台构建）
```bash
# 构建Windows版本（最好在Windows系统上运行）
python build_windows.py

# 或直接使用PyInstaller和Windows配置文件
pyinstaller DataToPDF_GUI_Windows.spec --clean --noconfirm
```

#### 分发方式
- **macOS**: `dist/DataToPDF_GUI` (ARM64架构，适用于M1/M2 Mac)
- **Windows**: `dist/DataToPDF_GUI.exe` (使用 `build_windows.py`)
- **跨平台**: 使用源代码分发配合 `requirements.txt`

## 架构设计

### 核心模块结构

项目采用模块化架构，职责清晰分离：

- **CLI层** (`src/cli/`): 命令行界面和参数解析
- **数据层** (`src/data/`): Excel读取和数据处理
- **模板系统** (`src/template/`): 模板管理和渲染（计划中）
- **PDF生成** (`src/pdf/`): 基于ReportLab的PDF创建
- **配置管理** (`src/config/`): 设置和配置管理
- **工具类** (`src/utils/`): 通用辅助函数

### 主要依赖

- `reportlab>=3.6.0` - PDF生成
- `pandas>=1.5.0` - 数据操作
- `openpyxl>=3.1.0` - Excel文件读取
- `click>=8.1.0` - CLI框架

### 模板系统设计

模板系统围绕以下核心概念设计：
- **BaseTemplate**: 定义模板接口的基类
- **Template Manager**: 处理模板创建、加载和管理
- **内置模板**: 预定义的常用模板
- 模板存储在 `/templates/` 目录中（待创建）

### 数据流程

1. **Excel读取**: `ExcelReader` 类使用pandas/openpyxl读取Excel文件
2. **数据处理**: `DataProcessor` 提取和验证特定字段
3. **模板应用**: 模板系统根据选定模板渲染数据
4. **PDF生成**: `PDFGenerator` 使用ReportLab创建最终PDF

### 配置管理

`Settings` 类管理：
- 项目根目录路径
- 模板目录位置 (`PROJECT_ROOT/templates`)
- 默认输出目录 (`PROJECT_ROOT/output`)
- 配置文件加载/保存

## 开发注意事项

### 当前实现状态
代码库目前包含方法存根的骨架类。大部分核心功能需要实现。

### 缺失目录
- `/templates/` - 模板文件存储
- `/tests/` - 测试文件（README中存在计划结构）

### 入口点
CLI入口点在setup.py中配置为 `data-to-pdf=src.cli.main:main`

### 文件命名约定
Python文件和模块使用小写加下划线（snake_case）命名。