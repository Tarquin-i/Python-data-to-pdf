# CLAUDE.md

此文件为 Claude Code (claude.ai/code) 在此代码库中工作时提供指导。

## 项目概述

Data to PDF Print (data-to-pdfprint) 是一个用于读取Excel数据并生成PDF标签的Python GUI应用程序。它支持多种模板（常规、分盒、套盒），允许从同一数据生成不同样式的PDF标签。

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

### 运行应用
```bash
# 运行GUI应用
python src/gui_app.py
```

### 构建GUI应用程序

使用统一的构建脚本 `scripts/build.py`：

```bash
# 自动检测当前系统并构建
python scripts/build.py

# 指定目标平台构建
python scripts/build.py --platform macOS     # 构建macOS版本
python scripts/build.py --platform Windows   # 构建Windows版本

# 查看帮助信息
python scripts/build.py --help
```

#### 构建要求
- **macOS**: 在macOS系统上构建，适用于M1/M2 Mac
- **Windows**: 建议在Windows系统上构建，需要中文字体文件 `src/fonts/msyh.ttf`
- **依赖**: 需要安装pyinstaller (`pip install pyinstaller`)

#### 分发方式
- **macOS**: `dist/DataToPDF_GUI` (可执行文件)
- **Windows**: `dist/DataToPDF_GUI.exe` + `windows_distribution/` (分发包)
- **跨平台**: 使用源代码分发配合 `requirements.txt`

## 架构设计

### 核心模块结构

项目采用模块化架构，职责清晰分离：

- **GUI层** (`src/gui_app.py`): 图形用户界面主入口
- **模板系统** (`src/pdf/`): 常规、分盒、套盒三种模板系统
- **PDF生成** (`src/pdf/`): 基于ReportLab的PDF创建
- **渲染器** (`src/pdf/*/renderer.py`): 各模板专属渲染器
- **UI对话框** (`src/pdf/*/ui_dialog.py`): 各模板参数设置界面
- **工具类** (`src/utils/`): 通用辅助函数和基础渲染器
- **构建脚本** (`scripts/`): 应用程序构建和打包脚本

### 主要依赖

- `reportlab>=3.6.0` - PDF生成
- `pandas>=1.5.0` - 数据操作
- `openpyxl>=3.1.0` - Excel文件读取

### 模板系统设计

模板系统围绕以下核心概念设计：
- **BaseTemplate**: 定义模板接口的基类
- **Template Manager**: 处理模板创建、加载和管理
- **内置模板**: 预定义的常用模板
- 模板存储在 `/templates/` 目录中（待创建）

### 数据流程

1. **Excel读取**: `ExcelDataExtractor` 类读取和解析Excel文件
2. **数据处理**: 各模板的 `DataProcessor` 处理特定模板数据
3. **模板应用**: 用户通过GUI选择模板和参数
4. **PDF生成**: 各模板的 `Template` 类生成相应的PDF文件

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

### 应用入口
GUI应用入口为 `src/gui_app.py`

### 文件命名约定
Python文件和模块使用小写加下划线（snake_case）命名。