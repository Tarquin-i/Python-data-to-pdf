"""
通用UI对话框基类
抽取所有模板共同的UI组件和逻辑，减少UI代码重复
"""

import tkinter as tk
from abc import ABC, abstractmethod
from tkinter import messagebox, ttk
from typing import Dict, List, Tuple


class BaseUIDialog(ABC):
    """通用UI对话框基类"""

    def __init__(self, main_app):
        """
        初始化基类UI对话框

        Args:
            main_app: 主GUI应用程序实例
        """
        self.main_app = main_app
        self.dialog = None
        self.input_vars = {}

    # ==================== 通用对话框创建方法 ====================

    def create_base_dialog(self, title: str, width=580, height=480):
        """创建基础对话框窗口"""
        if not self.main_app.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return None

        # 创建对话框
        dialog = tk.Toplevel(self.main_app.root)
        dialog.title(title)
        dialog.transient(self.main_app.root)
        dialog.grab_set()

        # 设置最小尺寸
        dialog.minsize(width - 30, height - 30)
        dialog.resizable(True, True)
        dialog.geometry(f"{width}x{height}")

        self.dialog = dialog
        return dialog

    def create_scrollable_frame(self, dialog):
        """创建可滚动的主框架"""
        # 创建滚动框架
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # 绑定鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<MouseWheel>", _on_mousewheel)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 主框架在可滚动区域内
        main_frame = ttk.Frame(scrollable_frame, padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_columnconfigure(0, weight=1)

        return main_frame

    def create_title_label(self, parent, title: str):
        """创建标题标签"""
        title_label = ttk.Label(parent, text=title, font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))
        return title_label

    # ==================== 通用组件创建方法 ====================

    def create_radio_group(
        self,
        parent,
        title: str,
        options: List[Tuple[str, str]],
        var_name: str,
        default_value: str = None,
        command=None,
    ):
        """
        创建单选按钮组

        Args:
            parent: 父容器
            title: 组标题
            options: 选项列表，格式为[(显示文本, 值), ...]
            var_name: 变量名
            default_value: 默认值
            command: 选择变化时的回调函数
        """
        frame = ttk.LabelFrame(parent, text=title, padding="15")
        frame.pack(fill=tk.X, pady=(0, 20))

        # 创建变量
        var = tk.StringVar(value=default_value or options[0][1])
        self.input_vars[var_name] = var

        # 居中布局的框架
        container = ttk.Frame(frame)
        container.pack(expand=True)

        # 创建单选按钮
        for i, (text, value) in enumerate(options):
            radio = ttk.Radiobutton(
                container,
                text=text,
                variable=var,
                value=value,
                command=command,
            )
            if i == 0:
                radio.grid(row=0, column=i, sticky=tk.W, pady=5)
            else:
                radio.grid(row=0, column=i, sticky=tk.W, padx=(20, 0), pady=5)

        return frame, var

    def create_input_section(self, parent, title: str, input_configs: List[Dict]):
        """
        创建输入参数区域

        Args:
            parent: 父容器
            title: 区域标题
            input_configs: 输入配置列表，每个元素包含：
                - label: 标签文本
                - var_name: 变量名
                - default_value: 默认值
                - width: 输入框宽度
                - help_text: 帮助文本（可选）
        """
        frame = ttk.LabelFrame(parent, text=title, padding="15")
        frame.pack(fill=tk.X, pady=(0, 20))

        for i, config in enumerate(input_configs):
            # 标签
            ttk.Label(frame, text=config["label"]).grid(
                row=i, column=0, sticky=tk.W, pady=5
            )

            # 输入框
            var = tk.StringVar(value=config.get("default_value", ""))
            self.input_vars[config["var_name"]] = var

            entry = ttk.Entry(frame, textvariable=var, width=config.get("width", 15))
            entry.grid(row=i, column=1, sticky=tk.W, padx=(10, 0), pady=5)

            # 帮助文本
            if config.get("help_text"):
                help_label = ttk.Label(
                    frame,
                    text=config["help_text"],
                    font=("Arial", 8),
                    foreground="gray",
                )
                help_label.grid(row=i, column=2, sticky=tk.W, padx=(5, 0), pady=5)

        return frame

    def create_data_display_section(self, parent):
        """创建当前数据显示区域"""
        data_frame = ttk.LabelFrame(parent, text="当前数据", padding="15")
        data_frame.pack(fill=tk.X, pady=(0, 20))

        data_text = f"客户名称编码: {self.main_app.current_data['客户名称编码']}\n"
        data_text += f"标签名称: {self.main_app.current_data['标签名称']}\n"
        data_text += f"总张数: {self.main_app.current_data['总张数']}"

        data_label = ttk.Label(data_frame, text=data_text, font=("Consolas", 10))
        data_label.pack(anchor=tk.W)

        return data_frame

    def create_button_section(self, parent, confirm_command, cancel_command=None):
        """创建按钮区域"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(pady=(10, 0))

        # 确认按钮
        confirm_btn = ttk.Button(
            button_frame,
            text="确认生成",
            command=confirm_command,
        )
        confirm_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 取消按钮
        cancel_cmd = cancel_command or self.dialog.destroy
        cancel_btn = ttk.Button(button_frame, text="取消", command=cancel_cmd)
        cancel_btn.pack(side=tk.LEFT)

        return button_frame

    # ==================== 通用验证方法 ====================

    def validate_integer_input(
        self, var_name: str, field_name: str, min_value=1, max_value=None
    ) -> int:
        """
        验证整数输入

        Args:
            var_name: 变量名
            field_name: 字段显示名称
            min_value: 最小值
            max_value: 最大值

        Returns:
            验证通过的整数值

        Raises:
            ValueError: 验证失败
        """
        value_str = self.input_vars[var_name].get().strip()

        if not value_str:
            raise ValueError(f"请输入'{field_name}'参数")

        try:
            value = int(value_str)
        except ValueError:
            raise ValueError(
                f"'{field_name}'必须为有效整数\\n\\n正确格式示例：{min_value}"
            )

        if value < min_value:
            raise ValueError(
                f"'{field_name}'必须大于等于{min_value}\\n\\n当前值：{value}"
            )

        if max_value is not None and value > max_value:
            raise ValueError(f"'{field_name}'不能超过{max_value}\\n\\n当前值：{value}")

        return value

    def validate_pieces_per_box(self) -> int:
        """验证张/盒参数"""
        total_pieces = int(self.main_app.current_data.get("总张数", 0))
        pieces_per_box = self.validate_integer_input(
            "pieces_per_box", "张/盒", 1, total_pieces
        )

        if pieces_per_box > total_pieces:
            raise ValueError(
                f"'张/盒'不能超过总张数\\n\\n当前设置：{pieces_per_box} 张/盒\\n总张数：{total_pieces} 张\\n\\n请输入不超过{total_pieces}的值"
            )

        return pieces_per_box

    def validate_required_string(self, var_name: str, field_name: str) -> str:
        """验证必填字符串字段"""
        value = self.input_vars[var_name].get().strip()
        if not value:
            raise ValueError(f"请输入'{field_name}'")
        return value

    def validate_font_size(
        self, var_name: str = "serial_font_size", min_size=6, max_size=14
    ) -> int:
        """验证字体大小"""
        try:
            font_size = self.validate_integer_input(
                var_name, "序列号字体大小", min_size, max_size
            )
            return font_size
        except ValueError as e:
            if "大于等于" in str(e) or "不能超过" in str(e):
                raise ValueError(
                    f"序列号字体大小必须在{min_size}-{max_size}之间\\n\\n当前值：{self.input_vars[var_name].get()}"
                )
            else:
                raise ValueError("请输入有效的序列号字体大小\\n\\n正确格式示例：10")

    # ==================== 通用工具方法 ====================

    def auto_resize_and_center_dialog(self, content_frame):
        """自动调整对话框大小并居中显示"""
        try:
            # 多次更新确保所有组件都已完全渲染
            for _ in range(3):
                self.dialog.update_idletasks()
                content_frame.update_idletasks()

            # 获取内容的实际所需尺寸
            content_width = content_frame.winfo_reqwidth()
            content_height = content_frame.winfo_reqheight()

            # 添加必要的边距
            padding_width = 60
            padding_height = 80

            # 计算对话框所需的实际尺寸
            required_width = content_width + padding_width
            required_height = content_height + padding_height

            # 获取屏幕尺寸，确保不会超出屏幕
            screen_width = self.dialog.winfo_screenwidth()
            screen_height = self.dialog.winfo_screenheight()

            # 最终尺寸：完全基于内容，但不超过屏幕90%
            final_width = min(required_width, int(screen_width * 0.9))
            final_height = min(required_height, int(screen_height * 0.9))

            # 计算居中位置
            x = (screen_width - final_width) // 2
            y = (screen_height - final_height) // 2

            # 设置对话框几何形状
            self.dialog.geometry(f"{final_width}x{final_height}+{x}+{y}")

        except Exception as e:
            print(f"⚠️ 自适应调整失败: {e}")
            # 备用方案：让系统自动计算
            self.dialog.update_idletasks()
            self.dialog.geometry("")
            self.dialog.update_idletasks()

            # 获取自动调整后的尺寸并居中
            width = self.dialog.winfo_width()
            height = self.dialog.winfo_height()
            x = (self.dialog.winfo_screenwidth() - width) // 2
            y = (self.dialog.winfo_screenheight() - height) // 2
            self.dialog.geometry(f"{width}x{height}+{x}+{y}")

    def show_error(self, message: str):
        """显示错误消息"""
        messagebox.showerror("参数错误", message)

    def get_var_value(self, var_name: str):
        """获取变量值"""
        return (
            self.input_vars.get(var_name).get() if var_name in self.input_vars else None
        )

    def get_var(self, var_name: str):
        """获取变量对象（用于访问tkinter变量）"""
        return self.input_vars.get(var_name)

    def prefill_from_excel(self, var_name: str, excel_key: str):
        """从Excel数据预填充输入框"""
        if excel_key in self.main_app.current_data:
            value = self.main_app.current_data[excel_key]
            if value and value != "N/A":
                self.input_vars[var_name].set(str(value))

    # ==================== 抽象方法 ====================

    @abstractmethod
    def show_parameters_dialog(self):
        """显示参数设置对话框"""

    @abstractmethod
    def confirm_parameters(self):
        """确认参数并生成PDF"""

    @abstractmethod
    def get_template_specific_inputs(self) -> List[Dict]:
        """获取模板特定的输入配置"""
