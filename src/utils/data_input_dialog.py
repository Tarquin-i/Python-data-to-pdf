"""
数据输入补充对话框
当Excel数据提取不完整时，让用户手动输入缺失的数据
"""

import tkinter as tk
from tkinter import messagebox, ttk


class DataInputDialog:
    """数据输入补充对话框 - 智能对比显示已有数据和缺失数据"""

    def __init__(self, parent, extracted_data):
        """
        初始化数据输入对话框

        Args:
            parent: 父窗口
            extracted_data: 已提取的数据字典
        """
        self.parent = parent
        self.extracted_data = extracted_data
        self.result_data = None
        self.input_vars = {}

        # 定义五个核心字段及其格式要求
        self.required_fields = {
            "客户名称编码": {
                "format_hint": "如：14KH0149、ABC123",
                "validation": self._validate_client_code,
            },
            "标签名称": {
                "format_hint": "如：LADIES NIGHT IN、产品名称",
                "validation": self._validate_theme,
            },
            "开始号": {
                "format_hint": "如：DSK00001、JAW01001-01",
                "validation": self._validate_start_number,
            },
            "总张数": {
                "format_hint": "正整数，如：57000、1000",
                "validation": self._validate_total_count,
            },
            "张/盒": {
                "format_hint": "正整数，如：300、500",
                "validation": self._validate_pieces_per_box,
            },
        }

    def show_dialog(self):
        """显示数据输入对话框"""
        # 检查哪些数据缺失
        missing_fields = []
        for field in self.required_fields.keys():
            value = self.extracted_data.get(field)
            if value is None or str(value).strip() == "" or str(value) == "0":
                missing_fields.append(field)

        # 如果没有缺失数据，直接返回原数据
        if not missing_fields:
            self.result_data = self.extracted_data.copy()
            return self.result_data

        # 创建对话框
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("数据补充 - Excel数据不完整")
        self.dialog.transient(self.parent)
        self.dialog.grab_set()

        # 设置对话框尺寸
        self.dialog.geometry("650x500")
        self.dialog.resizable(True, True)

        # 创建主框架
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(
            main_frame, text="📋 数据提取结果检查", font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 15))

        # 说明文字
        desc_label = ttk.Label(
            main_frame,
            text="Excel文件中部分数据缺失或为空，请补充以下信息：",
            font=("Arial", 10),
        )
        desc_label.pack(pady=(0, 20))

        # 数据对比表格框架
        table_frame = ttk.LabelFrame(main_frame, text="数据状态对比", padding="15")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # 创建表格式布局
        self._create_data_table(table_frame, missing_fields)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))

        # 确认按钮
        confirm_btn = ttk.Button(
            button_frame, text="✅ 确认并继续", command=self._confirm_data
        )
        confirm_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 取消按钮
        cancel_btn = ttk.Button(button_frame, text="❌ 取消", command=self._cancel)
        cancel_btn.pack(side=tk.LEFT)

        # 居中显示
        self._center_dialog()

        # 等待用户操作
        self.dialog.wait_window()

        return self.result_data

    def _create_data_table(self, parent, missing_fields):
        """创建数据对比表格"""
        # 表格标题行
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            header_frame, text="数据字段", font=("Arial", 11, "bold"), width=12
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(
            header_frame, text="当前值", font=("Arial", 11, "bold"), width=30
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(header_frame, text="状态", font=("Arial", 11, "bold"), width=8).pack(
            side=tk.LEFT
        )

        # 分隔线
        separator = ttk.Separator(parent, orient="horizontal")
        separator.pack(fill=tk.X, pady=(0, 15))

        # 数据行
        for field_name in self.required_fields.keys():
            self._create_field_row(parent, field_name, field_name in missing_fields)

    def _create_field_row(self, parent, field_name, is_missing):
        """创建单个数据字段行"""
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, pady=5)

        # 字段名
        field_label = ttk.Label(row_frame, text=field_name, width=12)
        field_label.pack(side=tk.LEFT, padx=(0, 10))

        # 当前值或输入框
        value_frame = ttk.Frame(row_frame)
        value_frame.pack(side=tk.LEFT, padx=(0, 10))

        current_value = self.extracted_data.get(field_name, "").strip()

        if is_missing:
            # 缺失数据 - 显示输入框
            self.input_vars[field_name] = tk.StringVar()

            input_frame = ttk.Frame(value_frame)
            input_frame.pack(anchor=tk.W)

            entry = ttk.Entry(
                input_frame,
                textvariable=self.input_vars[field_name],
                width=25,
                font=("Arial", 10),
            )
            entry.pack(side=tk.LEFT)

            # 格式提示
            hint_text = self.required_fields[field_name]["format_hint"]
            hint_label = ttk.Label(
                input_frame,
                text=f"  格式：{hint_text}",
                font=("Arial", 9),
                foreground="gray",
            )
            hint_label.pack(side=tk.LEFT)

            # 状态图标
            status_label = ttk.Label(
                row_frame, text="❌ 缺失", foreground="red", width=8
            )
            status_label.pack(side=tk.LEFT)

            # 第一个缺失字段自动获得焦点
            if len(self.input_vars) == 1:
                entry.focus()
        else:
            # 已有数据 - 显示当前值
            value_label = ttk.Label(
                value_frame,
                text=current_value,
                font=("Arial", 10, "bold"),
                width=30,
                anchor=tk.W,
            )
            value_label.pack(anchor=tk.W)

            # 状态图标
            status_label = ttk.Label(
                row_frame, text="✅ 已有", foreground="green", width=8
            )
            status_label.pack(side=tk.LEFT)

    def _confirm_data(self):
        """确认并验证用户输入的数据"""
        # 验证所有输入
        for field_name, var in self.input_vars.items():
            value = var.get().strip()

            # 检查是否为空
            if not value:
                messagebox.showerror("输入错误", f"请输入【{field_name}】")
                return

            # 格式验证
            validation_func = self.required_fields[field_name]["validation"]
            if not validation_func(value):
                return  # 验证函数内部会显示错误消息

        # 合并数据
        self.result_data = self.extracted_data.copy()
        for field_name, var in self.input_vars.items():
            self.result_data[field_name] = var.get().strip()

        self.dialog.destroy()

    def _cancel(self):
        """取消操作"""
        self.result_data = None
        self.dialog.destroy()

    def _center_dialog(self):
        """居中显示对话框"""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() - width) // 2
        y = (self.dialog.winfo_screenheight() - height) // 2
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")

    # 数据验证方法
    def _validate_client_code(self, value):
        """验证客户编码"""
        if len(value) < 2:
            messagebox.showerror(
                "格式错误",
                "客户编码长度至少2个字符\n\n正确格式：14KH0149、ABC123",
            )
            return False
        return True

    def _validate_theme(self, value):
        """验证主题"""
        if len(value) < 1:
            messagebox.showerror(
                "格式错误",
                "主题不能为空\n\n正确格式：LADIES NIGHT IN、产品名称",
            )
            return False
        return True

    def _validate_start_number(self, value):
        """验证开始号"""
        if len(value) < 3:
            messagebox.showerror(
                "格式错误",
                "开始号格式不正确\n\n正确格式：DSK00001、JAW01001-01",
            )
            return False
        return True

    def _validate_total_count(self, value):
        """验证总张数"""
        try:
            count = int(value)
            if count <= 0:
                messagebox.showerror(
                    "格式错误", "总张数必须为正整数\n\n正确格式：57000、1000"
                )
                return False
            return True
        except ValueError:
            messagebox.showerror(
                "格式错误", "总张数必须为数字\n\n正确格式：57000、1000"
            )
            return False

    def _validate_pieces_per_box(self, value):
        """验证张/盒"""
        try:
            pieces = int(value)
            if pieces <= 0:
                messagebox.showerror(
                    "格式错误", "张/盒必须为正整数\n\n正确格式：300、500"
                )
                return False

            # 检查张/盒是否超过总张数
            total_count_str = self.input_vars.get("总张数", tk.StringVar()).get()
            if total_count_str and total_count_str.strip():
                try:
                    total_count = int(total_count_str)
                    if pieces > total_count:
                        messagebox.showerror(
                            "数据错误",
                            f"张/盒不能超过总张数\n\n当前输入：\n张/盒：{pieces}\n总张数：{total_count}",
                        )
                        return False
                except ValueError:
                    pass  # 总张数格式错误时不做此验证，由总张数验证处理

            return True
        except ValueError:
            messagebox.showerror("格式错误", "张/盒必须为数字\n\n正确格式：300、500")
            return False


def show_data_input_dialog(parent, extracted_data):
    """显示数据输入对话框的便捷函数"""
    dialog = DataInputDialog(parent, extracted_data)
    return dialog.show_dialog()
