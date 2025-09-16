"""
套盒模板UI对话框
专门处理套盒模板的参数设置界面，基于通用基类重构
"""

from tkinter import ttk
from typing import Dict, List

from src.utils.base_ui_dialog import BaseUIDialog


class NestedBoxUIDialog(BaseUIDialog):
    """套盒模板UI对话框处理类"""

    def __init__(self, main_app):
        """初始化套盒模板UI对话框"""
        super().__init__(main_app)

    def show_parameters_dialog(self):
        """显示套盒模板的参数设置对话框"""
        # 创建对话框
        dialog = self.create_base_dialog("套盒模板 - 包装参数设置", 500, 420)
        if not dialog:
            return

        # 创建可滚动框架
        main_frame = self.create_scrollable_frame(dialog)

        # 创建标题
        self.create_title_label(main_frame, "套盒模板参数设置")

        # 创建包装类型选择
        self.create_radio_group(
            main_frame,
            "包装类型选择",
            [("正常（多套装箱）", "正常"), ("超重（一套拆多箱）", "超重")],
            "is_overweight",
            "正常",
            self.on_overweight_choice_changed,
        )

        # 创建小箱类型选择
        self.create_radio_group(
            main_frame,
            "小箱类型选择",
            [
                ("有小箱（三级包装）", "有小箱"),
                ("无小箱（二级包装）", "无小箱"),
            ],
            "has_small_box",
            "有小箱",
        )

        # 创建盒标类型选择
        self.create_radio_group(
            main_frame,
            "盒标类型选择",
            [("无盒标", "无盒标"), ("有盒标", "有盒标")],
            "has_box_label",
            "无盒标",
        )

        # 创建参数输入区域
        input_configs = [
            {
                "label": "张/盒:",
                "var_name": "pieces_per_box",
                "default_value": self.main_app.current_data.get("张/盒", ""),
                "width": 15,
            },
            {
                "label": "盒/套:",
                "var_name": "boxes_per_small_box",
                "default_value": "",
                "width": 15,
            },
            {
                "label": "套/箱:",
                "var_name": "small_boxes_per_large_box",
                "default_value": "",
                "width": 15,
            },
            {
                "label": "中文名称:",
                "var_name": "chinese_name",
                "default_value": "",
                "width": 15,
            },
            {
                "label": "序列号字体大小:",
                "var_name": "serial_font_size",
                "default_value": "10",
                "width": 15,
                "help_text": "(建议范围: 6-14, 序列号较长时可调小)",
            },
        ]

        params_frame = self.create_input_section(main_frame, "包装参数", input_configs)
        self.params_frame = params_frame  # 保存引用以便动态更新标签

        # 创建标签模版选择
        self.create_radio_group(
            main_frame,
            "标签模版选择",
            [("无纸卡备注", "无纸卡备注"), ("有纸卡备注", "有纸卡备注")],
            "template_type",
            "无纸卡备注",
        )

        # 创建当前数据显示区域
        self.create_data_display_section(main_frame)

        # 创建按钮区域
        self.create_button_section(main_frame, self.confirm_parameters)

        # 自适应大小和居中显示
        self.auto_resize_and_center_dialog(main_frame)

        # 设置第三参数标签引用
        self.third_param_label = None
        for widget in params_frame.winfo_children():
            if isinstance(widget, ttk.Label) and widget.cget("text") == "套/箱:":
                self.third_param_label = widget
                break

    def on_overweight_choice_changed(self):
        """处理超重选择变化"""
        if self.third_param_label:
            is_overweight = self.get_var_value("is_overweight") == "超重"
            if is_overweight:
                self.third_param_label.config(text="一套拆多少箱:")
            else:
                self.third_param_label.config(text="套/箱:")

    def confirm_parameters(self):
        """确认套盒模板参数并生成PDF"""
        try:
            # 获取和验证参数
            pieces_per_box = self.validate_pieces_per_box()
            boxes_per_small_box = self.validate_integer_input(
                "boxes_per_small_box", "盒/套", 1
            )

            is_overweight = self.get_var_value("is_overweight") == "超重"

            if is_overweight:
                small_boxes_per_large_box = self.validate_integer_input(
                    "small_boxes_per_large_box",
                    "一套拆多少箱",
                    1,
                    boxes_per_small_box,
                )
            else:
                small_boxes_per_large_box = self.validate_integer_input(
                    "small_boxes_per_large_box", "套/箱", 1
                )

            chinese_name = self.validate_required_string("chinese_name", "中文名称")
            serial_font_size = self.validate_font_size()

            # 设置参数
            self.main_app.packaging_params = {
                "张/盒": pieces_per_box,
                "盒/套": boxes_per_small_box,
                "套/箱": small_boxes_per_large_box,
                "是否超重": is_overweight,
                "选择外观": "外观一",
                "标签模版": self.get_var_value("template_type"),
                "中文名称": chinese_name,
                "序列号字体大小": serial_font_size,
                "是否有小箱": self.get_var_value("has_small_box") == "有小箱",
                "是否有盒标": self.get_var_value("has_box_label") == "有盒标",
            }

            self.dialog.destroy()
            self.main_app.generate_multi_level_pdfs()

        except ValueError as e:
            self.show_error(str(e))

    # ==================== 实现抽象方法 ====================

    def get_template_specific_inputs(self) -> List[Dict]:
        """获取套盒模板特定的输入配置"""
        return [
            {
                "label": "张/盒:",
                "var_name": "pieces_per_box",
                "default_value": "",
                "width": 15,
            },
            {
                "label": "盒/套:",
                "var_name": "boxes_per_small_box",
                "default_value": "",
                "width": 15,
            },
            {
                "label": "套/箱:",
                "var_name": "small_boxes_per_large_box",
                "default_value": "",
                "width": 15,
            },
            {
                "label": "中文名称:",
                "var_name": "chinese_name",
                "default_value": "",
                "width": 15,
            },
            {
                "label": "序列号字体大小:",
                "var_name": "serial_font_size",
                "default_value": "10",
                "width": 15,
            },
        ]


# 全局变量用于单例模式
nested_box_ui_dialog = None


def get_nested_box_ui_dialog(main_app):
    """获取套盒模板UI对话框实例（单例模式）"""
    global nested_box_ui_dialog
    if nested_box_ui_dialog is None:
        nested_box_ui_dialog = NestedBoxUIDialog(main_app)
    return nested_box_ui_dialog
