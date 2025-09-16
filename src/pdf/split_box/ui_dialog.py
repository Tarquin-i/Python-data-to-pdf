"""
分盒模板UI对话框
基于BaseUIDialog重构，专门处理分盒模板的参数设置界面
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict

from src.utils.base_ui_dialog import BaseUIDialog


class SplitBoxUIDialog(BaseUIDialog):
    """分盒模板UI对话框处理类"""
    
    def __init__(self, main_app):
        """初始化分盒模板UI对话框"""
        super().__init__(main_app)
    
    def show_parameters_dialog(self):
        """显示分盒模板的参数设置对话框（无外观选择）"""
        # 创建对话框
        dialog = self.create_base_dialog("分盒模板 - 包装参数设置", 500, 420)
        if not dialog:
            return
            
        # 创建可滚动框架
        main_frame = self.create_scrollable_frame(dialog)
        
        # 创建标题
        self.create_title_label(main_frame, "分盒模板参数设置")
        
        # 创建包装类型选择
        self.create_radio_group(
            main_frame, "包装类型选择",
            [("有小箱（三级包装）", "有小箱"), ("无小箱（二级包装）", "无小箱")],
            "has_small_box", "有小箱", self.on_small_box_choice_changed
        )
        
        # 创建盒标类型选择
        self.create_radio_group(
            main_frame, "盒标类型选择",
            [("无盒标", "无盒标"), ("有盒标", "有盒标")],
            "has_box_label", "无盒标"
        )
        
        # 创建参数输入区域
        input_configs = [
            {
                'label': '张/盒:',
                'var_name': 'pieces_per_box',
                'default_value': self.main_app.current_data.get('张/盒', ''),
                'width': 15
            },
            {
                'label': '盒/小箱:',
                'var_name': 'boxes_per_small_box',
                'default_value': '',
                'width': 15
            },
            {
                'label': '小箱/大箱:',
                'var_name': 'small_boxes_per_large_box',
                'default_value': '',
                'width': 15
            },
            {
                'label': '中文名称:',
                'var_name': 'chinese_name',
                'default_value': '',
                'width': 15
            },
            {
                'label': '序列号字体大小:',
                'var_name': 'serial_font_size',
                'default_value': '10',
                'width': 15,
                'help_text': '(建议范围: 6-14, 序列号较长时可调小)'
            }
        ]
        
        params_frame = self.create_input_section(main_frame, "包装参数", input_configs)
        self.params_frame = params_frame  # 保存引用以便动态更新标签
        
        # 创建标签模版选择
        self.create_radio_group(
            main_frame, "标签模版选择",
            [("无纸卡备注", "无纸卡备注"), ("有纸卡备注", "有纸卡备注")],
            "template_type", "无纸卡备注"
        )
        
        # 创建当前数据显示区域
        self.create_data_display_section(main_frame)
        
        # 创建按钮区域
        self.create_button_section(main_frame, self.confirm_parameters)
        
        # 自适应大小和居中显示
        self.auto_resize_and_center_dialog(main_frame)
        
        # 设置参数标签引用以便动态更新
        self.second_param_label = None
        self.third_param_label = None
        self.third_param_entry = None
        
        for widget in params_frame.winfo_children():
            if isinstance(widget, ttk.Label) and widget.cget("text") == "盒/小箱:":
                self.second_param_label = widget
            elif isinstance(widget, ttk.Label) and widget.cget("text") == "小箱/大箱:":
                self.third_param_label = widget
            elif isinstance(widget, ttk.Entry) and widget.cget("textvariable") == str(self.get_var('small_boxes_per_large_box')):
                self.third_param_entry = widget
        
        # 初始化显示状态
        self.on_small_box_choice_changed()

    def on_small_box_choice_changed(self):
        """处理小箱选择变化"""
        if self.second_param_label and self.third_param_label and self.third_param_entry:
            has_small_box = self.get_var_value("has_small_box") == "有小箱"
            if has_small_box:
                self.second_param_label.config(text="盒/小箱:")
                self.third_param_label.grid()
                self.third_param_entry.grid()
            else:
                self.second_param_label.config(text="盒/箱:")
                self.third_param_label.grid_remove()
                self.third_param_entry.grid_remove()
    
    def confirm_parameters(self):
        """确认分盒模板参数并生成PDF"""
        try:
            # 获取和验证参数
            pieces_per_box = self.validate_pieces_per_box()
            boxes_per_small_box = self.validate_integer_input('boxes_per_small_box', '盒/小箱', 1)
            
            has_small_box = self.get_var_value("has_small_box") == "有小箱"
            
            if has_small_box:
                small_boxes_per_large_box = self.validate_integer_input('small_boxes_per_large_box', '小箱/大箱', 1)
            else:
                small_boxes_per_large_box = 1  # 无小箱时设置为1
                
            chinese_name = self.validate_required_string('chinese_name', '中文名称')
            serial_font_size = self.validate_font_size()
            
            # 设置参数
            self.main_app.packaging_params = {
                "张/盒": pieces_per_box,
                "盒/小箱": boxes_per_small_box,
                "小箱/大箱": small_boxes_per_large_box,
                "选择外观": "外观一",  # 分盒模板固定为外观一
                "标签模版": self.get_var_value("template_type"),
                "中文名称": chinese_name,
                "序列号字体大小": serial_font_size,
                "是否有小箱": has_small_box,
                "是否有盒标": self.get_var_value("has_box_label") == "有盒标",
            }

            self.dialog.destroy()
            self.main_app.generate_multi_level_pdfs()
            
        except ValueError as e:
            self.show_error(str(e))

    # ==================== 实现抽象方法 ====================
    
    def get_template_specific_inputs(self) -> List[Dict]:
        """获取分盒模板特定的输入配置"""
        return [
            {'label': '张/盒:', 'var_name': 'pieces_per_box', 'default_value': '', 'width': 15},
            {'label': '盒/小箱:', 'var_name': 'boxes_per_small_box', 'default_value': '', 'width': 15},
            {'label': '小箱/大箱:', 'var_name': 'small_boxes_per_large_box', 'default_value': '', 'width': 15},
            {'label': '中文名称:', 'var_name': 'chinese_name', 'default_value': '', 'width': 15},
            {'label': '序列号字体大小:', 'var_name': 'serial_font_size', 'default_value': '10', 'width': 15}
        ]


# 全局变量用于单例模式
split_box_ui_dialog = None


def get_split_box_ui_dialog(main_app):
    """获取分盒模板UI对话框实例（单例模式）"""
    global split_box_ui_dialog
    if split_box_ui_dialog is None:
        split_box_ui_dialog = SplitBoxUIDialog(main_app)
    return split_box_ui_dialog