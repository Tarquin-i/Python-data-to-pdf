"""
分盒模板渲染器
基于BaseRenderer重构，负责分盒模板的所有PDF绘制和渲染逻辑
"""

from typing import Dict, Any
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

from src.utils.base_renderer import BaseRenderer


class SplitBoxRenderer(BaseRenderer):
    """分盒模板渲染器 - 基于BaseRenderer重构的简化版本"""
    
    def __init__(self):
        """初始化分盒渲染器"""
        super().__init__()
        self.renderer_type = "split_box"
    
    # ==================== 外观渲染方法 ====================
    
    def render_appearance_one(self, c, width, top_text, serial_number, top_text_y, serial_number_y):
        """分盒模板盒标外观一渲染"""
        return self.render_centered_title(c, width, top_text, top_text_y, serial_number, serial_number_y, 22)
    
    # ==================== 表格绘制方法 ====================
    
    def draw_split_box_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                                       serial_range, carton_no, remark_text, has_paper_card=True, serial_font_size=10):
        """绘制分盒小箱标表格"""
        return self.draw_standard_box_table(
            c, width, height, theme_text, pieces_per_small_box,
            serial_range, carton_no, remark_text, serial_font_size, has_paper_card
        )

    def draw_split_box_small_box_table_no_paper_card(self, c, width, height, theme_text, pieces_per_small_box, 
                                                      serial_range, carton_no, remark_text, serial_font_size=10):
        """绘制分盒小箱标表格 - 无纸卡备注版本"""
        return self.draw_standard_box_table(
            c, width, height, theme_text, pieces_per_small_box,
            serial_range, carton_no, remark_text, serial_font_size, has_paper_card=False
        )

    def draw_split_box_large_box_table(self, c, width, height, theme_text, pieces_per_box,
                                       boxes_per_small_box, small_boxes_per_large_box, serial_range, 
                                       carton_no, remark_text, serial_font_size=10):
        """绘制分盒大箱标表格"""
        # 计算大箱的总张数
        pieces_per_large_box = pieces_per_box * boxes_per_small_box * small_boxes_per_large_box
        return self.draw_standard_box_table(
            c, width, height, theme_text, pieces_per_large_box,
            serial_range, carton_no, remark_text, serial_font_size, has_paper_card=True
        )
    
    def draw_split_box_large_box_table_no_paper_card(self, c, width, height, theme_text, pieces_per_box,
                                                     boxes_per_small_box, small_boxes_per_large_box, serial_range, 
                                                     carton_no, remark_text, serial_font_size=10):
        """绘制分盒大箱标表格 - 无纸卡备注版本"""
        # 计算大箱的总张数
        pieces_per_large_box = pieces_per_box * boxes_per_small_box * small_boxes_per_large_box
        return self.draw_standard_box_table(
            c, width, height, theme_text, pieces_per_large_box,
            serial_range, carton_no, remark_text, serial_font_size, has_paper_card=False
        )
    
    # ==================== 空白标签渲染方法 - 按标签类型分离 ====================
    
    def render_empty_small_box_label(self, c, width, height, chinese_name, remark_text, has_paper_card=True):
        """渲染分盒模板小箱标空白标签 - 专门用于小箱标"""
        return super().render_empty_small_box_label(c, width, height, chinese_name, remark_text, has_paper_card)

    def render_empty_large_box_label(self, c, width, height, chinese_name, remark_text, has_paper_card=True):
        """渲染分盒模板大箱标空白标签 - 专门用于大箱标"""
        return super().render_empty_large_box_label(c, width, height, chinese_name, remark_text, has_paper_card)
    
    # ==================== 向后兼容的分盒专用方法 ====================
    
    def render_split_empty_box_label(self, c, width, height, chinese_name, remark_text):
        """兼容性方法 - 分盒空箱标签（有纸卡备注），重定向到大箱标方法"""
        return self.render_empty_large_box_label(c, width, height, chinese_name, remark_text, has_paper_card=True)

    def render_split_empty_box_label_no_paper_card(self, c, width, height, chinese_name, remark_text):
        """兼容性方法 - 分盒空箱标签（无纸卡备注），重定向到大箱标方法"""
        return self.render_empty_large_box_label(c, width, height, chinese_name, remark_text, has_paper_card=False)

    # ==================== 首页渲染方法 ====================
    
    def render_blank_box_first_page(self, c, width, height, chinese_name):
        """渲染分盒模板盒标的空白首页 - 专门用于盒标"""
        return super().render_blank_box_first_page(c, width, height, chinese_name, font_size=22)
    
    def render_split_blank_first_page(self, c, width, height, chinese_name):
        """渲染分盒模版盒标的空白首页 - 仅显示中文标题"""
        return self.render_centered_chinese_title(c, width, height, chinese_name, font_size=22)

    # ==================== 兼容性包装器方法 ====================
    
    def render_blank_first_page(self, c, width, height, chinese_name):
        """兼容性包装器 - 调用分盒模板的空白首页渲染方法"""
        return self.render_split_blank_first_page(c, width, height, chinese_name)
    
    def render_empty_box_label(self, c, width, height, chinese_name, remark_text, has_paper_card=True):
        """兼容性包装器 - 调用分盒空箱标签渲染方法（有纸卡备注）"""
        if has_paper_card:
            return self.render_split_empty_box_label(c, width, height, chinese_name, remark_text)
        else:
            return self.render_split_empty_box_label_no_paper_card(c, width, height, chinese_name, remark_text)

    def render_empty_box_label_no_paper_card(self, c, width, height, chinese_name, remark_text):
        """兼容性包装器 - 调用分盒空箱标签渲染方法（无纸卡备注）"""
        return self.render_split_empty_box_label_no_paper_card(c, width, height, chinese_name, remark_text)


# 创建全局实例供split_box模板使用  
split_box_renderer = SplitBoxRenderer()