"""
常规模板渲染器
基于BaseRenderer重构，专门负责常规模板的PDF绘制逻辑
"""

from typing import Dict, Any
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

from src.utils.base_renderer import BaseRenderer


class RegularRenderer(BaseRenderer):
    """常规模板渲染器 - 基于BaseRenderer重构的简化版本"""
    
    def __init__(self):
        """初始化常规渲染器"""
        super().__init__()
        self.renderer_type = "regular"
    
    # ==================== 外观渲染方法 ====================
    
    def render_appearance_one(self, c, width, top_text, serial_number, top_text_y, serial_number_y):
        """渲染外观一：简洁标准样式"""
        return self.render_centered_title(c, width, top_text, top_text_y, serial_number, serial_number_y, 22)

    def render_appearance_two(self, c, width, page_size, game_title, ticket_count, serial_number, top_y, bottom_y):
        """渲染外观二：精确的三行布局格式"""
        return self.render_three_line_layout(c, width, page_size, game_title, ticket_count, serial_number)
    
    # ==================== 表格绘制方法 ====================
    
    def draw_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                            serial_range, carton_no, remark_text, template_type="有纸卡备注", serial_font_size=10):
        """绘制小箱标表格"""
        has_paper_card = (template_type == "有纸卡备注")
        return self.draw_standard_box_table(
            c, width, height, theme_text, pieces_per_small_box,
            serial_range, carton_no, remark_text, serial_font_size, has_paper_card
        )

    def draw_large_box_table(self, c, width, height, theme_text, pieces_per_large_box,
                            serial_range, carton_no, remark_text, template_type="有纸卡备注", serial_font_size=10):
        """绘制大箱标表格"""
        has_paper_card = (template_type == "有纸卡备注")
        return self.draw_standard_box_table(
            c, width, height, theme_text, pieces_per_large_box,
            serial_range, carton_no, remark_text, serial_font_size, has_paper_card
        )
    
    # ==================== 空白标签渲染方法 ====================
    
    def render_empty_box_label(self, c, width, height, chinese_name, remark_text, has_paper_card=True):
        """渲染空箱标签 - 用于小箱标和大箱标的第一页（有纸卡备注）"""
        if has_paper_card:
            return self.render_empty_box_label_with_paper_card(c, width, height, chinese_name, remark_text)
        else:
            return self.render_empty_box_label_without_paper_card(c, width, height, chinese_name, remark_text)

    def render_empty_box_label_no_paper_card(self, c, width, height, chinese_name, remark_text):
        """渲染空箱标签 - 用于小箱标和大箱标的第一页（无纸卡备注）"""
        return self.render_empty_box_label_without_paper_card(c, width, height, chinese_name, remark_text)

    # ==================== 首页渲染方法 ====================
    
    def render_blank_first_page(self, c, width, height, chinese_name):
        """渲染常规模版盒标的空白首页 - 仅显示中文标题"""
        return self.render_centered_chinese_title(c, width, height, chinese_name, font_size=22)

    def render_blank_first_page_appearance_two(self, c, width, height, chinese_name):
        """渲染常规模版外观2的空白首页 - 完全按照外观2格式"""
        return self.render_blank_game_title_page(c, width, height, chinese_name)


# 创建全局实例供regular模板使用  
regular_renderer = RegularRenderer()