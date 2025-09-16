"""
套盒模板渲染器
负责套盒模板的所有PDF绘制和渲染逻辑，基于通用基类重构
"""

from reportlab.lib.colors import CMYKColor

# 导入基类和工具类
from src.utils.base_renderer import BaseRenderer
from src.utils.font_manager import font_manager
from src.utils.text_processor import text_processor


class NestedBoxRenderer(BaseRenderer):
    """套盒模板渲染器 - 专门负责套盒模板的PDF绘制逻辑"""

    def __init__(self):
        """初始化套盒渲染器"""
        super().__init__()
        self.renderer_type = "nested_box"

    def render_nested_appearance_one(
        self, c, width, top_text, current_number, top_text_y, serial_number_y
    ):
        """套盒模板盒标外观一渲染"""
        self.render_two_line_label(
            c, width, top_text, current_number, top_text_y, serial_number_y, 14
        )

    def render_nested_appearance_two(
        self, c, width, top_text, current_number, top_text_y, serial_number_y
    ):
        """套盒模板盒标外观二渲染"""
        # 外观二：左对齐的特殊处理
        clean_top_text = text_processor.clean_text_for_font(top_text)
        font_manager.set_best_font(c, 14, bold=True)

        max_width = width * 0.8
        title_lines = text_processor.wrap_text_to_fit(
            c,
            clean_top_text,
            max_width,
            font_manager.get_chinese_font_name(),
            14,
        )

        if len(title_lines) > 1:
            # 首行左对齐，其他行居中
            c.drawString(width * 0.1, top_text_y + 15, title_lines[0])
            for i, line in enumerate(title_lines[1:], 1):
                c.drawCentredString(width / 2, top_text_y + 15 - i * 16, line)
        else:
            c.drawString(width * 0.1, top_text_y, title_lines[0])

        # 绘制序列号
        c.drawCentredString(width / 2, serial_number_y, current_number)

    def draw_nested_small_box_table(
        self,
        c,
        width,
        height,
        theme_text,
        pieces_per_small_box,
        serial_range,
        carton_no,
        remark_text,
        serial_font_size=10,
    ):
        """绘制套盒套标表格"""
        self.draw_standard_box_table(
            c,
            width,
            height,
            theme_text,
            pieces_per_small_box,
            serial_range,
            carton_no,
            remark_text,
            serial_font_size,
            has_paper_card=True,
        )

    def draw_nested_small_box_table_no_paper_card(
        self,
        c,
        width,
        height,
        theme_text,
        pieces_per_small_box,
        serial_range,
        carton_no,
        remark_text,
        serial_font_size=10,
    ):
        """绘制套盒套标表格 - 无纸卡备注模版"""
        self.draw_standard_box_table(
            c,
            width,
            height,
            theme_text,
            pieces_per_small_box,
            serial_range,
            carton_no,
            remark_text,
            serial_font_size,
            has_paper_card=False,
        )

    def draw_nested_large_box_table(
        self,
        c,
        width,
        height,
        theme_text,
        pieces_per_large_box,
        serial_range,
        carton_no,
        remark_text,
        serial_font_size=10,
    ):
        """绘制套盒箱标表格"""
        self.draw_standard_box_table(
            c,
            width,
            height,
            theme_text,
            pieces_per_large_box,
            serial_range,
            carton_no,
            remark_text,
            serial_font_size,
            has_paper_card=True,
        )

    def draw_nested_large_box_table_no_paper_card(
        self,
        c,
        width,
        height,
        theme_text,
        pieces_per_large_box,
        serial_range,
        carton_no,
        remark_text,
        serial_font_size=10,
    ):
        """绘制套盒箱标表格 - 无纸卡备注模版"""
        self.draw_standard_box_table(
            c,
            width,
            height,
            theme_text,
            pieces_per_large_box,
            serial_range,
            carton_no,
            remark_text,
            serial_font_size,
            has_paper_card=False,
        )

    # ==================== 空白标签渲染方法 - 按标签类型分离 ====================

    def render_blank_box_first_page(self, c, width, height, chinese_name):
        """渲染套盒模板盒标的空白首页 - 专门用于盒标"""
        return super().render_blank_box_first_page(
            c, width, height, chinese_name, font_size=14
        )

    def render_empty_small_box_label(
        self, c, width, height, chinese_name, remark_text, has_paper_card=True
    ):
        """渲染套盒模板套标空白标签 - 专门用于套标（小箱标）"""
        return super().render_empty_small_box_label(
            c, width, height, chinese_name, remark_text, has_paper_card
        )

    def render_empty_large_box_label(
        self, c, width, height, chinese_name, remark_text, has_paper_card=True
    ):
        """渲染套盒模板箱标空白标签 - 专门用于箱标（大箱标）"""
        return super().render_empty_large_box_label(
            c, width, height, chinese_name, remark_text, has_paper_card
        )

    # ==================== 向后兼容方法 ====================

    def render_empty_box_label(
        self, c, width, height, chinese_name, remark_text, has_paper_card=True
    ):
        """兼容性方法 - 重定向到大箱标空白标签（套盒模板箱标）"""
        return self.render_empty_large_box_label(
            c, width, height, chinese_name, remark_text, has_paper_card
        )

    def render_empty_box_label_no_paper_card(
        self, c, width, height, chinese_name, remark_text
    ):
        """兼容性方法 - 重定向到大箱标空白标签（无纸卡版本）"""
        return self.render_empty_large_box_label(
            c, width, height, chinese_name, remark_text, has_paper_card=False
        )

    def render_blank_first_page(self, c, width, height, chinese_name):
        """渲染套盒模板盒标的空白首页"""
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        self.render_centered_chinese_title(c, width, height, chinese_name, 14)

    # ==================== 实现抽象方法 ====================

    def get_renderer_type(self) -> str:
        """返回渲染器类型"""
        return "nested_box"


# 创建全局实例供nested_box模板使用
nested_box_renderer = NestedBoxRenderer()
