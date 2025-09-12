"""
通用PDF渲染基类
抽取所有模板共同的PDF绘制逻辑，减少代码重复
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
from reportlab.lib.units import mm
from abc import ABC, abstractmethod
from typing import List, Dict, Any

# 导入工具类
from src.utils.font_manager import font_manager
from src.utils.text_processor import text_processor


class BaseRenderer(ABC):
    """通用PDF渲染基类"""
    
    def __init__(self):
        """初始化基类渲染器"""
        self.renderer_type = "base"
    
    # ==================== 通用外观渲染方法 ====================
    
    def render_centered_title(self, c, width, top_text, top_text_y, serial_number, serial_number_y, font_size=22):
        """渲染居中的标题（两行版本：标题+序列号）"""
        return self.render_two_line_label(c, width, top_text, serial_number, top_text_y, serial_number_y, font_size)
    
    def render_centered_chinese_title(self, c, width, height, title_text, font_size=14):
        """渲染居中的中文标题文本（支持自动换行）"""
        clean_title = text_processor.clean_text_for_font(title_text)
        font_manager.set_best_font(c, font_size, bold=True)
        
        # 计算页面中央位置
        center_x = width / 2
        center_y = height / 2
        
        # 计算最大文本宽度（页面宽度的80%）
        max_width = width * 0.8
        
        # 检查文本宽度并进行换行处理
        current_font_name = font_manager.get_chinese_font_name()
        text_width = c.stringWidth(clean_title, current_font_name, font_size)
        
        if text_width > max_width:
            # 需要换行
            title_lines = text_processor.wrap_text_to_fit(c, clean_title, max_width, current_font_name, font_size)
            
            if len(title_lines) > 1:
                # 多行：计算行高和总高度
                line_height = font_size * 1.2
                total_height = (len(title_lines) - 1) * line_height
                
                # 计算起始Y位置（垂直居中）
                start_y = center_y + total_height / 2
                
                # 绘制每一行，居中显示
                for i, line in enumerate(title_lines):
                    line_y = start_y - i * line_height
                    c.drawCentredString(center_x, line_y, line)
            else:
                # 单行但需要处理
                c.drawCentredString(center_x, center_y, title_lines[0])
        else:
            # 单行：直接绘制
            c.drawCentredString(center_x, center_y, clean_title)
    
    def render_two_line_label(self, c, width, title_text, serial_text, title_y, serial_y, title_font_size=14):
        """渲染两行标签：标题+序列号"""
        clean_title = text_processor.clean_text_for_font(title_text)
        font_manager.set_best_font(c, title_font_size, bold=True)
        
        # 计算最大文本宽度（页面宽度的80%）
        max_width = width * 0.8
        
        # 检查文本宽度并进行换行处理
        current_font_name = font_manager.get_chinese_font_name()
        text_width = c.stringWidth(clean_title, current_font_name, title_font_size)
        
        if text_width > max_width:
            # 需要换行
            title_lines = text_processor.wrap_text_to_fit(c, clean_title, max_width, current_font_name, title_font_size)
            
            if len(title_lines) > 1:
                # 多行：计算垂直居中位置
                line_height = 16
                total_height = (len(title_lines) - 1) * line_height
                start_y = title_y + total_height / 2
                
                # 绘制每一行，居中显示
                for i, line in enumerate(title_lines):
                    line_y = start_y - i * line_height
                    c.drawCentredString(width / 2, line_y, line)
            else:
                # 单行但需要处理
                c.drawCentredString(width / 2, title_y, title_lines[0])
        else:
            # 单行：直接绘制
            c.drawCentredString(width / 2, title_y, clean_title)
        
        # 绘制序列号
        c.drawCentredString(width / 2, serial_y, serial_text)
    
    # ==================== 通用表格绘制方法 ====================
    
    def draw_base_table_structure(self, c, width, height, rows_config: List[Dict], 
                                 label_col_ratio=1/3, margin=5*mm):
        """
        绘制基础表格结构
        
        Args:
            c: canvas对象
            width, height: 表格区域尺寸
            rows_config: 行配置列表，每个元素包含 {'height_ratio': float} 
            label_col_ratio: 标签列宽度比例
            margin: 边距
        
        Returns:
            dict: 包含表格位置信息的字典
        """
        # 表格尺寸和位置
        table_width = width - 2 * margin
        table_height = height - 2 * margin
        table_x = margin
        table_y = margin
        
        # 列宽
        label_col_width = table_width * label_col_ratio
        data_col_width = table_width * (1 - label_col_ratio)
        
        # 绘制表格外边框
        c.setStrokeColor(CMYKColor(0, 0, 0, 1))
        c.setLineWidth(1)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 计算行高度
        total_ratio = sum(row['height_ratio'] for row in rows_config)
        actual_heights = [table_height * row['height_ratio'] / total_ratio for row in rows_config]
        
        # 计算各行的Y坐标
        row_positions = []
        current_y = table_y
        for height_val in actual_heights:
            row_positions.append(current_y)
            current_y += height_val
        
        # 绘制横线
        for i in range(1, len(rows_config)):
            y = row_positions[i]
            c.line(table_x, y, table_x + table_width, y)
        
        # 绘制竖线
        col_x = table_x + label_col_width
        c.line(col_x, table_y, col_x, table_y + table_height)
        
        return {
            'table_x': table_x,
            'table_y': table_y,
            'table_width': table_width,
            'table_height': table_height,
            'label_col_width': label_col_width,
            'data_col_width': data_col_width,
            'row_positions': row_positions,
            'row_heights': actual_heights,
            'label_center_x': table_x + label_col_width / 2,
            'data_center_x': col_x + data_col_width / 2
        }
    
    def draw_table_cell_text(self, c, text: str, center_x: float, center_y: float, 
                           font_size=10, max_width=None, font_name=None):
        """
        在表格单元格中绘制文本（支持自动换行）
        
        Args:
            c: canvas对象
            text: 要绘制的文本
            center_x, center_y: 单元格中心位置
            font_size: 字体大小
            max_width: 最大宽度，如果超过则自动换行
            font_name: 字体名称
        """
        clean_text = text_processor.clean_text_for_font(text)
        font_manager.set_best_font(c, font_size, bold=True)
        
        if max_width is None:
            # 单行绘制
            c.drawCentredString(center_x, center_y - font_size/3, clean_text)
        else:
            # 支持换行的绘制
            current_font_name = font_name or font_manager.get_chinese_font_name()
            text_lines = text_processor.wrap_text_to_fit(c, clean_text, max_width, current_font_name, font_size)
            
            if len(text_lines) > 1:
                # 多行：调整字体大小并垂直居中
                smaller_font_size = max(6, font_size - 2)
                font_manager.set_best_font(c, smaller_font_size, bold=True)
                text_lines = text_processor.wrap_text_to_fit(c, clean_text, max_width, current_font_name, smaller_font_size)
                
                line_height = smaller_font_size + 2
                total_height = (len(text_lines) - 1) * line_height
                start_y = center_y + total_height / 2 - smaller_font_size/3
                
                for i, line in enumerate(text_lines):
                    line_y = start_y - i * line_height
                    c.drawCentredString(center_x, line_y, line)
                
                # 恢复字体大小
                font_manager.set_best_font(c, font_size, bold=True)
            else:
                c.drawCentredString(center_x, center_y - font_size/3, text_lines[0])
    
    def draw_standard_box_table(self, c, width, height, theme_text: str, pieces: int,
                              serial_range: str, carton_no: str, remark_text: str,
                              serial_font_size=10, has_paper_card=True):
        """
        绘制标准盒标表格（5行或4行版本）
        
        Args:
            has_paper_card: True=5行版本(有纸卡备注), False=4行版本(无纸卡备注)
        """
        if has_paper_card:
            # 5行版本：Item, Theme, Quantity(双倍), Carton No, Remark
            rows_config = [
                {'height_ratio': 1},    # Remark
                {'height_ratio': 1},    # Carton No
                {'height_ratio': 2},    # Quantity (双倍高度)
                {'height_ratio': 1},    # Theme
                {'height_ratio': 1}     # Item
            ]
        else:
            # 4行版本：Item, Quantity(双倍), Carton No, Remark
            rows_config = [
                {'height_ratio': 1},    # Remark
                {'height_ratio': 1},    # Carton No
                {'height_ratio': 2},    # Quantity (双倍高度)
                {'height_ratio': 1}     # Item
            ]
        
        # 绘制表格结构
        table_info = self.draw_base_table_structure(c, width, height, rows_config)
        
        # 绘制Quantity行的分隔线
        quantity_row_idx = 2
        quantity_split_y = table_info['row_positions'][quantity_row_idx] + table_info['row_heights'][quantity_row_idx] / 2
        c.line(table_info['table_x'] + table_info['label_col_width'], quantity_split_y, 
               table_info['table_x'] + table_info['table_width'], quantity_split_y)
        
        # 设置字体
        font_manager.set_best_font(c, 10, bold=True)
        text_offset = 10 / 3
        
        if has_paper_card:
            # 5行版本的内容
            # Item行 (第5行，从上往下)
            item_y = table_info['row_positions'][4] + table_info['row_heights'][4]/2 - text_offset
            c.drawCentredString(table_info['label_center_x'], item_y, "Item:")
            c.drawCentredString(table_info['data_center_x'], item_y, "Paper Cards")
            
            # Theme行 (第4行)
            theme_y = table_info['row_positions'][3] + table_info['row_heights'][3]/2 - text_offset
            c.drawCentredString(table_info['label_center_x'], theme_y, "Theme:")
            max_theme_width = table_info['data_col_width'] - 4*mm
            self.draw_table_cell_text(c, theme_text, table_info['data_center_x'], 
                                    table_info['row_positions'][3] + table_info['row_heights'][3]/2, 
                                    10, max_theme_width)
        else:
            # 4行版本的内容
            # Item行 (第4行，从上往下) - 显示主题内容
            item_y = table_info['row_positions'][3] + table_info['row_heights'][3]/2 - text_offset
            c.drawCentredString(table_info['label_center_x'], item_y, "Item:")
            max_theme_width = table_info['data_col_width'] - 4*mm
            self.draw_table_cell_text(c, theme_text, table_info['data_center_x'], 
                                    table_info['row_positions'][3] + table_info['row_heights'][3]/2, 
                                    10, max_theme_width)
        
        # Quantity行 (第3行，双倍高度)
        quantity_label_y = table_info['row_positions'][2] + table_info['row_heights'][2]/2 - text_offset
        c.drawCentredString(table_info['label_center_x'], quantity_label_y, "Quantity:")
        
        # 上层：票数
        upper_y = table_info['row_positions'][2] + table_info['row_heights'][2] * 3/4 - text_offset
        pcs_text = f"{pieces}PCS"
        c.drawCentredString(table_info['data_center_x'], upper_y, pcs_text)
        
        # 下层：序列号范围
        font_manager.set_best_font(c, serial_font_size, bold=True)
        lower_y = table_info['row_positions'][2] + table_info['row_heights'][2]/4 - serial_font_size/3
        clean_serial_range = text_processor.clean_text_for_font(serial_range)
        c.drawCentredString(table_info['data_center_x'], lower_y, clean_serial_range)
        font_manager.set_best_font(c, 10, bold=True)  # 恢复默认字体大小
        
        # Carton No行
        carton_y = table_info['row_positions'][1] + table_info['row_heights'][1]/2 - text_offset
        c.drawCentredString(table_info['label_center_x'], carton_y, "Carton No:")
        c.drawCentredString(table_info['data_center_x'], carton_y, carton_no)
        
        # Remark行
        remark_y = table_info['row_positions'][0] + table_info['row_heights'][0]/2 - text_offset
        clean_remark_text = text_processor.clean_text_for_font(remark_text)
        c.drawCentredString(table_info['label_center_x'], remark_y, "Remark:")
        c.drawCentredString(table_info['data_center_x'], remark_y, clean_remark_text)
    
    def render_empty_box_label(self, c, width, height, chinese_name: str, remark_text: str, has_paper_card=True):
        """渲染空箱标签"""
        if has_paper_card:
            # 5行版本：Item显示"Paper Cards"，Theme显示中文名称
            rows_config = [
                {'height_ratio': 1},    # Remark
                {'height_ratio': 1},    # Carton No (空白)
                {'height_ratio': 2},    # Quantity (空白)
                {'height_ratio': 1},    # Theme
                {'height_ratio': 1}     # Item
            ]
        else:
            # 4行版本：Item显示中文名称
            rows_config = [
                {'height_ratio': 1},    # Remark
                {'height_ratio': 1},    # Carton No (空白)
                {'height_ratio': 2},    # Quantity (空白)
                {'height_ratio': 1}     # Item
            ]
        
        # 绘制表格结构
        table_info = self.draw_base_table_structure(c, width, height, rows_config)
        
        # 绘制Quantity行的分隔线
        quantity_row_idx = 2
        quantity_split_y = table_info['row_positions'][quantity_row_idx] + table_info['row_heights'][quantity_row_idx] / 2
        c.line(table_info['table_x'] + table_info['label_col_width'], quantity_split_y, 
               table_info['table_x'] + table_info['table_width'], quantity_split_y)
        
        # 设置字体
        font_manager.set_best_font(c, 10, bold=True)
        text_offset = 10 / 3
        
        if has_paper_card:
            # 5行版本
            # Item行
            item_y = table_info['row_positions'][4] + table_info['row_heights'][4]/2 - text_offset
            c.drawCentredString(table_info['label_center_x'], item_y, "Item:")
            c.drawCentredString(table_info['data_center_x'], item_y, "Paper Cards")
            
            # Theme行
            theme_y = table_info['row_positions'][3] + table_info['row_heights'][3]/2 - text_offset
            c.drawCentredString(table_info['label_center_x'], theme_y, "Theme:")
            max_theme_width = table_info['data_col_width'] - 4*mm
            self.draw_table_cell_text(c, chinese_name, table_info['data_center_x'], 
                                    table_info['row_positions'][3] + table_info['row_heights'][3]/2, 
                                    10, max_theme_width)
        else:
            # 4行版本
            # Item行 - 显示中文名称
            item_y = table_info['row_positions'][3] + table_info['row_heights'][3]/2 - text_offset
            c.drawCentredString(table_info['label_center_x'], item_y, "Item:")
            max_theme_width = table_info['data_col_width'] - 4*mm
            self.draw_table_cell_text(c, chinese_name, table_info['data_center_x'], 
                                    table_info['row_positions'][3] + table_info['row_heights'][3]/2, 
                                    10, max_theme_width)
        
        # Quantity行 (空白)
        quantity_label_y = table_info['row_positions'][2] + table_info['row_heights'][2]/2 - text_offset
        c.drawCentredString(table_info['label_center_x'], quantity_label_y, "Quantity:")
        
        # Carton No行 (空白)
        carton_y = table_info['row_positions'][1] + table_info['row_heights'][1]/2 - text_offset
        c.drawCentredString(table_info['label_center_x'], carton_y, "Carton No:")
        
        # Remark行
        remark_y = table_info['row_positions'][0] + table_info['row_heights'][0]/2 - text_offset
        clean_remark_text = text_processor.clean_text_for_font(remark_text)
        c.drawCentredString(table_info['label_center_x'], remark_y, "Remark:")
        c.drawCentredString(table_info['data_center_x'], remark_y, clean_remark_text)
    
    # ==================== 子类调用的兼容方法 ====================
    
    def render_three_line_layout(self, c, width, page_size, game_title, ticket_count, serial_number):
        """渲染外观二：精确的三行布局格式"""
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        font_manager.set_best_font(c, 12, bold=True)
        
        # 获取页面高度
        page_height = page_size[1]
        
        # 左边距 - 统一的左边距
        left_margin = 4 * mm
        
        # 清理文本，移除不支持的字符
        clean_game_title = text_processor.clean_text_for_font(game_title)
        clean_serial_number = text_processor.clean_text_for_font(str(serial_number))
        
        # Game title: 距离顶部更大的距离，支持换行并居中显示
        game_title_y = page_height - 12 * mm  # 距离顶部12mm
        max_title_width = width - 2 * left_margin  # 留出左右边距
        
        # 将Game title文本分为标签部分和内容部分
        title_prefix = "Game title: "
        title_content = clean_game_title
        
        # 计算标签部分宽度
        used_font = font_manager.get_chinese_font_name() or "Helvetica"
        prefix_width = c.stringWidth(title_prefix, used_font, 12)
        
        # 计算内容可用宽度
        content_max_width = max_title_width - prefix_width
        
        # 对内容部分进行换行处理
        title_lines = text_processor.wrap_text_to_fit(c, title_content, content_max_width, font_manager.get_chinese_font_name(), 12)
        
        # 绘制Game title
        for i, line in enumerate(title_lines):
            current_y = game_title_y - i * 14  # 行间距14点
            if i == 0:
                # 第一行包含"Game title: "前缀，左对齐
                full_line = f"{title_prefix}{line}"
                c.drawString(left_margin, current_y, full_line)
            else:
                # 后续换行内容居中显示
                line_width = c.stringWidth(line, used_font, 12)
                center_x = (width - line_width) / 2
                c.drawString(center_x, current_y, line)
        
        # Ticket count: 左下区域
        ticket_count_y = 15 * mm  # 距离底部15mm
        ticket_text = f"Ticket count: {ticket_count}"
        c.drawString(left_margin, ticket_count_y, ticket_text)
        
        # Serial: 距离底部
        serial_y = 6 * mm  # 距离底部6mm
        serial_text = f"Serial: {clean_serial_number}"
        c.drawString(left_margin, serial_y, serial_text)

    def render_blank_game_title_page(self, c, width, height, chinese_name):
        """渲染外观2的空白首页 - 完全按照外观2格式"""
        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 设置字体 - 与外观2保持一致
        font_manager.set_best_font(c, 12, bold=True)
        
        # 左边距 - 与外观2保持一致
        left_margin = 4 * mm
        
        # 清理中文文本
        clean_chinese_name = text_processor.clean_text_for_font(chinese_name)
        
        # Game title: 距离顶部12mm
        game_title_y = height - 12 * mm
        max_title_width = width - 2 * left_margin  # 留出左右边距
        
        # 将Game title文本分为标签部分和内容部分
        title_prefix = "Game title: "
        title_content = clean_chinese_name
        
        # 计算标签部分宽度
        used_font = font_manager.get_chinese_font_name() or "Helvetica"
        prefix_width = c.stringWidth(title_prefix, used_font, 12)
        
        # 计算内容可用宽度
        content_max_width = max_title_width - prefix_width
        
        # 对内容部分进行换行处理
        title_lines = text_processor.wrap_text_to_fit(c, title_content, content_max_width, font_manager.get_chinese_font_name(), 12)
        
        # 绘制Game title
        for i, line in enumerate(title_lines):
            current_y = game_title_y - i * 14  # 行间距14点
            if i == 0:
                # 第一行包含"Game title: "前缀，左对齐
                full_line = f"{title_prefix}{line}"
                c.drawString(left_margin, current_y, full_line)
            else:
                # 后续行只包含内容，需要考虑前缀宽度的缩进
                indent_x = left_margin + prefix_width
                c.drawString(indent_x, current_y, line)
        
        # Ticket count: 空行
        ticket_count_y = 15 * mm  # 距离底部15mm
        c.drawString(left_margin, ticket_count_y, "Ticket count:")
        
        # Serial: 空行
        serial_y = 6 * mm  # 距离底部6mm
        c.drawString(left_margin, serial_y, "Serial:")
    
    def render_empty_box_label_with_paper_card(self, c, width, height, chinese_name, remark_text):
        """渲染有纸卡备注的空箱标签 - 直接调用主实现"""
        # 5行版本：Item显示"Paper Cards"，Theme显示中文名称
        rows_config = [
            {'height_ratio': 1},    # Remark
            {'height_ratio': 1},    # Carton No (空白)
            {'height_ratio': 2},    # Quantity (空白)
            {'height_ratio': 1},    # Theme
            {'height_ratio': 1}     # Item
        ]
        
        # 绘制表格结构
        table_info = self.draw_base_table_structure(c, width, height, rows_config)
        
        # 绘制Quantity行的分隔线
        quantity_row_idx = 2
        quantity_split_y = table_info['row_positions'][quantity_row_idx] + table_info['row_heights'][quantity_row_idx] / 2
        c.line(table_info['table_x'] + table_info['label_col_width'], quantity_split_y, 
               table_info['table_x'] + table_info['table_width'], quantity_split_y)
        
        # 设置字体
        font_manager.set_best_font(c, 10, bold=True)
        text_offset = 10 / 3
        
        # Item行
        item_y = table_info['row_positions'][4] + table_info['row_heights'][4]/2 - text_offset
        c.drawCentredString(table_info['label_center_x'], item_y, "Item:")
        c.drawCentredString(table_info['data_center_x'], item_y, "Paper Cards")
        
        # Theme行
        theme_y = table_info['row_positions'][3] + table_info['row_heights'][3]/2 - text_offset
        c.drawCentredString(table_info['label_center_x'], theme_y, "Theme:")
        max_theme_width = table_info['data_col_width'] - 4*mm
        self.draw_table_cell_text(c, chinese_name, table_info['data_center_x'], 
                                table_info['row_positions'][3] + table_info['row_heights'][3]/2, 
                                max_theme_width, 10)
        
        # Carton No行
        carton_y = table_info['row_positions'][1] + table_info['row_heights'][1]/2 - text_offset
        c.drawCentredString(table_info['label_center_x'], carton_y, "Carton No:")
        
        # Remark行
        remark_y = table_info['row_positions'][0] + table_info['row_heights'][0]/2 - text_offset
        clean_remark_text = text_processor.clean_text_for_font(remark_text)
        c.drawCentredString(table_info['label_center_x'], remark_y, "Remark:")
        c.drawCentredString(table_info['data_center_x'], remark_y, clean_remark_text)
        
    def render_empty_box_label_without_paper_card(self, c, width, height, chinese_name, remark_text):
        """渲染无纸卡备注的空箱标签 - 直接调用主实现"""
        # 4行版本：Item显示中文名称
        rows_config = [
            {'height_ratio': 1},    # Remark
            {'height_ratio': 1},    # Carton No (空白)
            {'height_ratio': 2},    # Quantity (空白)
            {'height_ratio': 1}     # Item
        ]
        
        # 绘制表格结构
        table_info = self.draw_base_table_structure(c, width, height, rows_config)
        
        # 绘制Quantity行的分隔线
        quantity_row_idx = 2
        quantity_split_y = table_info['row_positions'][quantity_row_idx] + table_info['row_heights'][quantity_row_idx] / 2
        c.line(table_info['table_x'] + table_info['label_col_width'], quantity_split_y, 
               table_info['table_x'] + table_info['table_width'], quantity_split_y)
        
        # 设置字体
        font_manager.set_best_font(c, 10, bold=True)
        text_offset = 10 / 3
        
        # Item行（显示中文名称）
        item_y = table_info['row_positions'][3] + table_info['row_heights'][3]/2 - text_offset
        c.drawCentredString(table_info['label_center_x'], item_y, "Item:")
        max_item_width = table_info['data_col_width'] - 4*mm
        self.draw_table_cell_text(c, chinese_name, table_info['data_center_x'], 
                                table_info['row_positions'][3] + table_info['row_heights'][3]/2, 
                                max_item_width, 10)
        
        # Carton No行
        carton_y = table_info['row_positions'][1] + table_info['row_heights'][1]/2 - text_offset
        c.drawCentredString(table_info['label_center_x'], carton_y, "Carton No:")
        
        # Remark行
        remark_y = table_info['row_positions'][0] + table_info['row_heights'][0]/2 - text_offset
        clean_remark_text = text_processor.clean_text_for_font(remark_text)
        c.drawCentredString(table_info['label_center_x'], remark_y, "Remark:")
        c.drawCentredString(table_info['data_center_x'], remark_y, clean_remark_text)
    
    # ==================== 非抽象方法 ====================
    
    def get_renderer_type(self) -> str:
        """返回渲染器类型 - 子类可以重写"""
        return self.renderer_type