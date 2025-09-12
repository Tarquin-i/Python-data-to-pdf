"""
PDF生成器

使用ReportLab生成PDF文档
"""

from reportlab.lib.pagesizes import A4, LETTER, A5
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import CMYKColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path
from typing import List, Dict, Any
import math
import os
import platform
import sys

# 添加utils目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "utils"))
from excel_data_extractor import ExcelDataExtractor


class PDFGenerator:
    
    """
    PDF生成器类
    """

    def __init__(self, max_pages_per_file: int = 100):
        """
        初始化PDF生成器
        
        Args:
            max_pages_per_file: 每个PDF文件的最大页数限制
        """
        self.page_size = (90 * mm, 50 * mm)  # 90mm x 50mm标签尺寸
        self.margin = 2 * mm
        self.font_name = "MicrosoftYaHei"  # 使用微软雅黑作为默认字体
        self.font_size = 8
        self.chinese_font_name = "MicrosoftYaHei"  # 统一使用微软雅黑
        self.max_pages_per_file = max_pages_per_file
        self._register_chinese_font()
        
    def _extract_excel_data_by_keywords(self, excel_file_path: str) -> Dict[str, Any]:
        """
        使用关键字从Excel文件中提取数据
        
        Args:
            excel_file_path: Excel文件路径
            
        Returns:
            提取的数据字典
        """
        try:
            extractor = ExcelDataExtractor(excel_file_path)
            
            # 配置要提取的数据 - 常规模板
            keyword_config = {
                '开始号': {
                    'keyword': '开始号',
                    'direction': 'down'  # 开始号标签下方是数据 (B9->B10)
                },
                '标签名称': {
                    'keyword': '标签名称',  # 查找包含"标签名称"四个字的单元格
                    'direction': 'right'  # 标签名称右侧是数据 (G10->H10)
                },
                '结束号': {
                    'keyword': '结束号', 
                    'direction': 'down'  # 结束号标签下方是数据
                },
                '客户编码': {
                    'keyword': '客户名称编码',
                    'direction': 'down'  # 客户名称编码标签下方是数据
                },
                '主题': {
                    'keyword': '主题',
                    'direction': 'down'  # 主题标签下方是数据
                },
                '总张数': {
                    'keyword': '总张数',
                    'direction': 'down'  # 总张数标签下方是数据
                }
            }
            
            # 提取数据
            extracted_data = extractor.extract_data_by_keywords(keyword_config)
            
            return extracted_data
            
        except Exception as e:
            print(f"⚠️ 关键字提取失败: {e}")
            # 返回空字典，让调用者使用默认值
            return {}
    
    def _wrap_text_to_fit(self, canvas_obj, text, max_width, font_size):
        """
        将文本按单词换行以适应指定宽度
        
        Args:
            canvas_obj: ReportLab画布对象
            text: 要换行的文本
            max_width: 最大宽度
            font_size: 字体大小
            
        Returns:
            换行后的文本行列表
        """
        # 使用最佳字体进行宽度计算
        used_font = self._set_best_font(canvas_obj, font_size, bold=True)
        
        # 按单词分割文本
        words = text.split()
        if not words:
            return [text]
        
        lines = []
        current_line = ""
        
        for word in words:
            # 测试加上新单词后的宽度
            test_line = current_line + " " + word if current_line else word
            text_width = canvas_obj.stringWidth(test_line, used_font, font_size)
            
            if text_width <= max_width:
                # 如果宽度允许，加上这个单词
                current_line = test_line
            else:
                # 如果宽度超出，开始新行
                if current_line:
                    lines.append(current_line)
                    current_line = word
                else:
                    # 单个单词就超宽，强制使用
                    lines.append(word)
        
        # 添加最后一行
        if current_line:
            lines.append(current_line)
        
        return lines if lines else [text]
    
    def _set_best_font(self, canvas_obj, font_size, bold=True):
        """
        设置最适合的字体，支持中文显示
        
        Args:
            canvas_obj: ReportLab canvas对象
            font_size: 字体大小
            bold: 是否使用粗体
            
        Returns:
            实际使用的字体名称
        """
        # 优先使用已注册的中文字体，确保中文标点符号正确显示
        font_candidates = []
        
        if self.chinese_font_name and self.chinese_font_name != "Helvetica":
            # 如果有中文字体，优先使用
            font_candidates = [self.chinese_font_name]
        else:
            # 回退到标准字体
            if bold:
                font_candidates = [
                    "Helvetica-Bold",    # 标准英文粗体
                    "Times-Bold",        # Times粗体
                    "Courier-Bold",      # 等宽粗体
                    "Helvetica",         # 退回到普通字体
                    "Times-Roman",       # Times普通
                    "Courier",           # 等宽普通
                ]
            else:
                font_candidates = [
                    "Helvetica",         # 标准英文字体
                    "Times-Roman",       # Times字体
                    "Courier",           # 等宽字体
                ]
        
        # 尝试设置字体
        used_font = self.font_name  # 使用实例的字体名称作为默认
        for font_name in font_candidates:
            try:
                canvas_obj.setFont(font_name, font_size)
                used_font = font_name
                break
            except Exception:
                continue
        
        return used_font
    
    def _clean_text_for_font(self, text):
        """
        清理文本，移除可能导致渲染问题的字符
        使用微软雅黑字体时，保留中文标点符号以确保原汁原味的显示
        
        Args:
            text: 原始文本
            
        Returns:
            清理后的文本
        """
        if not text:
            return text
        
        # 转换为字符串并去除首尾空白
        text = str(text).strip()
        
        # 统一的文本清理策略：无论使用什么字体都转换中文标点符号为可渲染的等价字符
        replacements = {
            '\ufffd': '',  # Unicode替换字符
            '\u2019': "'",  # 右单引号 → 英文单引号
            '\u2018': "'",  # 左单引号 → 英文单引号
            '\u201c': '"',  # 左双引号 → 英文双引号
            '\u201d': '"',  # 右双引号 → 英文双引号
            '\u2013': '-',  # en dash → 连字符
            '\u2014': '-',  # em dash → 连字符
            '\u00a0': ' ',  # 不间断空格 → 普通空格
            # 中文标点符号转换为ASCII等价字符，确保在任何字体下都能正常显示
            '！': '!',      # 中文感叹号 → 英文感叹号
            '？': '?',      # 中文问号 → 英文问号
            '，': ',',      # 中文逗号 → 英文逗号
            '。': '.',      # 中文句号 → 英文句号
            '；': ';',      # 中文分号 → 英文分号
            '：': ':',      # 中文冒号 → 英文冒号
            '（': '(',      # 中文左括号 → 英文左括号
            '）': ')',      # 中文右括号 → 英文右括号
            '【': '[',      # 中文左方括号 → 英文左方括号
            '】': ']',      # 中文右方括号 → 英文右方括号
            '「': '"',      # 中文左引号 → 英文双引号
            '」': '"',      # 中文右引号 → 英文双引号
            '《': '<',      # 中文左书名号 → 小于号
            '》': '>',      # 中文右书名号 → 大于号
            '—': '-',      # 中文破折号 → 英文连字符
            '…': '...',    # 中文省略号 → 英文省略号
        }
        
        for old_char, new_char in replacements.items():
            text = text.replace(old_char, new_char)
        
        # 移除明确的问题字符，但保留转换后的标点符号
        cleaned_chars = []
        for char in text:
            char_code = ord(char)
            
            # 只移除明确会导致渲染问题的字符
            if char_code == 0xFFFD:  # Unicode替换字符（黑方块的根源）
                continue
            elif char_code in [0xFEFF, 0x200B, 0x200C, 0x200D]:  # BOM和零宽字符
                continue
            elif (0 <= char_code <= 31) and char not in ['\n', '\r', '\t']:  # 控制字符但保留换行
                continue
            else:
                # 保留所有其他字符（包括中文字符和转换后的ASCII标点）
                cleaned_chars.append(char)
        
        result = ''.join(cleaned_chars)
        return result if result else text  # 如果清理后为空，返回原文本

    def set_page_size(self, size: str):
        """
        设置页面大小

        Args:
            size: 页面大小 ('A4', 'LETTER', 'A5')
        """
        size_map = {"A4": A4, "LETTER": LETTER, "A5": A5}

        if size.upper() in size_map:
            self.page_size = size_map[size.upper()]
        else:
            raise ValueError(f"不支持的页面大小: {size}")

    def generate_from_template(self, template, data: Dict[str, Any], output_path: str):
        """
        根据模板和数据生成PDF

        Args:
            template: 模板对象
            data: 数据字典
            output_path: 输出文件路径
        """
        if not output_path:
            raise ValueError("输出路径不能为空")

        # 确保输出目录存在
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # 创建PDF画布
        c = canvas.Canvas(str(output_file), pagesize=self.page_size)

        try:
            # 获取页面尺寸
            width, height = self.page_size

            # 设置基本字体
            c.setFont(self.font_name, self.font_size)

            # 如果有模板，使用模板渲染
            if template and hasattr(template, "render"):
                template.render(c, data, width, height)
            else:
                # 使用默认布局
                self._render_default_layout(c, data, width, height)

            # 保存PDF
            c.save()

        except Exception as e:
            raise Exception(f"生成PDF失败: {e}")

    def batch_generate(
        self, template, data_list: List[Dict[str, Any]], output_dir: str
    ) -> List[str]:
        """
        批量生成PDF

        Args:
            template: 模板对象
            data_list: 数据列表
            output_dir: 输出目录

        Returns:
            生成的文件路径列表
        """
        if not data_list:
            return []

        output_files = []
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        for i, data in enumerate(data_list):
            # 生成文件名
            filename = f"label_{i+1:03d}.pdf"
            if "name" in data:
                filename = f"label_{data['name']}_{i+1:03d}.pdf"

            file_path = output_path / filename

            try:
                self.generate_from_template(template, data, str(file_path))
                output_files.append(str(file_path))
            except Exception as e:
                print(f"生成第{i+1}个PDF失败: {e}")
                continue

        return output_files

    def create_multi_level_pdfs(
        self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None
    ) -> Dict[str, str]:
        """
        创建多级标签PDF

        Args:
            data: Excel数据
            params: 用户参数 (张/盒, 盒/小箱, 小箱/大箱, 选择外观)
            output_dir: 输出目录

        Returns:
            生成的文件路径字典
        """
        # 计算数量 - 三级结构：张→盒→小箱→大箱
        total_pieces = int(float(data["总张数"]))  # 处理Excel的float值
        pieces_per_box = int(params["张/盒"])  # 用户输入
        boxes_per_small_box = int(params["盒/小箱"])  # 用户输入
        small_boxes_per_large_box = int(params["小箱/大箱"])  # 用户输入

        # 计算各级数量 (使用向上取整处理余数)
        total_boxes = math.ceil(total_pieces / pieces_per_box)  # 总盒数
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)  # 总小箱数
        total_large_boxes = math.ceil(
            total_small_boxes / small_boxes_per_large_box
        )  # 总大箱数

        # 计算余数信息 - 基于新的三级结构
        remaining_pieces_in_last_box = total_pieces % pieces_per_box
        remaining_boxes_in_last_small_box = total_boxes % boxes_per_small_box
        remaining_small_boxes_in_last_large_box = (
            total_small_boxes % small_boxes_per_large_box
        )

        # 记录余数信息
        remainder_info = {
            "last_box_pieces": (
                remaining_pieces_in_last_box
                if remaining_pieces_in_last_box > 0
                else pieces_per_box
            ),
            "last_small_box_boxes": (
                remaining_boxes_in_last_small_box
                if remaining_boxes_in_last_small_box > 0
                else boxes_per_small_box
            ),
            "last_large_box_small_boxes": (
                remaining_small_boxes_in_last_large_box
                if remaining_small_boxes_in_last_large_box > 0
                else small_boxes_per_large_box
            ),
            "total_boxes": total_boxes,
        }

        # 创建输出目录 - 清理文件名中的特殊字符
        clean_theme = data['主题'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        generated_files = {}

        # 生成盒标 (只生成用户选择的外观)
        selected_appearance = params["选择外观"]
        box_label_path = (
            full_output_dir
            / f"{data['客户编码']}+{clean_theme}+盒标+{selected_appearance}.pdf"
        )

        self._create_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
        generated_files["盒标"] = str(box_label_path)

        # 生成小箱标
        small_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+小箱标.pdf"
        )
        self._create_small_box_label(
            data, params, str(small_box_path), total_small_boxes, remainder_info, excel_file_path
        )
        generated_files["小箱标"] = str(small_box_path)

        # 生成大箱标
        large_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+大箱标.pdf"
        )
        self._create_large_box_label(
            data,
            params,
            str(large_box_path),
            total_large_boxes,
            total_small_boxes,
            remainder_info,
            excel_file_path
        )
        generated_files["大箱标"] = str(large_box_path)

        return generated_files

    def create_fenhe_multi_level_pdfs(
        self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None
    ) -> Dict[str, str]:
        """
        创建分盒模板的多级标签PDF

        Args:
            data: Excel数据
            params: 用户参数 (张/盒, 盒/小箱, 小箱/大箱, 选择外观)
            output_dir: 输出目录

        Returns:
            生成的文件路径字典
        """
        # 计算数量 - 三级结构：张→盒→小箱→大箱
        total_pieces = int(float(data["总张数"]))  # 处理Excel的float值
        pieces_per_box = int(params["张/盒"])  # 用户输入
        boxes_per_small_box = int(params["盒/小箱"])  # 用户输入
        small_boxes_per_large_box = int(params["小箱/大箱"])  # 用户输入

        # 计算各级数量 (使用向上取整处理余数)
        total_boxes = math.ceil(total_pieces / pieces_per_box)  # 总盒数
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)  # 总小箱数
        total_large_boxes = math.ceil(
            total_small_boxes / small_boxes_per_large_box
        )  # 总大箱数

        # 计算余数信息 - 基于新的三级结构
        remaining_pieces_in_last_box = total_pieces % pieces_per_box
        remaining_boxes_in_last_small_box = total_boxes % boxes_per_small_box
        remaining_small_boxes_in_last_large_box = (
            total_small_boxes % small_boxes_per_large_box
        )

        # 记录余数信息
        remainder_info = {
            "last_box_pieces": (
                remaining_pieces_in_last_box
                if remaining_pieces_in_last_box > 0
                else pieces_per_box
            ),
            "last_small_box_boxes": (
                remaining_boxes_in_last_small_box
                if remaining_boxes_in_last_small_box > 0
                else boxes_per_small_box
            ),
            "last_large_box_small_boxes": (
                remaining_small_boxes_in_last_large_box
                if remaining_small_boxes_in_last_large_box > 0
                else small_boxes_per_large_box
            ),
            "total_boxes": total_boxes,
        }

        # 创建输出目录 - 清理文件名中的特殊字符
        clean_theme = data['主题'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        generated_files = {}

        # 生成分盒模板的盒标 (特殊序列号逻辑)
        selected_appearance = params["选择外观"]
        box_label_path = (
            full_output_dir
            / f"{data['客户编码']}+{clean_theme}+分盒盒标+{selected_appearance}.pdf"
        )

        self._create_fenhe_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
        generated_files["盒标"] = str(box_label_path)

        # 生成分盒模板专用的小箱标
        small_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+分盒小箱标.pdf"
        )
        self._create_fenhe_small_box_label(
            data, params, str(small_box_path), total_small_boxes, remainder_info, excel_file_path
        )
        generated_files["小箱标"] = str(small_box_path)

        # 生成分盒模板专用的大箱标 (显示序列号范围)
        large_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+分盒大箱标.pdf"
        )
        self._create_fenhe_large_box_label(
            data,
            params,
            str(large_box_path),
            total_large_boxes,
            excel_file_path
        )
        generated_files["大箱标"] = str(large_box_path)

        return generated_files

    def _create_box_label(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None
    ):
        """创建盒标 - 支持分页限制的多页PDF"""
        # 计算总盒数
        total_pieces = int(float(data["总张数"]))  # 处理Excel的float值
        pieces_per_box = int(params["张/盒"])
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 获取盒标内容 - 只使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用关键字提取
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        top_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        print(f"✅ 常规模板使用关键字提取: 标签名称='{top_text}', 开始号='{base_number}'")
        
        # 计算需要的文件数量
        total_files = math.ceil(total_boxes / self.max_pages_per_file)
        
        # 分页生成多个PDF文件
        for file_num in range(total_files):
            # 计算当前文件的页数范围
            start_box = file_num * self.max_pages_per_file + 1
            end_box = min((file_num + 1) * self.max_pages_per_file, total_boxes)
            
            # 生成文件名
            if total_files == 1:
                current_output_path = output_path
            else:
                path_obj = Path(output_path)
                current_output_path = str(path_obj.parent / f"{path_obj.stem}_part{file_num + 1:02d}{path_obj.suffix}")
            
            # 创建当前文件的PDF
            self._create_single_box_label_file(
                data, params, current_output_path, style, 
                start_box, end_box, top_text, base_number
            )

    def _create_single_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
        style: str, start_box: int, end_box: int, top_text: str, base_number: str
    ):
        """创建单个盒标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"盒标-{style}-{start_box}到{end_box}")
        c.setSubject("Product Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 真正的三等分留白布局：每个留白区域高度相等
        blank_height = height / 5  # 每个留白区域高度：10mm
        
        # 布局位置计算（确保三个留白区域等高）
        # 留白1: height到(height - blank_height) = 40mm到50mm
        # 产品名称区域: (height - blank_height)到(height - 2*blank_height) = 30mm到40mm  
        # 留白2: (height - 2*blank_height)到(height - 3*blank_height) = 20mm到30mm
        # 序列号区域: (height - 3*blank_height)到(height - 4*blank_height) = 10mm到20mm
        # 留白3: (height - 4*blank_height)到0 = 0mm到10mm
        
        top_text_y = height - 1.5 * blank_height      # 产品名称居中在区域2
        serial_number_y = height - 3.5 * blank_height # 序列号居中在区域4

        # 生成指定范围的盒标
        for box_num in range(start_box, end_box + 1):
            if box_num > start_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 生成当前盒标的编号 - 截取数字前面的所有字符作为前缀
            import re
            # 查找第一个数字的位置，数字前面的所有字符都是前缀
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字的起始位置
                digit_start = match.start()
                # 截取数字前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]  # 比如 "GLA-", "A—B@C", "XYZ_"
                base_num = int(match.group(1))  # 第一个连续数字 01001
                
                # 生成新序列号：前缀 + 新数字（保持5位格式）
                current_number = f"{prefix_part}{base_num + box_num - 1:05d}"
            else:
                # 备用方案：如果完全没有数字
                current_number = f"DSK{box_num:05d}"
            
            # 根据外观选择不同的样式
            if style == "外观一":
                self._render_appearance_one(c, width, top_text, current_number, top_text_y, serial_number_y)
            else:
                # 外观二需要额外的票数信息
                ticket_count = params["张/盒"]
                self._render_appearance_two(c, width, top_text, ticket_count, current_number, top_text_y, serial_number_y)

        c.save()

    def _render_appearance_one(self, c, width, top_text, serial_number, top_text_y, serial_number_y):
        """渲染外观一：简洁标准样式"""
        # 使用多次绘制实现加粗效果
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        
        # 设置最适合的字体
        font_name = self._set_best_font(c, 22, bold=True)
        
        # 清理文本，移除不支持的字符
        clean_top_text = self._clean_text_for_font(top_text)
        
        # 处理上部文本的自动换行和字体大小调整
        top_text_lines = self._wrap_text_to_fit(c, clean_top_text, width - 4*mm, 22)
        
        # 根据行数调整字体大小和位置
        if len(top_text_lines) > 1:
            # 多行时使用较小字体
            self._set_best_font(c, 18, bold=True)
            line_height = 20  # 行间距
            start_y = top_text_y + (len(top_text_lines) - 1) * line_height / 2
            
            for i, line in enumerate(top_text_lines):
                current_y = start_y - i * line_height
                # 多次绘制增加粗细
                for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                    c.drawCentredString(width / 2 + offset[0], current_y + offset[1], line)
        else:
            # 单行时使用原字体大小
            for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                c.drawCentredString(width / 2 + offset[0], top_text_y + offset[1], top_text_lines[0])

        # 重置字体大小绘制序列号
        self._set_best_font(c, 22, bold=True)
        # 下部序列号 - 多次绘制增加粗细  
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], serial_number_y + offset[1], serial_number)

    def _render_appearance_two(self, c, width, game_title, ticket_count, serial_number, top_y, bottom_y):
        """渲染外观二：三行信息格式"""
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        self._set_best_font(c, 12, bold=True)
        
        # 获取页面高度
        page_height = self.page_size[1]
        
        # 左边距 - 统一的左边距
        left_margin = 4 * mm
        
        # 清理文本，移除不支持的字符
        clean_game_title = self._clean_text_for_font(game_title)
        clean_serial_number = self._clean_text_for_font(str(serial_number))
        
        # Game title: 距离顶部更大的距离，支持换行并居中显示
        game_title_y = page_height - 12 * mm  # 距离顶部12mm
        max_title_width = width - 2 * left_margin  # 留出左右边距
        
        # 将Game title文本分为标签部分和内容部分
        title_prefix = "Game title: "
        title_content = clean_game_title
        
        # 计算标签部分宽度
        used_font = self._set_best_font(c, 12, bold=True)  # 使用最佳字体
        prefix_width = c.stringWidth(title_prefix, used_font, 12)
        
        # 计算内容可用宽度
        content_max_width = max_title_width - prefix_width
        
        # 对内容部分进行换行处理
        title_lines = self._wrap_text_to_fit(c, title_content, content_max_width, 12)
        
        # 绘制Game title - 加粗效果通过多次绘制实现
        for i, line in enumerate(title_lines):
            current_y = game_title_y - i * 14  # 行间距14点
            if i == 0:
                # 第一行包含"Game title: "前缀，左对齐，多次绘制加粗
                full_line = f"{title_prefix}{line}"
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawString(left_margin + offset[0], current_y + offset[1], full_line)
            else:
                # 后续换行内容居中显示，多次绘制加粗
                line_width = c.stringWidth(line, used_font, 12)
                center_x = (width - line_width) / 2
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawString(center_x + offset[0], current_y + offset[1], line)
        
        # Ticket count: 左下区域，增加与Serial的间距，多次绘制加粗
        ticket_count_y = 15 * mm  # 距离底部15mm
        ticket_text = f"Ticket count: {ticket_count}"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawString(left_margin + offset[0], ticket_count_y + offset[1], ticket_text)
        
        # Serial: 距离底部，多次绘制加粗
        serial_y = 6 * mm  # 距离底部6mm
        serial_text = f"Serial: {clean_serial_number}"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawString(left_margin + offset[0], serial_y + offset[1], serial_text)

    def _create_fenhe_box_label(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None
    ):
        """创建分盒模板的盒标 - 特殊序列号逻辑"""
        # 计算总盒数
        total_pieces = int(float(data["总张数"]))  # 处理Excel的float值
        pieces_per_box = int(params["张/盒"])
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 获取盒标内容 - 只使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        top_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        print(f"✅ 分盒盒标使用关键字提取: 标签名称='{top_text}', 开始号='{base_number}'")
        
        # 从用户输入的第三个参数获取分组大小（从"小箱/大箱"参数获取）
        try:
            group_size = int(params["小箱/大箱"])  # 用户的第三个参数，控制副号满几进一
            if group_size <= 0:  # 避免除零错误
                group_size = 2
            print(f"✅ 分盒盒标使用用户输入分组大小: {group_size} (小箱/大箱)")
        except (ValueError, KeyError) as e:
            print(f"⚠️ 获取小箱/大箱参数失败: {e}")
            group_size = 2  # 默认分组大小
        
        # 直接创建单个PDF文件，包含所有盒标（移除分页限制）
        self._create_single_fenhe_box_label_file(
            data, params, output_path, style, 
            1, total_boxes, top_text, base_number, group_size
        )

    def _create_single_fenhe_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
        style: str, start_box: int, end_box: int, top_text: str, base_number: str, group_size: int
    ):
        """创建单个分盒模板盒标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"分盒盒标-{style}-{start_box}到{end_box}")
        c.setSubject("Fenhe Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 真正的三等分留白布局：每个留白区域高度相等
        blank_height = height / 5  # 每个留白区域高度：10mm
        
        # 布局位置计算（确保三个留白区域等高）
        top_text_y = height - 1.5 * blank_height      # 产品名称居中在区域2
        serial_number_y = height - 3.5 * blank_height # 序列号居中在区域4

        # 生成指定范围的盒标
        for box_num in range(start_box, end_box + 1):
            if box_num > start_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 生成分盒模板的序列号 - 特殊处理主号和副号
            import re
            # 分盒模板格式：前缀+主号+"-"+副号，如 ABC01001-02
            # 需要找到主号（第一个数字）前面的字符作为前缀
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                main_digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:main_digit_start]  # 比如 "ABC", "GLA-", "A—B@C"
                base_main_num = int(match.group(1))  # 主号，如 01001
                
                # 计算主号码和后缀
                # box_num从1开始，需要转换为0基数来计算
                box_index = box_num - 1
                main_number_offset = box_index // group_size  # 主号码偏移
                suffix_number = (box_index % group_size) + 1  # 后缀号码(1开始)
                
                new_main_number = base_main_num + main_number_offset
                # 分盒模板格式：前缀 + 主号码 + "-" + 副号
                current_number = f"{prefix_part}{new_main_number:05d}-{suffix_number:02d}"
            else:
                # 备用方案
                current_number = f"DSK{box_num:05d}-01"
            
            # 分盒模板只有一种固定外观，使用简洁标准样式
            self._render_appearance_one(c, width, top_text, current_number, top_text_y, serial_number_y)

        c.save()

    def _create_fenhe_small_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_small_boxes: int,
        remainder_info: Dict[str, Any],
        excel_file_path: str = None,
    ):
        """创建分盒模板的小箱标"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        print(f"✅ 分盒小箱标使用关键字提取: 标签名称='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 获取用户输入的分组大小（从"小箱/大箱"参数获取）
        try:
            group_size = int(params["小箱/大箱"])  # 用户的第三个参数，控制副号满几进一
            if group_size <= 0:
                group_size = 2
            print(f"✅ 分盒小箱标使用用户输入分组大小: {group_size} (小箱/大箱)")
        except (ValueError, KeyError) as e:
            print(f"⚠️ 获取小箱/大箱参数失败: {e}")
            group_size = 2  # 默认分组大小
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 直接创建单个PDF文件，包含所有小箱标
        self._create_single_fenhe_small_box_label_file(
            data, params, output_path, 1, total_small_boxes,
            theme_text, base_number, remark_text, pieces_per_small_box, 
            boxes_per_small_box, total_small_boxes, group_size
        )

    def _create_single_fenhe_small_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_small_box: int, end_small_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_small_box: int, boxes_per_small_box: int, 
        total_small_boxes: int, group_size: int
    ):
        """创建单个分盒小箱标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"分盒小箱标-{start_small_box}到{end_small_box}")
        c.setSubject("Fenhe Small Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的分盒小箱标
        for small_box_num in range(start_small_box, end_small_box + 1):
            if small_box_num > start_small_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算分盒模板的序列号范围
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]
                base_main_num = int(match.group(1))  # 主号
                
                # 分盒模板小箱标的特殊逻辑：
                # 每个小箱标包含boxes_per_small_box个盒标的序列号范围
                # 计算当前小箱标包含的盒标范围
                start_box_index = (small_box_num - 1) * boxes_per_small_box  # 起始盒标索引(0基数)
                end_box_index = start_box_index + boxes_per_small_box - 1    # 结束盒标索引(0基数)
                
                # 计算起始盒标的序列号
                start_main_offset = start_box_index // group_size
                start_suffix = (start_box_index % group_size) + 1
                start_main_number = base_main_num + start_main_offset
                start_serial = f"{prefix_part}{start_main_number:05d}-{start_suffix:02d}"
                
                # 计算结束盒标的序列号
                end_main_offset = end_box_index // group_size
                end_suffix = (end_box_index % group_size) + 1
                end_main_number = base_main_num + end_main_offset
                end_serial = f"{prefix_part}{end_main_number:05d}-{end_suffix:02d}"
                
                # 分盒小箱标显示序列号范围
                serial_range = f"{start_serial}-{end_serial}"
            else:
                serial_range = f"DSK{small_box_num:05d}-DSK{small_box_num:05d}"

            # 计算分盒小箱标的Carton No（主箱号-副箱号格式）
            main_box_num = ((small_box_num - 1) // group_size) + 1  # 主箱号
            sub_box_num = ((small_box_num - 1) % group_size) + 1    # 副箱号
            carton_no = f"{main_box_num}-{sub_box_num}"

            # 绘制分盒小箱标表格
            self._draw_fenhe_small_box_table(c, width, height, theme_text, pieces_per_small_box, 
                                           serial_range, carton_no, remark_text)

        c.save()

    def _draw_fenhe_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                                  serial_range, carton_no, remark_text):
        """绘制分盒小箱标表格"""
        # 表格尺寸和位置
        table_width = width - 4 * mm
        table_height = height - 4 * mm
        table_x = 2 * mm
        table_y = 2 * mm
        
        # 高度分配：Quantity行占2/6，其他4行各占1/6
        base_row_height = table_height / 6
        quantity_row_height = base_row_height * 2  # Quantity行双倍高度
        
        # 列宽 (标签列:数据列 = 1:2)
        label_col_width = table_width / 3
        data_col_width = table_width * 2 / 3
        
        # 绘制表格边框
        c.setStrokeColor(CMYKColor(0, 0, 0, 1))
        c.setLineWidth(1)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 计算各行的Y坐标
        row_positions = []
        current_y = table_y
        # 从底部开始：Remark, Carton No, Quantity(双倍), Theme, Item
        for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height, base_row_height]:
            row_positions.append(current_y)
            current_y += height_val
        
        # 绘制行线
        for i in range(1, 5):
            y = row_positions[i]
            c.line(table_x, y, table_x + table_width, y)
        
        # 绘制列线
        col_x = table_x + label_col_width
        c.line(col_x, table_y, col_x, table_y + table_height)
        
        # 绘制Quantity行的分隔线（上层和下层之间）
        quantity_split_y = row_positions[2] + quantity_row_height / 2
        c.line(col_x, quantity_split_y, table_x + table_width, quantity_split_y)
        
        # 表格内容
        self._set_best_font(c, 10, bold=True)
        
        # 计算居中位置
        label_center_x = table_x + label_col_width / 2  # 标签列居中
        data_center_x = col_x + data_col_width / 2      # 数据列居中
        
        # 行1: Item (第5行，从上往下) - 多次绘制加粗
        item_y = row_positions[4] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], item_y + offset[1], "Item:")
            c.drawCentredString(data_center_x + offset[0], item_y + offset[1], "Paper Cards")
        
        # 行2: Theme (第4行) - 多次绘制加粗
        theme_y = row_positions[3] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], theme_y + offset[1], "Theme:")
        
        # 应用文本清理和换行处理
        clean_theme_text = self._clean_text_for_font(theme_text)
        max_theme_width = data_col_width - 4*mm  # 留出边距
        theme_lines = self._wrap_text_to_fit(c, clean_theme_text, max_theme_width, 10)
        
        # 绘制主题文本（支持多行） - 多次绘制加粗
        if len(theme_lines) > 1:
            # 多行：调整字体大小并垂直居中
            self._set_best_font(c, 8, bold=True)
            line_height = 10
            # 计算整个文本块的总高度
            total_text_height = (len(theme_lines) - 1) * line_height
            # 让文本块在单元格中垂直居中
            start_y = theme_y + total_text_height / 2
            for i, line in enumerate(theme_lines):
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawCentredString(data_center_x + offset[0], start_y - i * line_height + offset[1], line)
            self._set_best_font(c, 10, bold=True)  # 恢复字体大小
        else:
            for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                c.drawCentredString(data_center_x + offset[0], theme_y + offset[1], theme_lines[0])
        
        # 行3: Quantity (第3行，双倍高度) - 多次绘制加粗
        quantity_label_y = row_positions[2] + quantity_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], quantity_label_y + offset[1], "Quantity:")
        # 上层：票数（在分隔线上方居中）
        upper_y = row_positions[2] + quantity_row_height * 3/4
        pcs_text = f"{pieces_per_small_box}PCS"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], upper_y + offset[1], pcs_text)
        # 下层：序列号范围（在分隔线下方居中）
        lower_y = row_positions[2] + quantity_row_height/4
        clean_serial_range = self._clean_text_for_font(serial_range)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], lower_y + offset[1], clean_serial_range)
        
        # 行4: Carton No (第2行) - 多次绘制加粗，显示主箱号-副箱号格式
        carton_y = row_positions[1] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], carton_y + offset[1], "Carton No:")
            c.drawCentredString(data_center_x + offset[0], carton_y + offset[1], carton_no)
        
        # 行5: Remark (第1行) - 多次绘制加粗
        remark_y = row_positions[0] + base_row_height/2
        clean_remark_text = self._clean_text_for_font(remark_text)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], remark_y + offset[1], "Remark:")
            c.drawCentredString(data_center_x + offset[0], remark_y + offset[1], clean_remark_text)

    def _create_small_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_small_boxes: int,
        remainder_info: Dict[str, Any],
        excel_file_path: str = None,
    ):
        """创建小箱标 - 支持分页限制的多页PDF"""
        # 获取Excel数据 - 使用关键字提取或备用方案
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用关键字提取
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        print(f"✅ 小箱标使用关键字提取: 标签名称='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 计算需要的文件数量
        total_files = math.ceil(total_small_boxes / self.max_pages_per_file)
        
        # 分页生成多个PDF文件
        for file_num in range(total_files):
            # 计算当前文件的页数范围
            start_small_box = file_num * self.max_pages_per_file + 1
            end_small_box = min((file_num + 1) * self.max_pages_per_file, total_small_boxes)
            
            # 生成文件名
            if total_files == 1:
                current_output_path = output_path
            else:
                path_obj = Path(output_path)
                current_output_path = str(path_obj.parent / f"{path_obj.stem}_part{file_num + 1:02d}{path_obj.suffix}")
            
            # 创建当前文件的PDF
            self._create_single_small_box_label_file(
                data, params, current_output_path, start_small_box, end_small_box,
                theme_text, base_number, remark_text, pieces_per_small_box, 
                boxes_per_small_box, total_small_boxes
            )

    def _create_single_small_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_small_box: int, end_small_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_small_box: int, boxes_per_small_box: int, total_small_boxes: int
    ):
        """创建单个小箱标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"小箱标-{start_small_box}到{end_small_box}")
        c.setSubject("Small Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的小箱标
        for small_box_num in range(start_small_box, end_small_box + 1):
            if small_box_num > start_small_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算当前小箱的序列号范围
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字的起始位置
                digit_start = match.start()
                # 截取数字前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]  # 比如 "GLA-", "A—B@C"
                base_num = int(match.group(1))  # 第一个连续数字
                
                # 计算序列号范围
                start_serial = base_num + (small_box_num - 1) * boxes_per_small_box
                end_serial = start_serial + boxes_per_small_box - 1
                
                # 生成序列号格式
                start_format = f"{prefix_part}{start_serial:05d}"
                end_format = f"{prefix_part}{end_serial:05d}"
                
                if boxes_per_small_box == 1:
                    # 一盒一小箱：序列号相同
                    serial_range = f"{start_format}-{start_format}"
                else:
                    # 多盒一小箱：序列号范围
                    serial_range = f"{start_format}-{end_format}"
            else:
                serial_range = f"DSK{small_box_num:05d}-DSK{small_box_num:05d}"

            # 绘制表格
            self._draw_small_box_table(c, width, height, theme_text, pieces_per_small_box, 
                                     serial_range, small_box_num, total_small_boxes, remark_text)

        c.save()

    def _draw_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                            serial_range, small_box_num, total_small_boxes, remark_text):
        """绘制小箱标表格"""
        # 表格尺寸和位置
        table_width = width - 4 * mm
        table_height = height - 4 * mm
        table_x = 2 * mm
        table_y = 2 * mm
        
        # 高度分配：Quantity行占2/6，其他4行各占1/6
        base_row_height = table_height / 6
        quantity_row_height = base_row_height * 2  # Quantity行双倍高度
        
        # 列宽 (标签列:数据列 = 1:2)
        label_col_width = table_width / 3
        data_col_width = table_width * 2 / 3
        
        # 绘制表格边框
        c.setStrokeColor(CMYKColor(0, 0, 0, 1))
        c.setLineWidth(1)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 计算各行的Y坐标
        row_positions = []
        current_y = table_y
        # 从底部开始：Remark, Carton No, Quantity(双倍), Theme, Item
        for height in [base_row_height, base_row_height, quantity_row_height, base_row_height, base_row_height]:
            row_positions.append(current_y)
            current_y += height
        
        # 绘制行线
        for i in range(1, 5):
            y = row_positions[i]
            c.line(table_x, y, table_x + table_width, y)
        
        # 绘制列线
        col_x = table_x + label_col_width
        c.line(col_x, table_y, col_x, table_y + table_height)
        
        # 绘制Quantity行的分隔线（上层和下层之间）
        quantity_split_y = row_positions[2] + quantity_row_height / 2
        c.line(col_x, quantity_split_y, table_x + table_width, quantity_split_y)
        
        # 表格内容
        self._set_best_font(c, 10, bold=True)
        
        # 计算居中位置
        label_center_x = table_x + label_col_width / 2  # 标签列居中
        data_center_x = col_x + data_col_width / 2      # 数据列居中
        
        # 行1: Item (第5行，从上往下) - 多次绘制加粗
        item_y = row_positions[4] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], item_y + offset[1], "Item:")
            c.drawCentredString(data_center_x + offset[0], item_y + offset[1], "Paper Cards")
        
        # 行2: Theme (第4行) - 多次绘制加粗
        theme_y = row_positions[3] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], theme_y + offset[1], "Theme:")
        
        # 应用文本清理和换行处理
        clean_theme_text = self._clean_text_for_font(theme_text)
        max_theme_width = data_col_width - 4*mm  # 留出边距
        theme_lines = self._wrap_text_to_fit(c, clean_theme_text, max_theme_width, 10)
        
        # 绘制主题文本（支持多行） - 多次绘制加粗
        if len(theme_lines) > 1:
            # 多行：调整字体大小并垂直居中
            self._set_best_font(c, 8, bold=True)
            line_height = 10
            # 计算整个文本块的总高度
            total_text_height = (len(theme_lines) - 1) * line_height
            # 让文本块在单元格中垂直居中
            start_y = theme_y + total_text_height / 2
            for i, line in enumerate(theme_lines):
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawCentredString(data_center_x + offset[0], start_y - i * line_height + offset[1], line)
            self._set_best_font(c, 10, bold=True)  # 恢复字体大小
        else:
            for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                c.drawCentredString(data_center_x + offset[0], theme_y + offset[1], theme_lines[0])
        
        # 行3: Quantity (第3行，双倍高度) - 多次绘制加粗
        quantity_label_y = row_positions[2] + quantity_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], quantity_label_y + offset[1], "Quantity:")
        # 上层：票数（在分隔线上方居中）
        upper_y = row_positions[2] + quantity_row_height * 3/4
        pcs_text = f"{pieces_per_small_box}PCS"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], upper_y + offset[1], pcs_text)
        # 下层：序列号（在分隔线下方居中）
        lower_y = row_positions[2] + quantity_row_height/4
        clean_serial_range = self._clean_text_for_font(serial_range)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], lower_y + offset[1], clean_serial_range)
        
        # 行4: Carton No (第2行) - 多次绘制加粗
        carton_y = row_positions[1] + base_row_height/2
        carton_text = f"{small_box_num}/{total_small_boxes}"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], carton_y + offset[1], "Carton No:")
            c.drawCentredString(data_center_x + offset[0], carton_y + offset[1], carton_text)
        
        # 行5: Remark (第1行) - 多次绘制加粗
        remark_y = row_positions[0] + base_row_height/2
        clean_remark_text = self._clean_text_for_font(remark_text)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], remark_y + offset[1], "Remark:")
            c.drawCentredString(data_center_x + offset[0], remark_y + offset[1], clean_remark_text)

    def _create_large_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_large_boxes: int,
        total_small_boxes: int,
        remainder_info: Dict[str, Any],
        excel_file_path: str = None,
    ):
        """创建大箱标 - 支持分页限制的多页PDF"""
        # 获取Excel数据 - 使用关键字提取或备用方案
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用关键字提取
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        print(f"✅ 大箱标使用关键字提取: 标签名称='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])
        pieces_per_large_box = pieces_per_box * boxes_per_small_box * small_boxes_per_large_box
        
        # 计算需要的文件数量
        total_files = math.ceil(total_large_boxes / self.max_pages_per_file)
        
        # 分页生成多个PDF文件
        for file_num in range(total_files):
            # 计算当前文件的页数范围
            start_large_box = file_num * self.max_pages_per_file + 1
            end_large_box = min((file_num + 1) * self.max_pages_per_file, total_large_boxes)
            
            # 生成文件名
            if total_files == 1:
                current_output_path = output_path
            else:
                path_obj = Path(output_path)
                current_output_path = str(path_obj.parent / f"{path_obj.stem}_part{file_num + 1:02d}{path_obj.suffix}")
            
            # 创建当前文件的PDF
            self._create_single_large_box_label_file(
                data, params, current_output_path, start_large_box, end_large_box,
                theme_text, base_number, remark_text, pieces_per_large_box,
                boxes_per_small_box, small_boxes_per_large_box, total_large_boxes
            )

    def _create_single_large_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_large_box: int, boxes_per_small_box: int,
        small_boxes_per_large_box: int, total_large_boxes: int
    ):
        """创建单个大箱标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"大箱标-{start_large_box}到{end_large_box}")
        c.setSubject("Large Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的大箱标
        for large_box_num in range(start_large_box, end_large_box + 1):
            if large_box_num > start_large_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算当前大箱包含的盒序列号范围
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字的起始位置
                digit_start = match.start()
                # 截取数字前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]  # 比如 "GLA-", "A—B@C"
                base_num = int(match.group(1))  # 第一个连续数字
                
                # 计算大箱内盒序列号范围
                # 每大箱包含的盒数 = 盒/小箱 × 小箱/大箱
                boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
                start_box_serial = base_num + (large_box_num - 1) * boxes_per_large_box
                end_box_serial = start_box_serial + boxes_per_large_box - 1
                
                # 生成大箱序列号范围格式
                box_serial_range = f"{prefix_part}{start_box_serial:05d}-{prefix_part}{end_box_serial:05d}"
            else:
                box_serial_range = f"DSK{large_box_num:05d}-DSK{large_box_num:05d}"

            # 绘制表格
            self._draw_large_box_table(c, width, height, theme_text, pieces_per_large_box,
                                     box_serial_range, large_box_num, total_large_boxes, remark_text)

        c.save()

    def _draw_large_box_table(self, c, width, height, theme_text, pieces_per_large_box,
                            box_serial_range, large_box_num, total_large_boxes, remark_text):
        """绘制大箱标表格"""
        # 复用小箱标的表格绘制逻辑，只是数据不同
        # 表格尺寸和位置
        table_width = width - 4 * mm
        table_height = height - 4 * mm
        table_x = 2 * mm
        table_y = 2 * mm
        
        # 高度分配：Quantity行占2/6，其他4行各占1/6
        base_row_height = table_height / 6
        quantity_row_height = base_row_height * 2
        
        # 列宽 (标签列:数据列 = 1:2)
        label_col_width = table_width / 3
        
        # 绘制表格边框
        c.setStrokeColor(CMYKColor(0, 0, 0, 1))
        c.setLineWidth(1)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 计算各行的Y坐标
        row_positions = []
        current_y = table_y
        # 从底部开始：Remark, Carton No, Quantity(双倍), Theme, Item
        for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height, base_row_height]:
            row_positions.append(current_y)
            current_y += height_val
        
        # 绘制行线
        for i in range(1, 5):
            y = row_positions[i]
            c.line(table_x, y, table_x + table_width, y)
        
        # 绘制列线
        col_x = table_x + label_col_width
        c.line(col_x, table_y, col_x, table_y + table_height)
        
        # 绘制Quantity行的分隔线
        quantity_split_y = row_positions[2] + quantity_row_height / 2
        c.line(col_x, quantity_split_y, table_x + table_width, quantity_split_y)
        
        # 表格内容
        self._set_best_font(c, 10, bold=True)
        
        # 计算居中位置
        label_center_x = table_x + label_col_width / 2
        data_center_x = col_x + (table_width - label_col_width) / 2
        
        # 行1: Item - 多次绘制加粗
        item_y = row_positions[4] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], item_y + offset[1], "Item:")
            c.drawCentredString(data_center_x + offset[0], item_y + offset[1], "Paper Cards")
        
        # 行2: Theme - 多次绘制加粗
        theme_y = row_positions[3] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], theme_y + offset[1], "Theme:")
        
        # 应用文本清理和换行处理
        clean_theme_text = self._clean_text_for_font(theme_text)
        data_col_width = table_width - label_col_width
        max_theme_width = data_col_width - 4*mm  # 留出边距
        theme_lines = self._wrap_text_to_fit(c, clean_theme_text, max_theme_width, 10)
        
        # 绘制主题文本（支持多行） - 多次绘制加粗
        if len(theme_lines) > 1:
            # 多行：调整字体大小并垂直居中
            self._set_best_font(c, 8, bold=True)
            line_height = 10
            # 计算整个文本块的总高度
            total_text_height = (len(theme_lines) - 1) * line_height
            # 让文本块在单元格中垂直居中
            start_y = theme_y + total_text_height / 2
            for i, line in enumerate(theme_lines):
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawCentredString(data_center_x + offset[0], start_y - i * line_height + offset[1], line)
            self._set_best_font(c, 10, bold=True)  # 恢复字体大小
        else:
            for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                c.drawCentredString(data_center_x + offset[0], theme_y + offset[1], theme_lines[0])
        
        # 行3: Quantity (双倍高度) - 多次绘制加粗
        quantity_label_y = row_positions[2] + quantity_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], quantity_label_y + offset[1], "Quantity:")
        # 上层：每大箱票数
        upper_y = row_positions[2] + quantity_row_height * 3/4
        large_pcs_text = f"{pieces_per_large_box}PCS"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], upper_y + offset[1], large_pcs_text)
        # 下层：大箱内盒序列号范围
        lower_y = row_positions[2] + quantity_row_height/4
        clean_box_serial_range = self._clean_text_for_font(box_serial_range)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], lower_y + offset[1], clean_box_serial_range)
        
        # 行4: Carton No - 多次绘制加粗
        carton_y = row_positions[1] + base_row_height/2
        large_carton_text = f"{large_box_num}/{total_large_boxes}"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], carton_y + offset[1], "Carton No:")
            c.drawCentredString(data_center_x + offset[0], carton_y + offset[1], large_carton_text)
        
        # 行5: Remark - 多次绘制加粗
        remark_y = row_positions[0] + base_row_height/2
        clean_remark_text = self._clean_text_for_font(remark_text)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], remark_y + offset[1], "Remark:")
            c.drawCentredString(data_center_x + offset[0], remark_y + offset[1], clean_remark_text)

    def create_label_pdf(self, data: Dict[str, Any], output_path: str):
        """
        创建标签PDF (保持兼容性)

        Args:
            data: 标签数据
            output_path: 输出路径
        """
        self.generate_from_template(None, data, output_path)

    def _render_default_layout(
        self, canvas_obj, data: Dict[str, Any], width: float, height: float
    ):
        """
        渲染默认布局

        Args:
            canvas_obj: 画布对象
            data: 数据
            width: 页面宽度
            height: 页面高度
        """
        # 计算起始位置
        x_start = self.margin
        y_start = height - self.margin - self.font_size

        line_height = self.font_size + 5
        y_pos = y_start

        # 渲染标题
        title = "数据标签"
        if self._has_chinese(title):
            canvas_obj.setFont(self.chinese_font_name, self.font_size + 4)
        else:
            canvas_obj.setFont(self.font_name, self.font_size + 4)
        canvas_obj.drawString(x_start, y_pos, title)
        y_pos -= line_height * 2

        # 渲染数据字段
        canvas_obj.setFont(self.font_name, self.font_size)
        for key, value in data.items():
            text = f"{key}: {value}"
            # 清理文本，移除不支持的字符
            clean_text = self._clean_text_for_font(text)
            # 如果包含中文，使用中文字体
            if self._has_chinese(clean_text):
                canvas_obj.setFont(self.chinese_font_name, self.font_size)
            else:
                canvas_obj.setFont(self.font_name, self.font_size)
            canvas_obj.drawString(x_start, y_pos, clean_text)
            y_pos -= line_height

    def _register_chinese_font(self):
        """
        注册中文字体 - 优先使用微软雅黑
        """
        try:
            # 根据操作系统选择字体路径，优先微软雅黑
            system = platform.system()
            font_paths = []

            if system == "Darwin":  # macOS
                font_paths = [
                    "/System/Library/Fonts/Microsoft/msyh.ttf",  # 微软雅黑
                    "/Library/Fonts/Microsoft YaHei.ttf",       # 用户安装的微软雅黑
                    "/System/Library/Fonts/PingFang.ttc",       # 苹方字体
                    "/System/Library/Fonts/STHeiti Light.ttc",  # 黑体Light
                    "/System/Library/Fonts/STHeiti Medium.ttc", # 黑体Medium
                    "/System/Library/Fonts/Supplemental/Songti.ttc", # 宋体
                ]
            elif system == "Windows":
                font_paths = [
                    "C:/Windows/Fonts/msyh.ttc",     # 微软雅黑
                    "C:/Windows/Fonts/msyhbd.ttc",   # 微软雅黑粗体
                    "C:/Windows/Fonts/simhei.ttf",   # 黑体
                    "C:/Windows/Fonts/simsun.ttc",   # 宋体
                ]
            elif system == "Linux":
                font_paths = [
                    "/usr/share/fonts/truetype/msyh/msyh.ttf",    # 微软雅黑
                    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", # 文泉驿微米黑
                    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", # Noto中文
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", # 备选
                ]

            # 尝试注册第一个可用的字体
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(
                            TTFont(self.chinese_font_name, font_path)
                        )
                        self.font_name = self.chinese_font_name
                        print(f"✅ 成功注册中文字体: {font_path}")
                        return
                    except Exception as e:
                        print(f"❌ 注册字体失败 {font_path}: {e}")
                        continue

            # 如果没有找到中文字体，回退到默认字体但发出警告
            print("⚠️ 警告: 未找到微软雅黑或其他中文字体，回退到Helvetica")
            self.font_name = "Helvetica"
            self.chinese_font_name = "Helvetica"

        except Exception as e:
            print(f"❌ 注册中文字体过程失败: {e}")
            self.font_name = "Helvetica"
            self.chinese_font_name = "Helvetica"

    def _has_chinese(self, text: str) -> bool:
        """
        检查文本是否包含中文字符
        """
        for char in str(text):
            if "\u4e00" <= char <= "\u9fff":
                return True
        return False

    def _create_fenhe_large_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_large_boxes: int,
        excel_file_path: str = None,
    ):
        """创建分盒模板的大箱标 - 完全参考小箱标模式"""
        # 获取Excel数据 - 使用关键字提取，与小箱标相同
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        print(f"✅ 分盒大箱标使用关键字提取: 标签名称='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 获取用户输入的分组大小（从"小箱/大箱"参数获取）
        try:
            group_size = int(params["小箱/大箱"])  # 用户的第三个参数，控制副号满几进一
            print(f"✅ 分盒大箱标使用用户输入分组大小: {group_size} (小箱/大箱)")
        except (ValueError, KeyError) as e:
            print(f"⚠️ 获取小箱/大箱参数失败: {e}")
            group_size = 2  # 默认分组大小
        
        # 计算参数 - 大箱标专用
        pieces_per_box = int(params["张/盒"])  # 第一个参数：张/盒
        boxes_per_small_box = int(params["盒/小箱"])  # 第二个参数：盒/小箱
        small_boxes_per_large_box = int(params["小箱/大箱"])  # 第三个参数：小箱/大箱
        
        # 直接创建单个PDF文件，包含所有大箱标
        self._create_single_fenhe_large_box_label_file(
            data, params, output_path, 1, total_large_boxes,
            theme_text, base_number, remark_text, pieces_per_box, 
            small_boxes_per_large_box, total_large_boxes, group_size
        )

    def _create_single_fenhe_large_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_box: int, small_boxes_per_large_box: int, 
        total_large_boxes: int, group_size: int
    ):
        """创建单个分盒大箱标PDF文件 - 完全参考小箱标"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"分盒大箱标-{start_large_box}到{end_large_box}")
        c.setSubject("Fenhe Large Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的大箱标
        for large_box_num in range(start_large_box, end_large_box + 1):
            if large_box_num > start_large_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算当前大箱的序列号范围
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                main_digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:main_digit_start]  # 比如 "LGM"
                base_main_num = int(match.group(1))  # 主号，如 01001
                
                # 计算当前大箱的主号 (每个大箱主号递增)
                current_main_number = base_main_num + (large_box_num - 1)
                
                # 生成序列号范围（从01到group_size）
                start_serial = f"{prefix_part}{current_main_number:05d}-01"
                end_serial = f"{prefix_part}{current_main_number:05d}-{group_size:02d}"
                serial_range = f"{start_serial}-{end_serial}"
            else:
                # 备用方案
                serial_range = f"DSK{large_box_num:05d}-01-DSK{large_box_num:05d}-{group_size:02d}"
            
            # 绘制大箱标表格 - 完全使用小箱标的表格结构
            self._draw_fenhe_large_box_table(c, width, height, theme_text, pieces_per_box, 
                                           small_boxes_per_large_box, serial_range, 
                                           str(large_box_num), remark_text)

        c.save()

    def _draw_fenhe_large_box_table(self, c, width, height, theme_text, pieces_per_box, 
                                  small_boxes_per_large_box, serial_range, carton_no, remark_text):
        """绘制分盒大箱标表格 - 完全复制小箱标实现"""
        # 表格尺寸和位置
        table_width = width - 4 * mm
        table_height = height - 4 * mm
        table_x = 2 * mm
        table_y = 2 * mm
        
        # 高度分配：Quantity行占2/6，其他4行各占1/6
        base_row_height = table_height / 6
        quantity_row_height = base_row_height * 2  # Quantity行双倍高度
        
        # 列宽 (标签列:数据列 = 1:2)
        label_col_width = table_width / 3
        data_col_width = table_width * 2 / 3
        
        # 绘制表格边框
        c.setStrokeColor(CMYKColor(0, 0, 0, 1))
        c.setLineWidth(1)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 计算各行的Y坐标
        row_positions = []
        current_y = table_y
        # 从底部开始：Remark, Carton No, Quantity(双倍), Theme, Item
        for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height, base_row_height]:
            row_positions.append(current_y)
            current_y += height_val
        
        # 绘制行线
        for i in range(1, 5):
            y = row_positions[i]
            c.line(table_x, y, table_x + table_width, y)
        
        # 绘制列线
        col_x = table_x + label_col_width
        c.line(col_x, table_y, col_x, table_y + table_height)
        
        # 绘制Quantity行的分隔线（上层和下层之间）
        quantity_split_y = row_positions[2] + quantity_row_height / 2
        c.line(col_x, quantity_split_y, table_x + table_width, quantity_split_y)
        
        # 表格内容
        self._set_best_font(c, 10, bold=True)
        
        # 计算居中位置
        label_center_x = table_x + label_col_width / 2  # 标签列居中
        data_center_x = col_x + data_col_width / 2      # 数据列居中
        
        # 行1: Item (第5行，从上往下) - 多次绘制加粗
        item_y = row_positions[4] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], item_y + offset[1], "Item:")
            c.drawCentredString(data_center_x + offset[0], item_y + offset[1], "Paper Cards")
        
        # 行2: Theme (第4行) - 多次绘制加粗
        theme_y = row_positions[3] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], theme_y + offset[1], "Theme:")
        
        # 应用文本清理和换行处理
        clean_theme_text = self._clean_text_for_font(theme_text)
        max_theme_width = data_col_width - 4*mm  # 留出边距
        theme_lines = self._wrap_text_to_fit(c, clean_theme_text, max_theme_width, 10)
        
        # 绘制主题文本（支持多行） - 多次绘制加粗
        if len(theme_lines) > 1:
            # 多行：调整字体大小并垂直居中
            self._set_best_font(c, 8, bold=True)
            line_height = 10
            # 计算整个文本块的总高度
            total_text_height = (len(theme_lines) - 1) * line_height
            # 让文本块在单元格中垂直居中
            start_y = theme_y + total_text_height / 2
            for i, line in enumerate(theme_lines):
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawCentredString(data_center_x + offset[0], start_y - i * line_height + offset[1], line)
            self._set_best_font(c, 10, bold=True)  # 恢复字体大小
        else:
            for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                c.drawCentredString(data_center_x + offset[0], theme_y + offset[1], theme_lines[0])
        
        # 行3: Quantity (第3行，双倍高度) - 多次绘制加粗
        quantity_label_y = row_positions[2] + quantity_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], quantity_label_y + offset[1], "Quantity:")
        # 上层：计算并显示 (张/盒 * 小箱/大箱)PCS
        upper_y = row_positions[2] + quantity_row_height * 3/4
        pcs_count = pieces_per_box * small_boxes_per_large_box  # 张/盒 * 小箱/大箱
        pcs_text = f"{pcs_count}PCS"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], upper_y + offset[1], pcs_text)
        # 下层：序列号范围（在分隔线下方居中）
        lower_y = row_positions[2] + quantity_row_height/4
        clean_serial_range = self._clean_text_for_font(serial_range)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], lower_y + offset[1], clean_serial_range)
        
        # 行4: Carton No (第2行) - 多次绘制加粗，显示大箱编号
        carton_y = row_positions[1] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], carton_y + offset[1], "Carton No:")
            c.drawCentredString(data_center_x + offset[0], carton_y + offset[1], carton_no)
        
        # 行5: Remark (第1行) - 多次绘制加粗
        remark_y = row_positions[0] + base_row_height/2
        clean_remark_text = self._clean_text_for_font(remark_text)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], remark_y + offset[1], "Remark:")
            c.drawCentredString(data_center_x + offset[0], remark_y + offset[1], clean_remark_text)

    # ===================== 套盒模板方法 =====================

    def create_taohebox_multi_level_pdfs(
        self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None
    ) -> Dict[str, str]:
        """创建套盒模板的多级标签PDF"""
        # 创建输出目录
        clean_theme = data['主题'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        generated_files = {}

        # 生成套盒模板的盒标 - 第二个参数用于结束号逻辑（套盒无外观选择）
        box_label_path = full_output_dir / f"{data['客户编码']}+{clean_theme}+套盒盒标.pdf"

        self._create_taohebox_box_label(data, params, str(box_label_path), excel_file_path)
        generated_files["盒标"] = str(box_label_path)

        # 计算各级数量用于小箱标生成
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])

        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)

        # 生成套盒模板小箱标
        small_box_path = full_output_dir / f"{data['客户编码']}+{clean_theme}+套盒小箱标.pdf"
        self._create_taohebox_small_box_label(
            data, params, str(small_box_path), total_small_boxes, excel_file_path
        )
        generated_files["小箱标"] = str(small_box_path)

        # 生成套盒模板大箱标
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)
        large_box_path = full_output_dir / f"{data['客户编码']}+{clean_theme}+套盒大箱标.pdf"
        self._create_taohebox_large_box_label(
            data, params, str(large_box_path), total_large_boxes, excel_file_path
        )
        generated_files["大箱标"] = str(large_box_path)

        return generated_files

    def _create_taohebox_box_label(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, excel_file_path: str = None
    ):
        """创建套盒模板的盒标 - 基于Excel文件的开始号和结束号"""
        # 分析Excel文件获取套盒特有的数据
        excel_path = excel_file_path
        print(f"🔍 正在分析套盒模板Excel文件: {excel_path}")
        
        try:
            import pandas as pd
            df = pd.read_excel(excel_path, header=None)
            print(f"✅ Excel文件已加载: {df.shape[0]}行 x {df.shape[1]}列")
            
            # 根据分析的结果提取数据
            # 标签名称：第10行第9列 (索引9,8)
            theme_text = df.iloc[9, 8] if pd.notna(df.iloc[9, 8]) else 'Unknown Title'
            
            # 开始号：第10行第2列 (索引9,1) 
            base_number = df.iloc[9, 1] if pd.notna(df.iloc[9, 1]) else 'DEFAULT01001'
            
            # 结束号：第10行第3列 (索引9,2)
            end_number = df.iloc[9, 2] if pd.notna(df.iloc[9, 2]) else base_number
            
            # 主题：第4行第2列 (索引3,1)
            full_theme = df.iloc[3, 1] if pd.notna(df.iloc[3, 1]) else 'Unknown Theme'
            
            print(f"✅ 套盒模板数据提取:")
            print(f"   标签名称: '{theme_text}'")
            print(f"   开始号: '{base_number}'")
            print(f"   结束号: '{end_number}'")
            print(f"   完整主题: '{full_theme}'")
            
        except Exception as e:
            print(f"❌ 读取Excel文件失败: {e}")
            # 回退到关键字提取
            excel_data = self._extract_excel_data_by_keywords(excel_path)
            theme_text = excel_data.get('标签名称') or 'Unknown Title'
            base_number = excel_data.get('开始号') or 'DEFAULT01001'
            end_number = excel_data.get('结束号') or base_number
        
        # 套盒模板参数分析
        pieces_per_box = int(params["张/盒"])
        boxes_per_ending_unit = int(params["盒/小箱"])  # 在套盒模板中，这个参数用于结束号的范围计算
        group_size = int(params["小箱/大箱"])
        
        print(f"✅ 套盒模板参数:")
        print(f"   张/盒: {pieces_per_box}")
        print(f"   盒/小箱(结束号范围): {boxes_per_ending_unit}")
        print(f"   小箱/大箱(分组大小): {group_size}")
        
        # 解析开始号和结束号的格式
        import re
        start_match = re.search(r'(.+?)(\d+)-(\d+)', base_number)
        end_match = re.search(r'(.+?)(\d+)-(\d+)', end_number)
        
        if start_match and end_match:
            start_prefix = start_match.group(1)
            start_main = int(start_match.group(2))
            start_suffix = int(start_match.group(3))
            
            end_prefix = end_match.group(1)
            end_main = int(end_match.group(2))
            end_suffix = int(end_match.group(3))
            
            print(f"✅ 解析序列号格式:")
            print(f"   开始: {start_prefix}{start_main:05d}-{start_suffix:02d}")
            print(f"   结束: {end_prefix}{end_main:05d}-{end_suffix:02d}")
            
        else:
            print("⚠️ 无法解析序列号格式，使用默认逻辑")
            start_prefix = "JAW"
            start_main = 1001
            start_suffix = 1
            end_suffix = boxes_per_ending_unit
        
        # 计算需要生成的盒标数量
        total_pieces = int(float(data["总张数"]))
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        c.setTitle("套盒模板盒标")
        
        width, height = self.page_size
        blank_height = height / 5
        top_text_y = height - 1.5 * blank_height
        serial_number_y = height - 3.5 * blank_height
        
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 生成套盒盒标 - 基于开始号到结束号的范围
        print(f"📝 开始生成套盒盒标，预计生成 {total_boxes} 个标签")
        
        for box_num in range(1, total_boxes + 1):
            if box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 套盒模板序列号生成逻辑 - 基于开始号和结束号范围
            box_index = box_num - 1
            
            # 计算当前盒的序列号在范围内的位置
            main_offset = box_index // boxes_per_ending_unit
            suffix_in_range = (box_index % boxes_per_ending_unit) + start_suffix
            
            current_main = start_main + main_offset
            current_number = f"{start_prefix}{current_main:05d}-{suffix_in_range:02d}"
            
            print(f"📝 生成套盒盒标 #{box_num}: {current_number}")
            
            # 套盒模板使用固定外观（类似分盒模板外观一）
            self._render_taohebox_appearance_one(c, width, theme_text, current_number, top_text_y, serial_number_y)

        c.save()
        print(f"✅ 套盒模板盒标PDF已生成: {output_path}")

    def _render_taohebox_appearance_one(self, c, width, top_text, current_number, top_text_y, serial_number_y):
        """套盒模板盒标外观一渲染"""
        clean_top_text = self._clean_text_for_font(top_text)
        self._set_best_font(c, 14, bold=True)
        
        # 绘制Game title和序列号 - 加粗效果
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], top_text_y + offset[1], clean_top_text)
            c.drawCentredString(width / 2 + offset[0], serial_number_y + offset[1], current_number)

    def _render_taohebox_appearance_two(self, c, width, top_text, current_number, top_text_y, serial_number_y):
        """套盒模板盒标外观二渲染"""
        clean_top_text = self._clean_text_for_font(top_text)
        self._set_best_font(c, 14, bold=True)
        
        # 外观二：Game title左对齐，但溢出文本居中
        max_width = width * 0.8
        title_lines = self._wrap_text_to_fit(c, clean_top_text, max_width, 14)
        
        if len(title_lines) > 1:
            # 首行左对齐，其他行居中
            for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                c.drawString(width * 0.1 + offset[0], top_text_y + 15 + offset[1], title_lines[0])
            for i, line in enumerate(title_lines[1:], 1):
                for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                    c.drawCentredString(width / 2 + offset[0], top_text_y + 15 - i * 16 + offset[1], line)
        else:
            for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                c.drawString(width * 0.1 + offset[0], top_text_y + offset[1], title_lines[0])
        
        # 绘制序列号
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], serial_number_y + offset[1], current_number)

    def _create_taohebox_small_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_small_boxes: int,
        excel_file_path: str = None,
    ):
        """创建套盒模板的小箱标 - 借鉴分盒模板的计算逻辑"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        print(f"✅ 套盒小箱标使用关键字提取: 标签名称='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 获取用户输入的分组大小（从"小箱/大箱"参数获取） - 借鉴分盒模板逻辑
        try:
            group_size = int(params["小箱/大箱"])  # 用户的第三个参数，控制副号满几进一
            if group_size <= 0:
                group_size = 2
            print(f"✅ 套盒小箱标使用用户输入分组大小: {group_size} (小箱/大箱)")
        except (ValueError, KeyError) as e:
            print(f"⚠️ 获取小箱/大箱参数失败: {e}")
            group_size = 2  # 默认分组大小
        
        # 计算参数 - 借鉴分盒模板逻辑
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"套盒小箱标-1到{total_small_boxes}")
        c.setSubject("Taohebox Small Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的套盒小箱标
        for small_box_num in range(1, total_small_boxes + 1):
            if small_box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算套盒模板的序列号范围 - 简化逻辑
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]
                base_main_num = int(match.group(1))  # 主号
                
                # 套盒模板小箱标的简化逻辑：
                # 每个小箱标对应一个主号，包含连续的boxes_per_small_box个副号
                # 第1个小箱：主号base_main_num，副号01-06
                # 第2个小箱：主号base_main_num+1，副号01-06
                # 第3个小箱：主号base_main_num+2，副号01-06
                
                current_main_number = base_main_num + (small_box_num - 1)  # 当前小箱对应的主号
                
                # 副号始终从01开始，到boxes_per_small_box结束
                start_suffix = 1
                end_suffix = boxes_per_small_box
                
                start_serial = f"{prefix_part}{current_main_number:05d}-{start_suffix:02d}"
                end_serial = f"{prefix_part}{current_main_number:05d}-{end_suffix:02d}"
                
                # 套盒小箱标显示序列号范围
                serial_range = f"{start_serial}-{end_serial}"
                print(f"📝 小箱标 #{small_box_num}: 主号{current_main_number}, 副号{start_suffix}-{end_suffix} = {serial_range}")
            else:
                serial_range = f"DSK{small_box_num:05d}-DSK{small_box_num:05d}"

            # 计算套盒小箱标的Carton No（简单序号格式，不是主-副箱号格式）
            carton_no = small_box_num  # 套盒使用简单的递增序号

            # 绘制套盒小箱标表格
            self._draw_taohebox_small_box_table(c, width, height, theme_text, pieces_per_small_box, 
                                               serial_range, carton_no, remark_text)

        c.save()
        print(f"✅ 套盒模板小箱标PDF已生成: {output_path}")

    def _draw_taohebox_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                                      serial_range, carton_no, remark_text):
        """绘制套盒小箱标表格 - 直接复用分盒模板的表格绘制逻辑"""
        # 直接调用分盒模板的表格绘制方法，只修改数据内容
        # 套盒模板需要的特殊数据：
        # - Item: "Paper Cards" (固定值)
        # - Carton No: 格式为 "01", "02" 等，而不是 "1-1" 格式
        
        # 修正Carton No格式：套盒使用简单的序号格式
        simple_carton_no = f"{carton_no:02d}" if isinstance(carton_no, int) else str(carton_no).zfill(2)
        
        # 直接使用分盒模板的绘制逻辑
        self._draw_fenhe_small_box_table(c, width, height, theme_text, pieces_per_small_box, 
                                       serial_range, simple_carton_no, remark_text)

    def _create_taohebox_large_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_large_boxes: int,
        excel_file_path: str = None,
    ):
        """创建套盒模板的大箱标 - 借鉴分盒模板的大箱标逻辑"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        print(f"✅ 套盒大箱标使用关键字提取: 标签名称='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])  # 每个大箱包含的小箱数量
        
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        pieces_per_large_box = pieces_per_small_box * small_boxes_per_large_box
        
        print(f"✅ 套盒大箱标参数: 张/盒={pieces_per_box}, 盒/小箱={boxes_per_small_box}, 小箱/大箱={small_boxes_per_large_box}")
        print(f"✅ 计算结果: 每小箱{pieces_per_small_box}PCS, 每大箱{pieces_per_large_box}PCS")
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"套盒大箱标-1到{total_large_boxes}")
        c.setSubject("Taohebox Large Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的套盒大箱标
        for large_box_num in range(1, total_large_boxes + 1):
            if large_box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算当前大箱包含的小箱范围
            start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
            end_small_box = start_small_box + small_boxes_per_large_box - 1
            
            # 计算序列号范围 - 从第一个小箱的起始号到最后一个小箱的结束号
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]
                base_main_num = int(match.group(1))  # 主号
                
                # 第一个小箱的序列号范围
                first_main_number = base_main_num + (start_small_box - 1)
                first_start_serial = f"{prefix_part}{first_main_number:05d}-01"
                
                # 最后一个小箱的序列号范围
                last_main_number = base_main_num + (end_small_box - 1)
                last_end_serial = f"{prefix_part}{last_main_number:05d}-{boxes_per_small_box:02d}"
                
                # 大箱标显示完整序列号范围
                serial_range = f"{first_start_serial}-{last_end_serial}"
                print(f"📝 大箱标 #{large_box_num}: 包含小箱{start_small_box}-{end_small_box}, 序列号范围={serial_range}")
            else:
                serial_range = f"DSK{large_box_num:05d}-DSK{large_box_num:05d}"

            # 计算套盒大箱标的Carton No（小箱范围格式）
            carton_range = f"{start_small_box}-{end_small_box}"

            # 绘制套盒大箱标表格 - 直接复用分盒大箱标的绘制逻辑
            self._draw_taohebox_large_box_table(c, width, height, theme_text, pieces_per_large_box, 
                                               serial_range, carton_range, remark_text)

        c.save()
        print(f"✅ 套盒模板大箱标PDF已生成: {output_path}")

    def _draw_taohebox_large_box_table(self, c, width, height, theme_text, pieces_per_large_box, 
                                      serial_range, carton_range, remark_text):
        """绘制套盒大箱标表格 - 直接复用分盒大箱标的表格绘制逻辑"""
        # 分盒大箱标方法的参数：(c, width, height, theme_text, pieces_per_box, small_boxes_per_large_box, serial_range, carton_no, remark_text)
        # 套盒需要传递：pieces_per_large_box作为总PCS，carton_range作为Carton No
        # 中间的small_boxes_per_large_box参数在分盒中用于计算，在套盒中我们直接传递总PCS，设为1即可
        self._draw_fenhe_large_box_table(c, width, height, theme_text, pieces_per_large_box, 
                                        1, serial_range, carton_range, remark_text)