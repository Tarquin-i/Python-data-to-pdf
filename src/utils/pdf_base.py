"""
PDF生成基础工具类
只负责纯PDF操作相关的基础功能
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
from reportlab.lib.units import mm
from src.utils.font_manager import font_manager


class PDFBaseUtils:
    """PDF生成基础工具类，只负责纯PDF操作相关的基础功能"""
    
    def __init__(self):
        """
        初始化PDF生成器基础配置
        """
        self.page_size = (90 * mm, 50 * mm)  # 90mm x 50mm标签尺寸
        self.margin = 2 * mm
        self.font_size = 8
        # 使用全局字体管理器
        font_manager.register_chinese_font()






