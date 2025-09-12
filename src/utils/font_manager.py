"""
字体管理工具类
专门负责字体注册和字体相关操作
"""

import os
import sys
import platform
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


class FontManager:
    """字体管理工具类，负责字体注册和管理"""
    
    def __init__(self):
        """初始化字体管理器"""
        self.font_name = "MicrosoftYaHei"  # 默认字体名称
        self.chinese_font_name = "MicrosoftYaHei"  # 中文字体名称
        self.bold_font_name = "MicrosoftYaHei-Bold"  # 粗体字体名称
        self.font_registered = False
        self.bold_font_registered = False
        
    def register_chinese_font(self):
        """
        注册中文字体（常规和粗体）
        
        Returns:
            bool: 字体注册是否成功
        """
        if self.font_registered and self.bold_font_registered:
            return True
            
        try:
            # 获取项目字体路径
            regular_paths, bold_paths = self._get_font_paths()
            
            # 注册常规字体
            if not self.font_registered:
                for font_path in regular_paths:
                    if os.path.exists(font_path):
                        try:
                            pdfmetrics.registerFont(TTFont(self.font_name, font_path))
                            print(f"[OK] 成功注册常规中文字体: {font_path}")
                            self.font_registered = True
                            break
                        except Exception as e:
                            print(f"[WARNING] 常规字体注册失败 {font_path}: {str(e)}")
                            continue
            
            # 注册粗体字体
            if not self.bold_font_registered:
                for font_path in bold_paths:
                    if os.path.exists(font_path):
                        try:
                            if font_path.endswith('.ttc'):
                                # TTC文件需要指定字体索引，微软雅黑粗体通常是索引0
                                pdfmetrics.registerFont(TTFont(self.bold_font_name, font_path, subfontIndex=0))
                            else:
                                pdfmetrics.registerFont(TTFont(self.bold_font_name, font_path))
                            print(f"[OK] 成功注册粗体中文字体: {font_path}")
                            self.bold_font_registered = True
                            break
                        except Exception as e:
                            print(f"[WARNING] 粗体字体注册失败 {font_path}: {str(e)}")
                            continue

            # 如果没有找到合适的字体，使用Helvetica作为fallback
            if not self.font_registered:
                print("[WARNING] 未找到中文字体，将使用默认字体")
                self.font_name = "Helvetica"
                self.chinese_font_name = "Helvetica"
                
            if not self.bold_font_registered:
                print("[WARNING] 未找到粗体字体，粗体将使用常规字体")
                self.bold_font_name = self.font_name
                
            return self.font_registered

        except Exception as e:
            print(f"[WARNING] 字体注册过程出错: {str(e)}")
            self.font_name = "Helvetica"
            self.chinese_font_name = "Helvetica"
            self.bold_font_name = "Helvetica-Bold"
            return False
    
    def _get_font_paths(self) -> tuple:
        """
        获取项目fonts目录下的字体路径列表
        考虑打包后的路径兼容性
        
        Returns:
            tuple: (常规字体路径列表, 粗体字体路径列表)
        """
        regular_paths = []
        bold_paths = []
        
        # 方法1: PyInstaller打包环境 - 从临时目录读取
        try:
            if getattr(sys, 'frozen', False):
                # PyInstaller打包后，字体文件在临时目录中
                base_path = sys._MEIPASS
                regular_path = os.path.join(base_path, "fonts", "msyh.ttf")
                bold_path = os.path.join(base_path, "fonts", "msyhbd.ttc")
                regular_paths.append(regular_path)
                bold_paths.append(bold_path)
                print(f"[INFO] PyInstaller模式，常规字体路径: {regular_path}")
                print(f"[INFO] PyInstaller模式，粗体字体路径: {bold_path}")
        except Exception as e:
            print(f"[WARNING] PyInstaller路径查找失败: {e}")
        
        # 方法2: 开发环境 - 基于当前文件路径
        try:
            src_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            fonts_dir = os.path.join(src_dir, "fonts")
            regular_path = os.path.join(fonts_dir, "msyh.ttf")
            bold_path = os.path.join(fonts_dir, "msyhbd.ttc")
            regular_paths.append(regular_path)
            bold_paths.append(bold_path)
            print(f"[INFO] 开发环境，常规字体路径: {regular_path}")
            print(f"[INFO] 开发环境，粗体字体路径: {bold_path}")
        except Exception as e:
            print(f"[WARNING] 开发环境路径查找失败: {e}")
        
        # 方法3: 相对于当前工作目录
        try:
            regular_path = os.path.join("src", "fonts", "msyh.ttf")
            bold_path = os.path.join("src", "fonts", "msyhbd.ttc")
            regular_paths.append(regular_path)
            bold_paths.append(bold_path)
            print(f"[INFO] 相对路径，常规字体路径: {regular_path}")
            print(f"[INFO] 相对路径，粗体字体路径: {bold_path}")
        except Exception as e:
            print(f"[WARNING] 相对路径查找失败: {e}")
            
        # 去重并返回
        unique_regular = []
        unique_bold = []
        for path in regular_paths:
            if path not in unique_regular:
                unique_regular.append(path)
        for path in bold_paths:
            if path not in unique_bold:
                unique_bold.append(path)
                
        print(f"[INFO] 所有可能的常规字体路径: {unique_regular}")
        print(f"[INFO] 所有可能的粗体字体路径: {unique_bold}")
        return unique_regular, unique_bold
    
    def set_best_font(self, canvas_obj, font_size: int, bold: bool = True):
        """
        设置最适合的字体
        
        Args:
            canvas_obj: ReportLab Canvas对象
            font_size: 字体大小
            bold: 是否加粗
        """
        try:
            if bold and self.bold_font_registered:
                canvas_obj.setFont(self.bold_font_name, font_size)
            else:
                canvas_obj.setFont(self.chinese_font_name, font_size)
        except Exception as e:
            print(f"[WARNING] 设置字体失败 {e}，使用常规字体")
            canvas_obj.setFont(self.chinese_font_name, font_size)
    
    def has_chinese(self, text: str) -> bool:
        """
        检查文本是否包含中文字符
        
        Args:
            text: 要检查的文本
            
        Returns:
            bool: 是否包含中文字符
        """
        if not text:
            return False
        for char in text:
            if '\u4e00' <= char <= '\u9fff':
                return True
        return False
    
    def get_font_name(self) -> str:
        """获取当前字体名称"""
        return self.font_name
    
    def get_chinese_font_name(self) -> str:
        """获取中文字体名称"""
        return self.chinese_font_name
    
    def is_font_registered(self) -> bool:
        """检查字体是否已注册"""
        return self.font_registered


# 全局字体管理器实例
font_manager = FontManager()