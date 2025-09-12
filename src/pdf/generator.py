"""
PDF生成器 - 重构版本
使用委托模式将不同模板的逻辑分离到独立文件中
"""

from typing import Dict, Any
from src.pdf.regular_box.template import RegularTemplate
from src.pdf.split_box.template import SplitBoxTemplate  
from src.pdf.nested_box.template import NestedBoxTemplate


def _get_template_class(template_name: str):
    """获取模板类"""
    template_classes = {
        "regular_box": RegularTemplate,
        "split_box": SplitBoxTemplate,
        "nested_box": NestedBoxTemplate
    }
    
    if template_name not in template_classes:
        raise ValueError(f"未知的模板类型: {template_name}")
    
    return template_classes[template_name]


class PDFGenerator:
    """
    PDF生成器主类 - 重构版本
    通过委托模式调用不同的模板类
    """

    def __init__(self):
        """
        初始化PDF生成器
        """
        
        # 模板实例将在需要时延迟创建
        self._regular_template = None
        self._split_box_template = None
        self._nested_box_template = None
    
    @property
    def regular_template(self):
        """延迟创建常规模板实例"""
        if self._regular_template is None:
            RegularTemplate = _get_template_class("regular_box")
            self._regular_template = RegularTemplate()
        return self._regular_template
    
    @property
    def split_box_template(self):
        """延迟创建分盒模板实例"""
        if self._split_box_template is None:
            SplitBoxTemplate = _get_template_class("split_box")
            self._split_box_template = SplitBoxTemplate()
        return self._split_box_template
    
    @property
    def nested_box_template(self):
        """延迟创建套盒模板实例"""
        if self._nested_box_template is None:
            NestedBoxTemplate = _get_template_class("nested_box")
            self._nested_box_template = NestedBoxTemplate()
        return self._nested_box_template

    def create_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        创建常规模板的多级标签PDF
        """
        return self.regular_template.create_multi_level_pdfs(data, params, output_dir, excel_file_path)

    def create_split_box_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        Create multi-level PDF labels for split box template
        """
        return self.split_box_template.create_multi_level_pdfs(data, params, output_dir, excel_file_path)

    def create_nested_box_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        Create multi-level PDF labels for nested box template
        """
        return self.nested_box_template.create_multi_level_pdfs(data, params, output_dir, excel_file_path)

    # 保持向后兼容的一些通用方法
    def set_page_size(self, size: str):
        """设置页面尺寸"""
        # 为保持兼容性，同步设置所有模板的页面尺寸
        self.regular_template.set_page_size(size) if hasattr(self.regular_template, 'set_page_size') else None
        self.split_box_template.set_page_size(size) if hasattr(self.split_box_template, 'set_page_size') else None
        self.nested_box_template.set_page_size(size) if hasattr(self.nested_box_template, 'set_page_size') else None