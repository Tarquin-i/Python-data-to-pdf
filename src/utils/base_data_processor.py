"""
通用数据处理基类
抽取所有模板共同的数据处理逻辑，统一序列号生成和数量计算
"""

import math
import re
from abc import ABC
from typing import Any, Dict

from src.utils.excel_data_extractor import ExcelDataExtractor


class BaseDataProcessor(ABC):
    """通用数据处理基类"""

    def __init__(self):
        """初始化基类数据处理器"""
        self.template_type = getattr(self, "template_type", "unknown")

    # ==================== 通用数据提取方法 ====================

    def extract_common_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取通用标签数据（为了向后兼容子类调用）
        所有模板都需要的基础数据：标签名称、开始号、客户编码等
        """
        extractor = ExcelDataExtractor(excel_file_path)
        return extractor.extract_common_data()

    def extract_common_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取通用标签数据
        所有模板都需要的基础数据：标签名称、开始号、客户编码等
        """
        return self.extract_common_data(excel_file_path)

    def extract_fallback_data(
        self, excel_file_path: str, keyword_config: Dict[str, Dict]
    ) -> Dict[str, Any]:
        """
        回退到关键字提取方式
        当标准提取失败时使用的备用方法
        """
        extractor = ExcelDataExtractor(excel_file_path)
        excel_data = extractor.extract_data_by_keywords(keyword_config)

        theme_text = excel_data.get("标签名称") or "Unknown Title"
        base_number = excel_data.get("开始号") or "DEFAULT01001"
        end_number = excel_data.get("结束号") or base_number

        return {
            "标签名称": theme_text,
            "开始号": base_number,
            "结束号": end_number,
            "主题": excel_data.get("主题", "Unknown Theme"),
            "客户编码": excel_data.get("客户编码", "Unknown Client"),
        }

    # ==================== 通用序列号处理方法 ====================

    def parse_simple_serial_format(self, serial_number: str) -> Dict[str, Any]:
        """
        解析简单序列号格式（为常规模板设计）
        格式：DSK00001 等简单格式
        """
        if not serial_number:
            return {"prefix": "DSK", "start_number": 1, "digits": 5}

        # 尝试解析DSK00001这种格式
        match = re.search(r"([A-Z]+)(\d+)", serial_number)
        if match:
            prefix = match.group(1)
            start_number = int(match.group(2))
            digits = len(match.group(2))

            return {
                "prefix": prefix,
                "start_number": start_number,
                "digits": digits,
            }
        else:
            # 如果无法解析，使用默认格式
            return {"prefix": "DSK", "start_number": 1, "digits": 5}

    def parse_complex_serial_format(self, serial_number: str) -> Dict[str, Any]:
        """
        解析复杂序列号格式（为分盒模板设计）
        格式：DSK12345-01 等复杂格式
        """
        # 分盒模板使用简单的数字提取逻辑（与原代码一致）
        match = re.search(r"(\d+)", serial_number)
        if match:
            # 获取第一个数字（主号）的起始位置
            digit_start = match.start()
            # 截取主号前面的所有字符作为前缀
            prefix_part = serial_number[:digit_start]
            base_main_num = int(match.group(1))  # 主号

            return {
                "prefix": prefix_part,
                "main_number": base_main_num,
                "digit_start": digit_start,
            }
        else:
            return {"prefix": "DSK", "main_number": 1001, "digit_start": 0}

    def parse_serial_number_format(self, serial_number: str) -> Dict[str, Any]:
        """
        解析序列号格式
        支持格式：PREFIX12345-01 或类似的格式
        """
        # 支持多种格式的正则表达式
        patterns = [
            r"(.+?)(\d+)-(\d+)",  # 标准格式：PREFIX12345-01
            r"(.+?)(\d+)",  # 简单格式：PREFIX12345
        ]

        for pattern in patterns:
            match = re.search(pattern, serial_number)
            if match:
                if len(match.groups()) == 3:
                    # 标准格式
                    prefix = match.group(1)
                    main_number = int(match.group(2))
                    suffix = int(match.group(3))

                    print("✅ 解析序列号格式:")
                    print(f"   开始: {prefix}{main_number:05d}-{suffix:02d}")

                    return {
                        "prefix": prefix,
                        "main_number": main_number,
                        "suffix": suffix,
                        "main_digits": 5,
                        "suffix_digits": 2,
                        "has_suffix": True,
                    }
                elif len(match.groups()) == 2:
                    # 简单格式
                    prefix = match.group(1)
                    main_number = int(match.group(2))

                    print("✅ 解析序列号格式（简单）:")
                    print(f"   开始: {prefix}{main_number:05d}")

                    return {
                        "prefix": prefix,
                        "main_number": main_number,
                        "suffix": 0,
                        "main_digits": 5,
                        "suffix_digits": 0,
                        "has_suffix": False,
                    }

        # 解析失败，使用默认格式
        print("⚠️ 无法解析序列号格式，使用默认逻辑")
        return {
            "prefix": "DSK",
            "main_number": 1001,
            "suffix": 1,
            "main_digits": 5,
            "suffix_digits": 2,
            "has_suffix": True,
        }

    def generate_serial_number(
        self, base_number: str, index: int, increment_type="suffix"
    ) -> str:
        """
        生成序列号

        Args:
            base_number: 基础序列号
            index: 索引（从0开始）
            increment_type: 递增类型 'suffix'=副号递增, 'main'=主号递增, 'both'=混合递增
        """
        serial_info = self.parse_serial_number_format(base_number)

        if increment_type == "suffix" and serial_info["has_suffix"]:
            # 副号递增
            new_suffix = serial_info["suffix"] + index
            current_number = f"{serial_info['prefix']}{serial_info['main_number']:0{serial_info['main_digits']}d}-{new_suffix:0{serial_info['suffix_digits']}d}"
        elif increment_type == "main":
            # 主号递增
            new_main = serial_info["main_number"] + index
            if serial_info["has_suffix"]:
                current_number = f"{serial_info['prefix']}{new_main:0{serial_info['main_digits']}d}-{serial_info['suffix']:0{serial_info['suffix_digits']}d}"
            else:
                current_number = (
                    f"{serial_info['prefix']}{new_main:0{serial_info['main_digits']}d}"
                )
        else:
            # 混合递增或其他情况，子类可以重写实现
            current_number = self._generate_custom_serial_number(serial_info, index)

        return current_number

    def generate_serial_range(
        self, base_number: str, start_index: int, count: int
    ) -> str:
        """
        生成序列号范围字符串

        Args:
            base_number: 基础序列号
            start_index: 起始索引
            count: 数量
        """
        if count <= 0:
            return base_number

        start_serial = self.generate_serial_number(base_number, start_index)
        if count == 1:
            return start_serial

        end_serial = self.generate_serial_number(base_number, start_index + count - 1)
        return f"{start_serial}-{end_serial}"

    def _generate_custom_serial_number(
        self, serial_info: Dict[str, Any], index: int
    ) -> str:
        """
        自定义序列号生成逻辑，子类可以重写
        """
        # 默认实现：副号递增
        if serial_info["has_suffix"]:
            new_suffix = serial_info["suffix"] + index
            return f"{serial_info['prefix']}{serial_info['main_number']:0{serial_info['main_digits']}d}-{new_suffix:0{serial_info['suffix_digits']}d}"
        else:
            new_main = serial_info["main_number"] + index
            return f"{serial_info['prefix']}{new_main:0{serial_info['main_digits']}d}"

    # ==================== 通用数量计算方法 ====================

    def calculate_basic_quantities(
        self, total_pieces: int, pieces_per_unit: int
    ) -> Dict[str, int]:
        """
        计算基础数量

        Args:
            total_pieces: 总票数
            pieces_per_unit: 每单位票数
        """
        total_units = math.ceil(total_pieces / pieces_per_unit)
        return {
            "total_pieces": total_pieces,
            "pieces_per_unit": pieces_per_unit,
            "total_units": total_units,
        }

    def calculate_multi_level_quantities(
        self,
        total_pieces: int,
        pieces_per_box: int,
        boxes_per_small_box: int,
        small_boxes_per_large_box: int = None,
    ) -> Dict[str, int]:
        """
        计算多级包装数量
        支持二级（盒→箱）或三级（盒→套→箱）包装

        Args:
            total_pieces: 总票数
            pieces_per_box: 每盒票数
            boxes_per_small_box: 每套盒数（或每箱盒数，如果是二级包装）
            small_boxes_per_large_box: 每箱套数（仅三级包装需要）
        """
        total_boxes = math.ceil(total_pieces / pieces_per_box)

        if small_boxes_per_large_box is None:
            # 二级包装：盒→箱
            total_large_boxes = math.ceil(total_boxes / boxes_per_small_box)
            return {
                "total_pieces": total_pieces,
                "total_boxes": total_boxes,
                "total_large_boxes": total_large_boxes,
                "pieces_per_box": pieces_per_box,
                "boxes_per_large_box": boxes_per_small_box,
            }
        else:
            # 三级包装：盒→套→箱
            total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
            total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)
            return {
                "total_pieces": total_pieces,
                "total_boxes": total_boxes,
                "total_small_boxes": total_small_boxes,
                "total_large_boxes": total_large_boxes,
                "pieces_per_box": pieces_per_box,
                "boxes_per_small_box": boxes_per_small_box,
                "small_boxes_per_large_box": small_boxes_per_large_box,
            }

    def calculate_overweight_distribution(
        self,
        items_per_set: int,
        sets_per_container: int,
        item_in_container: int,
    ) -> int:
        """
        计算超重模式下的分配
        使用均分+余数分配策略

        Args:
            items_per_set: 每套包含的项目数
            sets_per_container: 一套拆成多少容器
            item_in_container: 当前容器在套内的编号（1-based）
        """
        # 基础分配
        base_items = items_per_set // sets_per_container
        # 余数
        remainder = items_per_set % sets_per_container

        # 余数分配策略：前面的容器多分配一个
        if item_in_container <= remainder:
            return base_items + 1
        else:
            return base_items

    # ==================== 子类可重写的抽象方法 ====================

    def get_template_specific_config(self) -> Dict[str, Any]:
        """返回模板特定的配置信息 - 子类可以重写"""
        return {
            "template_type": self.template_type,
            "supports_overweight": False,
            "supports_empty_labels": True,
        }

    def get_fallback_keyword_config(self) -> Dict[str, Dict]:
        """返回回退关键字配置 - 子类可以重写"""
        return {
            "标签名称": {"keywords": ["标签名称", "主题", "theme", "title"]},
            "开始号": {"keywords": ["开始号", "起始序列号", "start_number"]},
            "客户编码": {"keywords": ["客户编码", "客户名称", "client_code"]},
        }
