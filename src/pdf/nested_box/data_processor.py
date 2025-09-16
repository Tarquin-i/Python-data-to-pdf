"""
套盒模板专属数据处理器
封装套盒模板的所有数据提取和处理逻辑，基于通用基类重构
"""

import math
import re
from typing import Any, Dict

# 导入基类和工具
from src.utils.base_data_processor import BaseDataProcessor


class NestedBoxDataProcessor(BaseDataProcessor):
    """套盒模板专属数据处理器 - 封装现有逻辑，确保功能完全一致"""

    def __init__(self):
        """初始化套盒数据处理器"""
        super().__init__()
        self.template_type = "nested_box"

    def extract_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """提取套盒盒标所需的数据"""
        return self.extract_common_label_data(excel_file_path)

    def extract_small_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """提取套盒套标所需的数据"""
        return self.extract_common_label_data(excel_file_path)

    def extract_large_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """提取套盒箱标所需的数据"""
        return self.extract_common_label_data(excel_file_path)

    def _fallback_keyword_extraction_for_box_label(
        self, excel_file_path: str
    ) -> Dict[str, Any]:
        """回退到关键字提取方式"""
        return self.extract_fallback_data(
            excel_file_path, self.get_fallback_keyword_config()
        )

    # 使用基类的parse_serial_number_format方法

    def calculate_quantities(
        self,
        total_pieces: int,
        pieces_per_box: int,
        boxes_per_small_box: int,
        small_boxes_per_large_box: int,
    ) -> Dict[str, int]:
        """计算各级数量"""
        return self.calculate_multi_level_quantities(
            total_pieces,
            pieces_per_box,
            boxes_per_small_box,
            small_boxes_per_large_box,
        )

    def generate_box_serial_number(
        self, base_number: str, box_num: int, boxes_per_ending_unit: int
    ) -> str:
        """
        生成套盒盒标的序列号 - 与原有逻辑完全一致
        对应原来 _create_nested_box_label 中的序列号生成逻辑
        """
        serial_info = self.parse_serial_number_format(base_number)

        # 套盒模板序列号生成逻辑 - 基于开始号和结束号范围（与原代码完全一致）
        box_index = box_num - 1

        # 计算当前盒的序列号在范围内的位置
        main_offset = box_index // boxes_per_ending_unit
        suffix_in_range = (box_index % boxes_per_ending_unit) + serial_info["suffix"]

        current_main = serial_info["main_number"] + main_offset
        current_number = (
            f"{serial_info['prefix']}{current_main:05d}-{suffix_in_range:02d}"
        )

        print(f"📝 生成套盒盒标 #{box_num}: {current_number}")
        return current_number

    def generate_small_box_serial_range(
        self,
        base_number: str,
        small_box_num: int,
        boxes_per_small_box: int,
        total_boxes: int = None,
    ) -> str:
        """
        生成套盒套标的序列号范围 - 修复边界计算问题
        对应原来 _create_nested_small_box_label 中的序列号范围计算逻辑
        添加total_boxes边界检查，确保序列号不超出实际盒数
        """
        match = re.search(r"(\d+)", base_number)
        if match:
            # 获取第一个数字（主号）的起始位置
            digit_start = match.start()
            # 截取主号前面的所有字符作为前缀
            prefix_part = base_number[:digit_start]
            base_main_num = int(match.group(1))  # 主号

            # 套盒模板套标的简化逻辑：
            # 每个套标对应一个主号，包含连续的boxes_per_small_box个副号
            current_main_number = base_main_num + (
                small_box_num - 1
            )  # 当前套对应的主号

            # 计算当前套实际包含的盒数范围
            start_box = (small_box_num - 1) * boxes_per_small_box + 1
            end_box = start_box + boxes_per_small_box - 1

            # 🔧 边界检查：确保end_box不超过总盒数
            if total_boxes is not None:
                end_box = min(end_box, total_boxes)

            # 副号从01开始，根据实际盒数计算结束副号
            start_suffix = 1
            actual_boxes_in_small_box = end_box - start_box + 1
            end_suffix = start_suffix + actual_boxes_in_small_box - 1

            start_serial = f"{prefix_part}{current_main_number:05d}-{start_suffix:02d}"
            end_serial = f"{prefix_part}{current_main_number:05d}-{end_suffix:02d}"

            # 始终显示为范围格式，即使首尾序列号相同
            serial_range = f"{start_serial}-{end_serial}"

            print(
                f"📝 套盒套标 #{small_box_num}: 主号{current_main_number}, 副号{start_suffix}-{end_suffix}, 包含盒{start_box}-{end_box} = {serial_range}"
            )
            return serial_range
        else:
            return f"DSK{small_box_num:05d}-DSK{small_box_num:05d}"

    def generate_large_box_serial_range(
        self,
        base_number: str,
        large_box_num: int,
        small_boxes_per_large_box: int,
        boxes_per_small_box: int,
        total_boxes: int = None,
    ) -> str:
        """
        生成套盒箱标的序列号范围 - 修复边界计算问题
        对应原来 _create_nested_large_box_label 中的序列号范围计算逻辑
        添加total_boxes边界检查，确保序列号不超出实际盒数
        """
        # 计算当前箱包含的套范围
        start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
        end_small_box = start_small_box + small_boxes_per_large_box - 1

        # 计算当前箱包含的总盒子范围
        start_box = (start_small_box - 1) * boxes_per_small_box + 1
        end_box = end_small_box * boxes_per_small_box

        # 🔧 边界检查：确保end_box不超过总盒数
        if total_boxes is not None:
            end_box = min(end_box, total_boxes)
            # 根据实际的end_box重新计算最后一个套
            actual_end_small_box = math.ceil(end_box / boxes_per_small_box)
            end_small_box = min(end_small_box, actual_end_small_box)

        # 计算序列号范围 - 从第一个套的起始号到最后一个套的结束号
        match = re.search(r"(\d+)", base_number)
        if match:
            # 获取第一个数字（主号）的起始位置
            digit_start = match.start()
            # 截取主号前面的所有字符作为前缀
            prefix_part = base_number[:digit_start]
            base_main_num = int(match.group(1))  # 主号

            # 第一个套的序列号范围
            first_main_number = base_main_num + (start_small_box - 1)
            first_start_serial = f"{prefix_part}{first_main_number:05d}-01"

            # 最后一个套的序列号范围（考虑边界）
            last_main_number = base_main_num + (end_small_box - 1)
            last_box_in_small_box = end_box - (end_small_box - 1) * boxes_per_small_box
            last_suffix = min(boxes_per_small_box, last_box_in_small_box)
            last_end_serial = f"{prefix_part}{last_main_number:05d}-{last_suffix:02d}"

            # 始终显示为范围格式，即使首尾序列号相同
            serial_range = f"{first_start_serial}-{last_end_serial}"

            print(
                f"📝 套盒箱标 #{large_box_num}: 包含套{start_small_box}-{end_small_box}, 盒{start_box}-{end_box}, 序列号范围={serial_range}"
            )
            return serial_range
        else:
            return f"DSK{large_box_num:05d}-DSK{large_box_num:05d}"

    def calculate_carton_number_for_small_box(self, small_box_num: int) -> str:
        """计算套盒套标的Carton No - 与原有逻辑完全一致"""
        return str(small_box_num)

    def calculate_carton_range_for_large_box(
        self, large_box_num: int, small_boxes_per_large_box: int
    ) -> str:
        """计算套盒箱标的Carton No范围 - 与原有逻辑完全一致"""
        start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
        end_small_box = start_small_box + small_boxes_per_large_box - 1
        return f"{start_small_box}-{end_small_box}"

    def calculate_overweight_box_distribution(
        self, boxes_per_set: int, boxes_per_large_box: int, box_in_set: int
    ) -> int:
        """计算超重模式下每箱的盒子分配"""
        return self.calculate_overweight_distribution(
            boxes_per_set, boxes_per_large_box, box_in_set
        )

    def generate_overweight_serial_range(
        self,
        base_number: str,
        set_num: int,
        box_in_set: int,
        boxes_per_set: int,
        boxes_per_large_box: int,
    ) -> str:
        """
        生成超重模式的序列号范围
        使用套盒模板的正确逻辑：副号先递增，满"盒/套"参数时主号进一，副号重置

        Args:
            base_number: 基础序列号
            set_num: 套编号（1-based）
            box_in_set: 箱在套内的编号（1-based）
            boxes_per_set: 每套包含的盒数
            boxes_per_large_box: 一套拆成多少箱

        Returns:
            序列号范围字符串
        """
        # 计算当前箱包含的盒数
        boxes_in_current_box = self.calculate_overweight_box_distribution(
            boxes_per_set, boxes_per_large_box, box_in_set
        )

        # 计算当前箱的起始盒编号（在当前套内，1-based）
        start_box_in_set = 0
        for i in range(1, box_in_set):
            start_box_in_set += self.calculate_overweight_box_distribution(
                boxes_per_set, boxes_per_large_box, i
            )
        start_box_in_set += 1  # 转换为1-based

        # 计算结束盒编号（在当前套内）
        end_box_in_set = start_box_in_set + boxes_in_current_box - 1

        # 解析基础序列号格式
        match = re.search(r"(.+?)(\d+)-(\d+)", base_number)
        if match:
            start_prefix = match.group(1)
            base_main_num = int(match.group(2))
            start_suffix = int(match.group(3))

            # 套盒模板序列号逻辑：副号先递增，满"盒/套"参数时主号进一
            # 计算当前套的主号
            current_main = base_main_num + (set_num - 1)

            # 计算起始和结束副号（在当前套内）
            start_suffix_in_set = start_suffix + (start_box_in_set - 1)
            end_suffix_in_set = start_suffix + (end_box_in_set - 1)

            start_serial = f"{start_prefix}{current_main:05d}-{start_suffix_in_set:02d}"
            end_serial = f"{start_prefix}{current_main:05d}-{end_suffix_in_set:02d}"

            # 始终显示为范围格式
            serial_range = f"{start_serial}-{end_serial}"

            print(
                f"📝 超重箱标 套{set_num}-箱{box_in_set}: 主号{current_main}, 副号{start_suffix_in_set}-{end_suffix_in_set}, 包含盒{start_box_in_set}-{end_box_in_set}(套内), 序列号={serial_range}"
            )
            return serial_range
        else:
            print("⚠️ 无法解析序列号格式，使用默认逻辑")
            return f"DSK{set_num:05d}-{box_in_set:02d}"

    # ==================== 实现抽象方法 ====================

    def get_template_specific_config(self) -> Dict[str, Any]:
        """返回套盒模板特定的配置信息"""
        return {
            "template_name": "套盒模板",
            "supports_overweight": True,
            "supports_small_box": True,
            "default_prefix": "JAW",
            "default_main_number": 1001,
            "default_suffix": 1,
        }

    def get_fallback_keyword_config(self) -> Dict[str, Dict]:
        """返回回退关键字配置"""
        return {
            "标签名称": {"keyword": "标签名称", "direction": "right"},
            "开始号": {"keyword": "开始号", "direction": "down"},
            "结束号": {"keyword": "结束号", "direction": "down"},
            "客户编码": {"keyword": "客户名称编码", "direction": "down"},
        }

    # ==================== 基类抽象方法实现 ====================

    def generate_linear_serial_number(self, base_number: str, box_num: int) -> str:
        """
        生成线性序列号 - 套盒模板不使用线性序列号，委托给复杂算法
        """
        # 套盒模板使用复杂的分组算法，默认每套1个盒
        return self.generate_box_serial_number(base_number, box_num, 1)

    def generate_linear_serial_range(
        self,
        base_number: str,
        container_num: int,
        items_per_container: int,
        container_type: str = "小箱",
        total_items: int = None,
    ) -> str:
        """
        生成线性序列号范围 - 套盒模板不使用线性序列号，委托给复杂算法
        """
        if container_type == "小箱":
            return self.generate_small_box_serial_range(
                base_number, container_num, items_per_container, total_items
            )
        else:
            return self.generate_large_box_serial_range(
                base_number, container_num, 1, items_per_container, total_items
            )

    def generate_grouped_serial_number(
        self, base_number: str, box_num: int, group_size: int
    ) -> str:
        """
        生成分组序列号 - 套盒模板的核心序列号生成方法
        """
        return self.generate_box_serial_number(base_number, box_num, group_size)

    def generate_grouped_serial_range(
        self,
        base_number: str,
        container_num: int,
        items_per_container: int,
        group_size: int,
        container_type: str = "小箱",
        total_items: int = None,
    ) -> str:
        """
        生成分组序列号范围 - 套盒模板的核心范围生成方法
        """
        if container_type == "小箱":
            return self.generate_small_box_serial_range(
                base_number, container_num, items_per_container, total_items
            )
        else:
            return self.generate_large_box_serial_range(
                base_number,
                container_num,
                group_size,
                items_per_container,
                total_items,
            )


# 创建全局实例供nested_box模板使用
nested_box_data_processor = NestedBoxDataProcessor()
