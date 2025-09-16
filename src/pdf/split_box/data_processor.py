"""
分盒模板专属数据处理器
基于BaseDataProcessor重构，封装分盒模板的所有数据提取和处理逻辑
"""

from typing import Any, Dict

from src.utils.base_data_processor import BaseDataProcessor


class SplitBoxDataProcessor(BaseDataProcessor):
    """分盒模板专属数据处理器 - 封装现有逻辑，确保功能完全一致"""

    def __init__(self):
        """初始化分盒数据处理器"""
        super().__init__()
        self.template_type = "split_box"

    def extract_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取分盒盒标所需的数据 - 使用基类的通用方法
        """
        return self.extract_common_data(excel_file_path)

    def extract_small_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取分盒小箱标所需的数据 - 使用基类的通用方法
        """
        return self.extract_common_data(excel_file_path)

    def extract_large_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取分盒大箱标所需的数据 - 使用基类的通用方法
        """
        return self.extract_common_data(excel_file_path)

    def parse_serial_number_format(self, serial_number: str) -> Dict[str, Any]:
        """
        解析序列号格式 - 分盒模板使用复杂格式解析
        重写基类方法以适应分盒模板的特殊格式（主号-副号）
        """
        return self.parse_complex_serial_format(serial_number)

    def calculate_multi_level_quantities(
        self,
        total_pieces: int,
        pieces_per_box: int,
        boxes_per_small_box: int,
        small_boxes_per_large_box: int,
    ) -> Dict[str, int]:
        """
        计算各级数量 - 使用基类的标准计算方法
        分盒模板使用标准的向上取整逻辑
        """
        return super().calculate_multi_level_quantities(
            total_pieces,
            pieces_per_box,
            boxes_per_small_box,
            small_boxes_per_large_box,
        )

    def generate_grouped_serial_number(
        self, base_number: str, box_num: int, group_size: int
    ) -> str:
        """
        生成分盒盒标的序列号 - 实现基类抽象方法
        分盒模板使用复杂的主号-副号格式
        """
        serial_info = self.parse_serial_number_format(base_number)

        box_index = box_num - 1  # 转换为0-based索引

        main_increments = box_index // group_size  # 主号增加的次数
        suffix_in_group = (box_index % group_size) + 1  # 当前组内的副号（1-based）

        current_main = serial_info["main_number"] + main_increments
        current_number = (
            f"{serial_info['prefix']}{current_main:05d}-{suffix_in_group:02d}"
        )

        print(
            f"📝 分盒盒标 #{box_num}: 主号{current_main}, 副号{suffix_in_group}, 分组大小{group_size} → {current_number}"
        )
        return current_number

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
        生成分盒模板序列号范围 - 实现基类抽象方法
        使用复杂的主号-副号逻辑计算序列号范围
        """
        serial_info = self.parse_serial_number_format(base_number)

        start_item = (container_num - 1) * items_per_container + 1
        end_item = start_item + items_per_container - 1

        if total_items is not None:
            end_item = min(end_item, total_items)

        # 计算范围内第一个和最后一个的序列号
        first_item_index = start_item - 1
        first_main_increments = first_item_index // group_size
        first_suffix = (first_item_index % group_size) + 1
        first_main = serial_info["main_number"] + first_main_increments
        first_serial = f"{serial_info['prefix']}{first_main:05d}-{first_suffix:02d}"

        last_item_index = end_item - 1
        last_main_increments = last_item_index // group_size
        last_suffix = (last_item_index % group_size) + 1
        last_main = serial_info["main_number"] + last_main_increments
        last_serial = f"{serial_info['prefix']}{last_main:05d}-{last_suffix:02d}"

        serial_range = f"{first_serial}-{last_serial}"

        print(
            f"📝 分盒{container_type}标 #{container_num}: 包含盒{start_item}-{end_item}, 序列号范围={serial_range}"
        )
        return serial_range

    # 为保持向后兼容，添加原有方法名的包装器
    def generate_split_box_serial_number(
        self,
        base_number: str,
        box_num: int,
        boxes_per_small_box: int,
        small_boxes_per_large_box: int,
    ) -> str:
        """向后兼容包装器"""
        group_size = boxes_per_small_box * small_boxes_per_large_box
        return self.generate_grouped_serial_number(base_number, box_num, group_size)

    def generate_split_small_box_serial_range(
        self,
        base_number: str,
        small_box_num: int,
        boxes_per_small_box: int,
        small_boxes_per_large_box: int,
        total_boxes: int = None,
    ) -> str:
        """向后兼容包装器"""
        group_size = boxes_per_small_box * small_boxes_per_large_box
        return self.generate_grouped_serial_range(
            base_number,
            small_box_num,
            boxes_per_small_box,
            group_size,
            "小箱",
            total_boxes,
        )

    def generate_split_large_box_serial_range(
        self,
        base_number: str,
        large_box_num: int,
        small_boxes_per_large_box: int,
        boxes_per_small_box: int,
        total_boxes: int = None,
    ) -> str:
        """生成分盒大箱标序列号范围"""
        group_size = boxes_per_small_box * small_boxes_per_large_box

        # 计算当前大箱包含的小箱范围
        start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
        end_small_box = start_small_box + small_boxes_per_large_box - 1

        # 计算当前大箱包含的总盒子范围
        start_box = (start_small_box - 1) * boxes_per_small_box + 1
        end_box = end_small_box * boxes_per_small_box

        if total_boxes is not None:
            end_box = min(end_box, total_boxes)

        return self.generate_grouped_serial_range(
            base_number,
            large_box_num,
            end_box - start_box + 1,
            group_size,
            "大箱",
            total_boxes,
        )

    # ==================== 分盒模板特有方法 ====================

    def calculate_carton_number_for_small_box(
        self, small_box_num: int, small_boxes_per_large_box: int
    ) -> str:
        """计算分盒小箱标的Carton No - 格式：大箱号-小箱号"""
        large_box_num = ((small_box_num - 1) // small_boxes_per_large_box) + 1
        small_box_in_large_box = ((small_box_num - 1) % small_boxes_per_large_box) + 1
        return f"{large_box_num}-{small_box_in_large_box}"

    def calculate_carton_range_for_large_box(
        self, large_box_num: int, small_boxes_per_large_box: int
    ) -> str:
        """计算分盒大箱标的Carton No范围"""
        start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
        end_small_box = start_small_box + small_boxes_per_large_box - 1
        return f"{start_small_box}-{end_small_box}"

    def calculate_group_size(
        self, boxes_per_small_box: int, small_boxes_per_large_box: int
    ) -> int:
        """计算副号进位阈值 - 副号进位阈值 = 盒/小箱 × 小箱/大箱"""
        group_size = boxes_per_small_box * small_boxes_per_large_box
        print(
            f"✅ 分盒模板副号进位阈值: {group_size} (盒/小箱{boxes_per_small_box} × 小箱/大箱{small_boxes_per_large_box})"
        )
        return group_size

    # ==================== 基类抽象方法实现 ====================

    def generate_linear_serial_number(self, base_number: str, box_num: int) -> str:
        """
        生成线性序列号 - 分盒模板不使用线性序列号，委托给分组算法
        """
        # 分盒模板使用复杂的主号-副号格式，默认group_size=1
        return self.generate_grouped_serial_number(base_number, box_num, 1)

    def generate_linear_serial_range(
        self,
        base_number: str,
        container_num: int,
        items_per_container: int,
        container_type: str = "小箱",
        total_items: int = None,
    ) -> str:
        """
        生成线性序列号范围 - 分盒模板不使用线性序列号，委托给分组算法
        """
        # 分盒模板使用复杂的主号-副号格式，默认group_size=1
        return self.generate_grouped_serial_range(
            base_number,
            container_num,
            items_per_container,
            1,
            container_type,
            total_items,
        )


# 创建全局实例供split_box模板使用
split_box_data_processor = SplitBoxDataProcessor()
