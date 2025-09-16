"""
常规模板专属数据处理器
封装常规模板的所有数据提取和处理逻辑，基于BaseDataProcessor重构
"""

from typing import Any, Dict

from src.utils.base_data_processor import BaseDataProcessor


class RegularDataProcessor(BaseDataProcessor):
    """常规模板专属数据处理器 - 封装现有逻辑，确保功能完全一致"""

    def __init__(self):
        """初始化常规数据处理器"""
        super().__init__()
        self.template_type = "regular"

    def extract_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取常规盒标所需的数据 - 使用基类的通用方法
        """
        return self.extract_common_data(excel_file_path)

    def extract_small_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取常规小箱标所需的数据 - 使用基类的通用方法
        """
        return self.extract_common_data(excel_file_path)

    def extract_large_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取常规大箱标所需的数据 - 使用基类的通用方法
        """
        return self.extract_common_data(excel_file_path)

    def parse_serial_number_format(self, serial_number: str) -> Dict[str, Any]:
        """
        解析序列号格式 - 常规模板使用简单的线性递增逻辑
        重写基类方法以适应常规模板的特殊格式
        """
        return self.parse_simple_serial_format(serial_number)

    def calculate_multi_level_quantities(
        self,
        total_pieces: int,
        pieces_per_box: int,
        boxes_per_small_box: int,
        small_boxes_per_large_box: int,
    ) -> Dict[str, int]:
        """
        计算各级数量 - 常规模板特有的余数计算
        重写基类方法以包含常规模板的余数信息
        """
        # 使用基类的基本计算
        quantities = super().calculate_multi_level_quantities(
            total_pieces,
            pieces_per_box,
            boxes_per_small_box,
            small_boxes_per_large_box,
        )

        # 添加常规模板特有的余数信息
        quantities.update(
            {
                "remaining_pieces_in_last_box": total_pieces % pieces_per_box,
                "remaining_boxes_in_last_small_box": quantities["total_boxes"]
                % boxes_per_small_box,
                "remaining_small_boxes_in_last_large_box": quantities[
                    "total_small_boxes"
                ]
                % small_boxes_per_large_box,
            }
        )

        return quantities

    def generate_linear_serial_number(self, base_number: str, box_num: int) -> str:
        """
        生成常规盒标的序列号 - 实现基类抽象方法
        常规模板使用简单的线性递增逻辑
        """
        serial_info = self.parse_serial_number_format(base_number)

        current_number = serial_info["start_number"] + (box_num - 1)
        formatted_number = (
            f"{serial_info['prefix']}{current_number:0{serial_info['digits']}d}"
        )

        print(f"📝 常规盒标 #{box_num}: {formatted_number}")
        return formatted_number

    def generate_linear_serial_range(
        self,
        base_number: str,
        container_num: int,
        items_per_container: int,
        container_type: str = "小箱",
        total_items: int = None,
    ) -> str:
        """
        生成常规模板序列号范围 - 实现基类抽象方法
        使用线性递增逻辑计算序列号范围
        """
        serial_info = self.parse_serial_number_format(base_number)

        start_item = (container_num - 1) * items_per_container + 1
        end_item = start_item + items_per_container - 1

        if total_items is not None:
            end_item = min(end_item, total_items)

        first_serial_num = serial_info["start_number"] + (start_item - 1)
        last_serial_num = serial_info["start_number"] + (end_item - 1)

        first_serial = (
            f"{serial_info['prefix']}{first_serial_num:0{serial_info['digits']}d}"
        )
        last_serial = (
            f"{serial_info['prefix']}{last_serial_num:0{serial_info['digits']}d}"
        )

        serial_range = f"{first_serial}-{last_serial}"

        print(
            f"📝 常规{container_type}标 #{container_num}: 包含盒{start_item}-{end_item}, 序列号范围={serial_range}"
        )
        return serial_range

    # 为保持向后兼容，添加原有方法名的包装器
    def generate_regular_box_serial_number(self, base_number: str, box_num: int) -> str:
        """向后兼容包装器"""
        return self.generate_linear_serial_number(base_number, box_num)

    def generate_regular_small_box_serial_range(
        self,
        base_number: str,
        small_box_num: int,
        boxes_per_small_box: int,
        total_boxes: int = None,
    ) -> str:
        """向后兼容包装器"""
        return self.generate_linear_serial_range(
            base_number,
            small_box_num,
            boxes_per_small_box,
            "小箱",
            total_boxes,
        )

    def generate_regular_large_box_serial_range(
        self,
        base_number: str,
        large_box_num: int,
        small_boxes_per_large_box: int,
        boxes_per_small_box: int,
        total_boxes: int = None,
    ) -> str:
        """生成常规大箱标序列号范围"""
        start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
        end_small_box = start_small_box + small_boxes_per_large_box - 1

        start_box = (start_small_box - 1) * boxes_per_small_box + 1
        end_box = end_small_box * boxes_per_small_box

        if total_boxes is not None:
            end_box = min(end_box, total_boxes)

        return self.generate_linear_serial_range(
            base_number,
            large_box_num,
            end_box - start_box + 1,
            "大箱",
            total_boxes,
        )

    # ==================== 常规模板特有方法 ====================

    def calculate_carton_number_for_small_box(
        self, small_box_num: int, total_small_boxes: int
    ) -> str:
        """计算常规小箱标的Carton No - 格式：第几小箱/总小箱数"""
        return f"{small_box_num}/{total_small_boxes}"

    def calculate_carton_range_for_large_box(
        self, large_box_num: int, total_large_boxes: int
    ) -> str:
        """计算常规大箱标的Carton No - 格式：第几大箱/总大箱数"""
        return f"{large_box_num}/{total_large_boxes}"

    def calculate_pieces_for_small_box(
        self,
        small_box_num: int,
        total_small_boxes: int,
        pieces_per_small_box: int,
        remaining_pieces: int,
    ) -> int:
        """计算常规小箱的实际数量 - 处理最后一个小箱的余数情况"""
        if small_box_num == total_small_boxes and remaining_pieces > 0:
            return remaining_pieces
        return pieces_per_small_box

    def calculate_pieces_for_large_box(
        self,
        large_box_num: int,
        total_large_boxes: int,
        pieces_per_large_box: int,
        remaining_pieces: int,
    ) -> int:
        """计算常规大箱的实际数量 - 处理最后一个大箱的余数情况"""
        if large_box_num == total_large_boxes and remaining_pieces > 0:
            return remaining_pieces
        return pieces_per_large_box

    # ==================== 基类抽象方法实现 ====================

    def generate_grouped_serial_number(
        self, base_number: str, box_num: int, group_size: int
    ) -> str:
        """
        生成分组序列号 - 常规模板不使用分组，委托给线性算法
        """
        # 常规模板使用简单的线性递增，忽略group_size参数
        return self.generate_linear_serial_number(base_number, box_num)

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
        生成分组序列号范围 - 常规模板不使用分组，委托给线性算法
        """
        # 常规模板使用简单的线性递增，忽略group_size参数
        return self.generate_linear_serial_range(
            base_number,
            container_num,
            items_per_container,
            container_type,
            total_items,
        )


# 创建全局实例供regular模板使用
regular_data_processor = RegularDataProcessor()
