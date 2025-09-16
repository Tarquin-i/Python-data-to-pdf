"""
通用Excel数据提取器
根据关键字动态查找并提取数据
"""

from typing import Any, Dict, List, Optional

import pandas as pd


class ExcelDataExtractor:
    """
    Excel数据提取器
    通过关键字查找对应的数据位置
    """

    def __init__(self, file_path: str):
        """
        初始化提取器

        Args:
            file_path: Excel文件路径
        """
        self.file_path = file_path
        self.df = None
        self.keyword_positions = {}
        self._load_excel()

    def _load_excel(self):
        """加载Excel文件"""
        try:
            # 首先尝试读取第一个工作表（默认）
            self.df = pd.read_excel(
                self.file_path, header=None, sheet_name=0, engine="openpyxl"
            )
            print(
                f"✅ Excel文件已加载: {self.df.shape[0]}行 x {self.df.shape[1]}列 (工作表: 0)"
            )

            # 显示前几行内容用于调试
            print("📋 Excel前5行内容预览:")
            for i in range(min(5, self.df.shape[0])):
                row_content = []
                for j in range(min(5, self.df.shape[1])):
                    cell_value = self.df.iloc[i, j]
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip()
                        row_content.append(f"[{j}]='{cell_str}'")
                    else:
                        row_content.append(f"[{j}]=<空>")
                print(f"   行{i+1}: {' '.join(row_content)}")

        except Exception as e:
            raise Exception(f"无法加载Excel文件: {e}")

    def find_keyword(self, keyword: str) -> List[Dict[str, Any]]:
        """
        查找关键字在Excel中的位置

        Args:
            keyword: 要查找的关键字

        Returns:
            包含位置信息的字典列表
        """
        if self.df is None:
            return []

        positions = []
        print(f"🔍 搜索关键字: '{keyword}'")

        # 记录所有包含关键字的单元格，用于调试
        potential_matches = []

        for row_idx in range(self.df.shape[0]):
            for col_idx in range(self.df.shape[1]):
                cell_value = self.df.iloc[row_idx, col_idx]

                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()

                    # 检查是否包含关键字（用于调试）
                    if keyword in cell_str:
                        potential_matches.append(
                            f"({row_idx+1},{col_idx+1})='{cell_str}'"
                        )

                        # 精确匹配或包含匹配的逻辑
                        match_found = False

                        # 优先精确匹配
                        if cell_str == keyword:
                            match_found = True
                            print(
                                f"   ✅ 精确匹配在({row_idx+1},{col_idx+1}): '{cell_str}'"
                            )
                        # 其次包含匹配
                        elif keyword in cell_str:
                            match_found = True
                            print(
                                f"   ✅ 包含匹配在({row_idx+1},{col_idx+1}): '{cell_str}'"
                            )

                        if match_found:
                            col_letter = self._col_index_to_letter(col_idx)
                            positions.append(
                                {
                                    "row": row_idx,
                                    "col": col_idx,
                                    "excel_ref": f"{col_letter}{row_idx + 1}",
                                    "value": cell_str,
                                    "keyword": keyword,
                                }
                            )

        # 调试输出
        if potential_matches:
            print(f"   📍 发现 {len(potential_matches)} 个包含关键字的单元格:")
            for match in potential_matches[:5]:  # 只显示前5个
                print(f"      {match}")
            if len(potential_matches) > 5:
                print(f"      ... 还有 {len(potential_matches) - 5} 个")
        else:
            print(f"   ❌ 未找到包含关键字 '{keyword}' 的单元格")

        print(f"   📊 最终匹配结果: {len(positions)} 个位置")

        return positions

    def _col_index_to_letter(self, col_idx: int) -> str:
        """
        将列索引转换为Excel列字母

        Args:
            col_idx: 列索引 (0-based)

        Returns:
            Excel列字母 (如 'A', 'B', 'AA')
        """
        result = ""
        while col_idx >= 0:
            result = chr(65 + col_idx % 26) + result
            col_idx = col_idx // 26 - 1
        return result

    def get_nearby_value(self, row: int, col: int, direction: str) -> Optional[Any]:
        """
        获取指定位置附近的值

        Args:
            row: 行索引
            col: 列索引
            direction: 方向 ('right', 'down', 'left', 'up', 'right_down', etc.)

        Returns:
            附近单元格的值
        """
        if self.df is None:
            return None

        direction_map = {
            "right": (0, 1),
            "down": (1, 0),
            "left": (0, -1),
            "up": (-1, 0),
            "right_down": (1, 1),
            "left_down": (1, -1),
            "right_up": (-1, 1),
            "left_up": (-1, -1),
        }

        if direction not in direction_map:
            return None

        dr, dc = direction_map[direction]
        new_row, new_col = row + dr, col + dc

        # 检查边界
        if 0 <= new_row < self.df.shape[0] and 0 <= new_col < self.df.shape[1]:
            value = self.df.iloc[new_row, new_col]
            return value if pd.notna(value) else None

        return None

    def extract_data_by_keywords(
        self, keyword_config: Dict[str, Dict]
    ) -> Dict[str, Any]:
        """
        根据关键字配置提取数据

        Args:
            keyword_config: 关键字配置字典
            格式: {
                'field_name': {
                    'keyword': '关键字',  # 单个关键字（向后兼容）
                    'keywords': ['关键字1', '关键字2'],  # 多个可能的关键字
                    'direction': 'right',  # 数据相对于关键字的位置
                    'offset': (0, 1)  # 可选，额外偏移
                }
            }

        Returns:
            提取的数据字典
        """
        extracted_data = {}

        for field_name, config in keyword_config.items():
            keyword = config["keyword"]
            direction = config.get("direction", "right")
            offset = config.get("offset", (0, 0))

            # 查找关键字
            positions = self.find_keyword(keyword)

            if positions:
                # 使用第一个匹配的位置
                pos = positions[0]
                row, col = pos["row"], pos["col"]

                # 应用方向偏移
                if direction == "right":
                    target_row, target_col = row, col + 1
                elif direction == "down":
                    target_row, target_col = row + 1, col
                elif direction == "left":
                    target_row, target_col = row, col - 1
                elif direction == "up":
                    target_row, target_col = row - 1, col
                else:
                    # 直接使用关键字位置
                    target_row, target_col = row, col

                # 应用额外偏移
                target_row += offset[0]
                target_col += offset[1]

                # 获取目标位置的值
                if (
                    0 <= target_row < self.df.shape[0]
                    and 0 <= target_col < self.df.shape[1]
                ):
                    value = self.df.iloc[target_row, target_col]
                    extracted_data[field_name] = value if pd.notna(value) else None

                    col_letter = self._col_index_to_letter(target_col)
                    print(
                        f"✅ {field_name}: 匹配关键字 '{keyword}', 从 {col_letter}{target_row + 1} 提取 = {value}"
                    )
                else:
                    print(f"❌ {field_name}: 目标位置超出范围")
                    extracted_data[field_name] = None
            else:
                print(f"❌ {field_name}: 未找到关键字 '{keyword}'")
                extracted_data[field_name] = None

        return extracted_data

    def extract_common_data(self) -> Dict[str, Any]:
        """
        提取所有模板都需要的公共数据：客户编码、标签名称、开始号、总张数、张/盒

        Returns:
            包含公共数据的字典
        """
        print("🔍 提取公共数据字段...")

        # 定义公共数据的关键字配置 - 精确匹配指定关键字
        keyword_config = {
            "标签名称": {"keyword": "标签名称", "direction": "right"},
            "开始号": {"keyword": "开始号", "direction": "down"},
            "客户名称编码": {"keyword": "客户名称编码", "direction": "down"},
            "张/盒": {"keyword": "张/盒", "direction": "down"},
        }

        # 使用关键字提取数据
        extracted_data = self.extract_data_by_keywords(keyword_config)

        # 提取总张数（使用专门的逻辑）
        from src.utils.text_processor import text_processor

        total_count = text_processor.extract_total_count_by_keyword(self.df)
        extracted_data["总张数"] = (
            total_count if total_count and total_count > 0 else None
        )

        # 清理提取的数据：只保留真正有效的数据，无效或空的设为None
        cleaned_data = {}
        for field in ["客户名称编码", "标签名称", "开始号", "总张数", "张/盒"]:
            value = extracted_data.get(field)
            if value is not None and str(value).strip() != "" and str(value) != "0":
                cleaned_data[field] = value
            else:
                cleaned_data[field] = None

        extracted_data = cleaned_data

        print("✅ 公共数据提取完成:")
        for key, value in extracted_data.items():
            print(f"   {key}: {value}")

        return extracted_data

    def get_unified_standard_data(self, user_supplemented_data=None):
        """
        获取统一标准化的五个字段数据
        优先使用Excel提取的数据，用户输入数据补充缺失字段
        确保三个模板获得格式一致的标准数据

        Args:
            user_supplemented_data: 用户补充的数据字典（可选）

        Returns:
            包含五个标准字段的字典，格式统一，供三个模板使用
        """
        print("🔄 开始统一数据处理...")

        # 首先从Excel提取基础数据
        excel_data = self.extract_common_data()

        # 创建标准化的五字段数据结构（统一使用Excel原始字段名）
        standard_data = {
            "客户名称编码": None,
            "标签名称": None,
            "开始号": None,
            "总张数": None,
            "张/盒": None,
        }

        # 第一步：填入Excel中提取的有效数据（严格验证数据有效性）
        if (
            excel_data.get("客户名称编码") is not None
            and str(excel_data["客户名称编码"]).strip()
        ):
            standard_data["客户名称编码"] = str(excel_data["客户名称编码"]).strip()
            print(f"✅ 客户名称编码: 从Excel提取 = '{standard_data['客户名称编码']}'")

        if (
            excel_data.get("标签名称") is not None
            and str(excel_data["标签名称"]).strip()
        ):
            standard_data["标签名称"] = str(excel_data["标签名称"]).strip()
            print(f"✅ 标签名称: 从Excel提取 = '{standard_data['标签名称']}'")

        if excel_data.get("开始号") is not None and str(excel_data["开始号"]).strip():
            standard_data["开始号"] = str(excel_data["开始号"]).strip()
            print(f"✅ 开始号: 从Excel提取 = '{standard_data['开始号']}'")

        if excel_data.get("总张数") is not None and excel_data.get("总张数") != 0:
            standard_data["总张数"] = int(excel_data["总张数"])
            print(f"✅ 总张数: 从Excel提取 = {standard_data['总张数']}")

        if excel_data.get("张/盒") is not None and excel_data.get("张/盒") != 0:
            standard_data["张/盒"] = int(excel_data["张/盒"])
            print(f"✅ 张/盒: 从Excel提取 = {standard_data['张/盒']}")

        # 第二步：用户补充数据填补缺失字段
        if user_supplemented_data:
            for field in [
                "客户名称编码",
                "标签名称",
                "开始号",
                "总张数",
                "张/盒",
            ]:
                # 检查Excel数据是否缺失，以及用户是否提供了补充数据
                excel_has_data = standard_data[field] is not None
                user_has_data = (
                    user_supplemented_data.get(field)
                    and str(user_supplemented_data.get(field)).strip()
                )

                # 如果Excel没有数据但用户提供了数据，使用用户数据
                if not excel_has_data and user_has_data:
                    if field in ["总张数", "张/盒"]:
                        standard_data[field] = int(user_supplemented_data[field])
                    else:
                        standard_data[field] = str(
                            user_supplemented_data[field]
                        ).strip()
                    print(f"✅ {field}: 从用户输入补充 = '{standard_data[field]}'")

        # 第三步：验证数据完整性
        missing_fields = [
            field for field, value in standard_data.items() if value is None
        ]
        if missing_fields:
            print(f"⚠️ 仍有缺失字段: {missing_fields}")
        else:
            print("✅ 五个标准字段数据完整")

        print("🔄 统一数据处理完成")
        print(f"📋 最终标准数据: {standard_data}")

        return standard_data


def test_extractor():
    """测试数据提取器"""
    file_path = "/Users/trq/Desktop/常规-LADIES NIGHT IN 女士夜._副本.xlsx"

    print("🚀 测试Excel数据提取器")
    print("=" * 50)

    try:
        extractor = ExcelDataExtractor(file_path)

        # 定义要提取的数据配置
        keyword_config = {
            "开始号": {
                "keyword": "开始号",
                "direction": "down",  # 开始号在B10，数据在B11
            },
            "标签名称": {
                "keyword": "标签名称:",
                "direction": "right",  # 标签名称在G11，数据在H11
            },
            "客户编码": {
                "keyword": "14KH0149",  # 直接查找客户编码
                "direction": "none",
            },
            "主题": {
                "keyword": "主题",
                "direction": "down",
            },  # 主题标签在B3，数据在B4
            "总张数": {
                "keyword": "总张数",
                "direction": "down",  # 总张数标签，数据在下方
            },
        }

        # 提取数据
        print("\n📊 开始提取数据...")
        extracted_data = extractor.extract_data_by_keywords(keyword_config)

        print("\n📋 提取结果:")
        for field, value in extracted_data.items():
            print(f"  {field}: {value}")

        return extracted_data

    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return None


if __name__ == "__main__":
    test_extractor()
