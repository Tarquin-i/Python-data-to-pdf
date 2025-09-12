"""
Excel文件读取器

负责读取Excel文件并解析数据
"""

import pandas as pd
from pathlib import Path
from typing import List, Optional


class ExcelReader:
    """
    Excel文件读取器类
    """

    def __init__(self, file_path: str):
        """
        初始化Excel读取器

        Args:
            file_path: Excel文件路径
        """
        self.file_path = Path(file_path) if isinstance(file_path, str) else file_path

        if not self.file_path.exists():
            raise FileNotFoundError(f"Excel文件不存在: {self.file_path}")

        if self.file_path.suffix.lower() not in [".xlsx", ".xls"]:
            raise ValueError(f"不支持的文件格式: {self.file_path.suffix}")

    def read_data(self, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """
        读取Excel数据

        Args:
            sheet_name: 工作表名称，None表示读取第一个工作表

        Returns:
            解析后的数据DataFrame
        """
        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            return df
        except Exception as e:
            raise Exception(f"读取Excel文件失败: {e}")

    def get_sheet_names(self) -> List[str]:
        """
        获取Excel文件中所有工作表名称

        Returns:
            工作表名称列表
        """
        try:
            excel_file = pd.ExcelFile(self.file_path)
            return excel_file.sheet_names
        except Exception as e:
            raise Exception(f"获取工作表名称失败: {e}")

    def read_multiple_sheets(self) -> dict:
        """
        读取所有工作表数据

        Returns:
            字典，键为工作表名称，值为DataFrame
        """
        try:
            return pd.read_excel(self.file_path, sheet_name=None)
        except Exception as e:
            raise Exception(f"读取多个工作表失败: {e}")
