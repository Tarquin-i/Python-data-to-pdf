"""
文本处理工具类
专门负责文本清理、换行、格式化
"""

from typing import List, Optional

from reportlab.pdfgen import canvas


class TextProcessor:
    """文本处理工具类，负责文本清理、换行和格式化"""

    def __init__(self):
        """初始化文本处理器"""

    def clean_text_for_font(self, text) -> str:
        """
        清理文本以适应字体显示

        Args:
            text: 需要清理的文本

        Returns:
            清理后的文本
        """
        if not text:
            return ""

        # 移除不可打印的字符，保留空格和换行
        cleaned = "".join(
            char for char in str(text) if char.isprintable() or char.isspace()
        )
        return cleaned.strip()

    def wrap_text_to_fit(
        self,
        canvas_obj: canvas.Canvas,
        text: str,
        max_width: float,
        font_name: str,
        font_size: int,
    ) -> List[str]:
        """
        将文本包装以适应指定宽度

        Args:
            canvas_obj: ReportLab Canvas对象
            text: 要包装的文本
            max_width: 最大宽度
            font_name: 字体名称
            font_size: 字体大小

        Returns:
            包装后的文本行列表
        """
        if not text:
            return [""]

        # 先清理文本
        text = self.clean_text_for_font(text)
        words = text.split()
        lines = []
        current_line = ""

        for word in words:
            test_line = current_line + (" " if current_line else "") + word
            text_width = canvas_obj.stringWidth(test_line, font_name, font_size)

            if text_width <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                    current_line = word
                else:
                    # 如果单个词太长，强制分行
                    lines.append(word)
                    current_line = ""

        if current_line:
            lines.append(current_line)

        return lines if lines else [""]

    def has_chinese(self, text: str) -> bool:
        """
        检查文本是否包含中文字符

        Args:
            text: 要检查的文本

        Returns:
            是否包含中文字符
        """
        if not text:
            return False

        for char in text:
            if "\u4e00" <= char <= "\u9fff":
                return True
        return False

    def format_serial_number(
        self,
        prefix: str,
        main_number: int,
        suffix: int,
        main_digits: int = 5,
        suffix_digits: int = 2,
    ) -> str:
        """
        格式化序列号

        Args:
            prefix: 前缀
            main_number: 主号
            suffix: 后缀
            main_digits: 主号位数
            suffix_digits: 后缀位数

        Returns:
            格式化的序列号
        """
        return f"{prefix}{main_number:0{main_digits}d}-{suffix:0{suffix_digits}d}"

    def clean_filename(self, filename: str) -> str:
        """
        清理文件名，移除不合法字符

        Args:
            filename: 原文件名

        Returns:
            清理后的文件名
        """
        if not filename:
            return "untitled"

        # 替换不合法字符
        illegal_chars = [
            "\\",
            "/",
            ":",
            "*",
            "?",
            '"',
            "<",
            ">",
            "|",
            "!",
            "\n",
        ]
        cleaned = filename

        for char in illegal_chars:
            cleaned = cleaned.replace(char, "_")

        return cleaned.strip()

    def split_text_by_length(self, text: str, max_length: int) -> List[str]:
        """
        按指定长度分割文本

        Args:
            text: 要分割的文本
            max_length: 每行最大长度

        Returns:
            分割后的文本行列表
        """
        if not text:
            return [""]

        if len(text) <= max_length:
            return [text]

        lines = []
        current_pos = 0

        while current_pos < len(text):
            end_pos = min(current_pos + max_length, len(text))

            # 尝试在单词边界分割
            if end_pos < len(text) and text[end_pos] != " ":
                # 向前查找空格
                space_pos = text.rfind(" ", current_pos, end_pos)
                if space_pos > current_pos:
                    end_pos = space_pos

            lines.append(text[current_pos:end_pos].strip())
            current_pos = end_pos

            # 跳过空格
            while current_pos < len(text) and text[current_pos] == " ":
                current_pos += 1

        return [line for line in lines if line]  # 移除空行

    def normalize_whitespace(self, text: str) -> str:
        """
        标准化空白字符

        Args:
            text: 输入文本

        Returns:
            标准化后的文本
        """
        if not text:
            return ""

        # 替换所有空白字符为单个空格，移除首尾空白
        import re

        normalized = re.sub(r"\s+", " ", text).strip()
        return normalized

    def extract_number_from_text(self, text: str) -> Optional[int]:
        """
        从文本中提取第一个数字

        Args:
            text: 输入文本

        Returns:
            提取的数字，如果没有找到返回None
        """
        if not text:
            return None

        import re

        match = re.search(r"\d+", str(text))
        return int(match.group()) if match else None

    def format_quantity_text(self, quantity: int, unit: str = "PCS") -> str:
        """
        格式化数量文本

        Args:
            quantity: 数量
            unit: 单位

        Returns:
            格式化的数量文本
        """
        return f"{quantity}{unit}"

    def truncate_text(self, text: str, max_length: int, suffix: str = "...") -> str:
        """
        截断文本到指定长度

        Args:
            text: 输入文本
            max_length: 最大长度
            suffix: 截断后缀

        Returns:
            截断后的文本
        """
        if not text or len(text) <= max_length:
            return text

        if max_length <= len(suffix):
            return text[:max_length]

        return text[: max_length - len(suffix)] + suffix

    def extract_total_count_by_keyword(self, df):
        """
        通过关键字搜索提取总张数

        Args:
            df: pandas DataFrame对象

        Returns:
            提取的总张数
        """
        import pandas as pd

        try:
            # 搜索包含"总张数"的单元格
            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    cell_value = df.iloc[i, j]
                    if pd.notna(cell_value) and "总张数" in str(cell_value):
                        print(
                            f"✅ 找到总张数关键字: 位置({i+1},{j+1}) = '{cell_value}'"
                        )

                        # 只从下方单元格获取数值
                        if i + 1 < df.shape[0]:
                            total_value = df.iloc[i + 1, j]
                            if pd.notna(total_value):
                                print(f"✅ 从下方提取总张数: {total_value}")
                                return int(float(total_value))

            # 如果没找到关键字，返回None让用户输入
            print("❌ 未找到总张数关键字，需要用户手动输入")
            return None

        except Exception as e:
            print(f"❌ 提取总张数失败: {e}")
            return None


# 全局文本处理器实例
text_processor = TextProcessor()
