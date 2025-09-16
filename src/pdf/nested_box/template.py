"""
Nested Box Template - Multi-level PDF generation with Excel serial number ranges
"""
import math
from datetime import datetime
from pathlib import Path
from typing import Dict, Any
from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
from reportlab.lib.units import mm

# 导入基础工具类
from src.utils.pdf_base import PDFBaseUtils
from src.utils.font_manager import font_manager
from src.utils.text_processor import text_processor
from src.utils.excel_data_extractor import ExcelDataExtractor

# 导入套盒模板专属数据处理器和渲染器
from src.pdf.nested_box.data_processor import nested_box_data_processor
from src.pdf.nested_box.renderer import nested_box_renderer


class NestedBoxTemplate(PDFBaseUtils):
    """Nested Box Template Handler Class"""
    
    def __init__(self):
        """Initialize Nested Box Template"""
        super().__init__()
    
    def create_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        Create multi-level PDF labels for nested box template

        Args:
            data: Excel数据
            params: 用户参数 (张/盒, 盒/套, 套/箱, 选择外观, 是否超重)
            output_dir: 输出目录
            excel_file_path: Excel文件路径

        Returns:
            生成的文件路径字典
        """
        # 检查是否为超重模式
        is_overweight = params.get("是否超重", False)
        
        if is_overweight:
            return self._create_overweight_pdfs(data, params, output_dir, excel_file_path)
        else:
            return self._create_normal_pdfs(data, params, output_dir, excel_file_path)
    
    def _create_normal_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """正常模式：多套装箱，生成盒标+套标+箱标"""
        # 检查是否有小箱
        has_small_box = params.get("是否有小箱", True)
        
        # 计算数量 - 三级结构：张→盒→套→箱 或 二级结构：张→盒→箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/套"])  # 这个参数用于确定结束号
        small_boxes_per_large_box = int(params["套/箱"])

        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        if has_small_box:
            total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
            total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)
        else:
            # 无小箱模式：直接从盒到箱
            total_large_boxes = math.ceil(total_boxes / boxes_per_small_box)

        # 创建输出目录
        clean_theme = data['标签名称'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户名称编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        # 获取参数和日期时间戳
        chinese_name = params.get("中文名称", "")
        english_name = clean_theme  # 英文名称使用清理后的主题
        customer_code = data['客户名称编码']  # 客户编号
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        generated_files = {}

        # 检查是否需要生成盒标
        has_box_label = params.get("是否有盒标", False)
        
        if has_box_label:
            # 生成套盒模板的盒标 - 第二个参数用于结束号
            selected_appearance = params["选择外观"]
            # 文件名格式：客户编号_中文名称_英文名称_套盒盒标_日期时间戳
            box_label_filename = f"{customer_code}_{chinese_name}_{english_name}_套盒盒标_{timestamp}.pdf"
            box_label_path = full_output_dir / box_label_filename

            self._create_nested_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
            generated_files["盒标"] = str(box_label_path)
        else:
            print("⏭️ 用户选择无盒标，跳过盒标生成")

        # 只有在有小箱模式下才生成套标
        if has_small_box:
            # 生成套盒模板套标
            # 文件名格式：客户编号_中文名称_英文名称_套盒套标_日期时间戳
            small_box_filename = f"{customer_code}_{chinese_name}_{english_name}_套盒套标_{timestamp}.pdf"
            small_box_path = full_output_dir / small_box_filename
            self._create_nested_small_box_label(
                data, params, str(small_box_path), excel_file_path
            )
            generated_files["套标"] = str(small_box_path)

        # 生成套盒模板箱标
        # 文件名格式：客户编号_中文名称_英文名称_套盒箱标_日期时间戳
        large_box_filename = f"{customer_code}_{chinese_name}_{english_name}_套盒箱标_{timestamp}.pdf"
        large_box_path = full_output_dir / large_box_filename
        self._create_nested_large_box_label(
            data, params, str(large_box_path), excel_file_path
        )
        generated_files["箱标"] = str(large_box_path)

        return generated_files
    
    def _create_overweight_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """超重模式：一套拆多箱，生成盒标+箱标（无套标）"""
        # 计算数量 - 二级结构：张→盒→箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_set = int(params["盒/套"])  # 每套包含的盒数
        boxes_per_large_box = int(params["套/箱"])  # 一套拆成多少箱（即每箱的盒数）

        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_sets = math.ceil(total_boxes / boxes_per_set)
        
        # 超重模式：每套拆成多箱，所以总箱数 = 套数 × 每套拆成的箱数
        total_large_boxes = total_sets * boxes_per_large_box

        # 创建输出目录
        clean_theme = data['标签名称'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户名称编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        # 获取参数和日期时间戳
        chinese_name = params.get("中文名称", "")
        english_name = clean_theme  # 英文名称使用清理后的主题
        customer_code = data['客户名称编码']  # 客户编号
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        generated_files = {}

        # 检查是否需要生成盒标
        has_box_label = params.get("是否有盒标", False)
        
        if has_box_label:
            # 生成套盒模板的盒标（与正常模式相同）
            selected_appearance = params["选择外观"]
            # 文件名格式：客户编号_中文名称_英文名称_套盒盒标_日期时间戳
            box_label_filename = f"{customer_code}_{chinese_name}_{english_name}_套盒盒标_{timestamp}.pdf"
            box_label_path = full_output_dir / box_label_filename

            self._create_nested_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
            generated_files["盒标"] = str(box_label_path)
        else:
            print("⏭️ 用户选择无盒标，跳过盒标生成")

        # 超重模式不生成套标，直接生成箱标（使用大箱标的逻辑）
        # 文件名格式：客户编号_中文名称_英文名称_套盒箱标_日期时间戳
        large_box_filename = f"{customer_code}_{chinese_name}_{english_name}_套盒箱标_{timestamp}.pdf"
        large_box_path = full_output_dir / large_box_filename
        self._create_overweight_large_box_label(
            data, params, str(large_box_path), excel_file_path
        )
        generated_files["箱标"] = str(large_box_path)

        return generated_files

    def _create_nested_box_label(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None
    ):
        """创建套盒模板的盒标 - 基于Excel文件的开始号和结束号"""
        # 分析Excel文件获取套盒特有的数据
        excel_path = excel_file_path
        print(f"🔍 正在分析套盒模板Excel文件: {excel_path}")
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        print(f"✅ 套盒盒标使用统一数据: 主题='{theme_text}', 开始号='{base_number}'")
        
        # 套盒模板参数分析
        pieces_per_box = int(params["张/盒"])
        boxes_per_ending_unit = int(params["盒/套"])  # 在套盒模板中，这个参数用于结束号的范围计算
        group_size = int(params["套/箱"])
        
        print(f"✅ 套盒模板参数:")
        print(f"   张/盒: {pieces_per_box}")
        print(f"   盒/套(结束号范围): {boxes_per_ending_unit}")
        print(f"   套/箱(分组大小): {group_size}")
        
        # 解析开始号的格式
        import re
        start_match = re.search(r'(.+?)(\d+)-(\d+)', base_number)
        
        if start_match:
            start_prefix = start_match.group(1)
            start_main = int(start_match.group(2))
            start_suffix = int(start_match.group(3))
            
            print(f"✅ 解析序列号格式:")
            print(f"   开始: {start_prefix}{start_main:05d}-{start_suffix:02d}")
            print(f"   结束范围由用户参数控制: {boxes_per_ending_unit}")
            
        else:
            print("⚠️ 无法解析序列号格式，使用默认逻辑")
            start_prefix = "JAW"
            start_main = 1001
            start_suffix = 1
        
        # 计算需要生成的盒标数量
        total_pieces = int(float(data["总张数"]))
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        c.setTitle("套盒模板盒标")
        
        width, height = self.page_size
        blank_height = height / 5
        top_text_y = height - 1.5 * blank_height
        serial_number_y = height - 3.5 * blank_height
        
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 生成套盒盒标 - 基于开始号到结束号的范围
        print(f"📝 开始生成套盒盒标，预计生成 {total_boxes} 个标签")
        
        # 获取中文名称用于空白首页
        chinese_name = params.get("中文名称", "")
        
        # 添加第一页空白标签（仅在处理第一个盒标时）
        nested_box_renderer.render_blank_box_first_page(c, width, height, chinese_name)
        
        c.showPage()
        c.setFillColor(cmyk_black)
        
        for box_num in range(1, total_boxes + 1):
            # 保持与原始版本一致的页面管理逻辑
            if box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 套盒模板序列号生成逻辑 - 基于开始号和结束号范围
            box_index = box_num - 1
            
            # 计算当前盒的序列号在范围内的位置
            main_offset = box_index // boxes_per_ending_unit
            suffix_in_range = (box_index % boxes_per_ending_unit) + start_suffix
            
            current_main = start_main + main_offset
            current_number = f"{start_prefix}{current_main:05d}-{suffix_in_range:02d}"
            
            print(f"📝 生成套盒盒标 #{box_num}: {current_number}")
            
            # 渲染外观
            if style == "外观一":
                nested_box_renderer.render_nested_appearance_one(c, width, theme_text, current_number, top_text_y, serial_number_y)
            else:
                nested_box_renderer.render_nested_appearance_two(c, width, theme_text, current_number, top_text_y, serial_number_y)

        c.save()
        print(f"✅ 套盒模板盒标PDF已生成: {output_path}")

    def _create_nested_small_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        excel_file_path: str = None,
    ):
        """创建套盒模板的套标 - 借鉴分盒模板的计算逻辑"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        # 获取序列号字体大小参数
        serial_font_size = int(params.get("序列号字体大小", 10))
        print(f"✅ 套盒套标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}', 序列号字体大小={serial_font_size}")
        
        # 套盒模板不需要复杂的分组逻辑，直接使用简化逻辑
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/套"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 计算套数量
        total_pieces = int(float(data["总张数"]))
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"套盒套标-1到{total_small_boxes}")
        c.setSubject("Taohebox Set Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成套标专用空白标签（第一页）- 根据标签模版类型选择函数
        chinese_name = params.get("中文名称", "")
        template_type = params.get("标签模版", "有纸卡备注")
        
        # 套标使用专门的小箱标空白标签（区分纸卡类型）
        if template_type == "有纸卡备注":
            nested_box_renderer.render_empty_small_box_label(c, width, height, chinese_name, remark_text, has_paper_card=True)
        else:  # "无纸卡备注"
            nested_box_renderer.render_empty_small_box_label(c, width, height, chinese_name, remark_text, has_paper_card=False)
        
        c.showPage()
        c.setFillColor(cmyk_black)

        # 生成指定范围的套盒套标
        for small_box_num in range(1, total_small_boxes + 1):
            if small_box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 🔧 使用修复后的数据处理器计算序列号范围（包含边界检查）
            serial_range = nested_box_data_processor.generate_small_box_serial_range(
                base_number, small_box_num, boxes_per_small_box, total_boxes
            )

            # 🔧 计算当前套的实际张数（考虑最后一套的边界情况）
            # 计算当前套实际包含的盒数
            start_box = (small_box_num - 1) * boxes_per_small_box + 1
            end_box = min(start_box + boxes_per_small_box - 1, total_boxes)
            actual_boxes_in_small_box = end_box - start_box + 1
            actual_pieces_in_small_box = actual_boxes_in_small_box * pieces_per_box

            # 计算套盒套标的Carton No（简单的套编号）
            carton_no = str(small_box_num)

            # 获取标签模版类型 - 参照分盒模版的实现方式
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制套盒套标表格（使用实际张数，根据模版类型选择函数）
            if template_type == "有纸卡备注":
                nested_box_renderer.draw_nested_small_box_table(c, width, height, theme_text, actual_pieces_in_small_box, 
                                                                 serial_range, carton_no, remark_text, serial_font_size)
            else:  # "无纸卡备注"
                nested_box_renderer.draw_nested_small_box_table_no_paper_card(c, width, height, theme_text, actual_pieces_in_small_box, 
                                                                 serial_range, carton_no, remark_text, serial_font_size)

        c.save()
        print(f"✅ 套盒模板套标PDF已生成: {output_path}")

    def _create_nested_large_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, excel_file_path: str = None):
        """创建套盒模板的箱标"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        # 获取序列号字体大小参数
        serial_font_size = int(params.get("序列号字体大小", 10))
        print(f"✅ 套盒箱标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}', 序列号字体大小={serial_font_size}")
        
        # 获取参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/套"])
        small_boxes_per_large_box = int(params["套/箱"])
        
        print(f"✅ 套盒箱标参数: 张/盒={pieces_per_box}, 盒/套={boxes_per_small_box}, 套/箱={small_boxes_per_large_box}")
        
        # 计算每套和每箱的数量
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        pieces_per_large_box = pieces_per_small_box * small_boxes_per_large_box
        
        print(f"✅ 计算结果: 每套{pieces_per_small_box}PCS, 每箱{pieces_per_large_box}PCS")
        
        # 计算箱数量
        total_pieces = int(float(data["总张数"]))
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF属性
        c.setPageCompression(1)
        c.setTitle(f"套盒箱标-1到{total_large_boxes}")
        c.setSubject("Taohebox Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成箱标专用空白标签（第一页）- 根据标签模版类型选择函数
        chinese_name = params.get("中文名称", "")
        template_type = params.get("标签模版", "有纸卡备注")
        
        # 箱标使用专门的大箱标空白标签（区分纸卡类型）
        if template_type == "有纸卡备注":
            nested_box_renderer.render_empty_large_box_label(c, width, height, chinese_name, remark_text, has_paper_card=True)
        else:  # "无纸卡备注"
            nested_box_renderer.render_empty_large_box_label(c, width, height, chinese_name, remark_text, has_paper_card=False)
        
        c.showPage()
        c.setFillColor(cmyk_black)

        # 生成箱标
        for large_box_num in range(1, total_large_boxes + 1):
            if large_box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 🔧 使用修复后的数据处理器计算序列号范围（包含边界检查）
            serial_range = nested_box_data_processor.generate_large_box_serial_range(
                base_number, large_box_num, small_boxes_per_large_box, boxes_per_small_box, total_boxes
            )

            # 🔧 计算当前箱的实际张数（考虑最后一箱的边界情况）
            # 计算当前箱实际包含的盒数
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
            start_box = (large_box_num - 1) * boxes_per_large_box + 1
            end_box = min(start_box + boxes_per_large_box - 1, total_boxes)
            actual_boxes_in_large_box = end_box - start_box + 1
            actual_pieces_in_large_box = actual_boxes_in_large_box * pieces_per_box

            # 计算套盒箱标的Carton No（套范围格式）
            start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
            end_small_box = start_small_box + small_boxes_per_large_box - 1
            carton_range = f"{start_small_box}-{end_small_box}"

            # 获取标签模版类型 - 参照分盒模版的实现方式
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制套盒箱标表格（使用实际张数，根据模版类型选择函数）
            if template_type == "有纸卡备注":
                nested_box_renderer.draw_nested_large_box_table(c, width, height, theme_text, actual_pieces_in_large_box, 
                                                                 serial_range, carton_range, remark_text, serial_font_size)
            else:  # "无纸卡备注"
                nested_box_renderer.draw_nested_large_box_table_no_paper_card(c, width, height, theme_text, actual_pieces_in_large_box, 
                                                                 serial_range, carton_range, remark_text, serial_font_size)

        c.save()
        print(f"✅ 套盒模板箱标PDF已生成: {output_path}")

    def _create_overweight_large_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, excel_file_path: str = None):
        """创建超重模式的箱标（使用套号-箱号格式的Carton No）"""
        # 使用统一数据处理后的标准四字段
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        # 获取序列号字体大小参数
        serial_font_size = int(params.get("序列号字体大小", 10))
        print(f"✅ 超重模式箱标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}', 序列号字体大小={serial_font_size}")
        
        # 获取参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_set = int(params["盒/套"])  # 每套包含的盒数
        boxes_per_large_box = int(params["套/箱"])  # 一套拆成多少箱（即每箱的盒数）
        
        print(f"✅ 超重模式箱标参数: 张/盒={pieces_per_box}, 盒/套={boxes_per_set}, 一套拆{boxes_per_large_box}箱")
        
        # 计算数量
        total_pieces = int(float(data["总张数"]))
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_sets = math.ceil(total_boxes / boxes_per_set)
        total_large_boxes = total_sets * boxes_per_large_box
        
        print(f"✅ 超重模式计算结果: 总盒={total_boxes}, 总套={total_sets}, 总箱={total_large_boxes}")
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF属性
        c.setPageCompression(1)
        c.setTitle(f"套盒超重箱标-1到{total_large_boxes}")
        c.setSubject("Nested Box Overweight Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成空箱标签（第一页）
        chinese_name = params.get("中文名称", "")
        template_type = params.get("标签模版", "有纸卡备注")
        
        # 箱标使用专门的大箱标空白标签（区分纸卡类型）
        if template_type == "有纸卡备注":
            nested_box_renderer.render_empty_large_box_label(c, width, height, chinese_name, remark_text, has_paper_card=True)
        else:  # "无纸卡备注"
            nested_box_renderer.render_empty_large_box_label(c, width, height, chinese_name, remark_text, has_paper_card=False)
        
        c.showPage()
        c.setFillColor(cmyk_black)

        # 生成超重模式箱标
        for large_box_num in range(1, total_large_boxes + 1):
            if large_box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算当前箱属于哪一套
            set_num = ((large_box_num - 1) // boxes_per_large_box) + 1
            box_in_set = ((large_box_num - 1) % boxes_per_large_box) + 1
            
            # 计算当前箱包含的盒子范围（使用超重模式的拆分逻辑）
            boxes_in_each_box = nested_box_data_processor.calculate_overweight_box_distribution(
                boxes_per_set, boxes_per_large_box, box_in_set
            )
            
            # 计算序列号范围
            serial_range = nested_box_data_processor.generate_overweight_serial_range(
                base_number, set_num, box_in_set, boxes_per_set, boxes_per_large_box
            )
            
            # 计算实际张数
            actual_pieces_in_large_box = boxes_in_each_box * pieces_per_box
            
            # 超重模式的Carton No格式：套号-箱号
            carton_no = f"{set_num}-{box_in_set}"

            # 绘制超重模式箱标表格
            if template_type == "有纸卡备注":
                nested_box_renderer.draw_nested_large_box_table(c, width, height, theme_text, actual_pieces_in_large_box, 
                                                                 serial_range, carton_no, remark_text, serial_font_size)
            else:  # "无纸卡备注"
                nested_box_renderer.draw_nested_large_box_table_no_paper_card(c, width, height, theme_text, actual_pieces_in_large_box, 
                                                                 serial_range, carton_no, remark_text, serial_font_size)

        c.save()
        print(f"✅ 超重模式箱标PDF已生成: {output_path}")



