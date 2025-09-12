"""
GUI应用程序
支持选择Excel文件进行处理
支持Windows和macOS
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from src.pdf.generator import PDFGenerator
from src.pdf.regular_box.ui_dialog import get_regular_ui_dialog
from src.pdf.split_box.ui_dialog import get_split_box_ui_dialog
from src.pdf.nested_box.ui_dialog import get_nested_box_ui_dialog
from src.utils.text_processor import text_processor
from src.utils.excel_data_extractor import ExcelDataExtractor
from src.utils.font_manager import font_manager
from src.utils.data_input_dialog import show_data_input_dialog

# 在应用启动时初始化字体管理器
print("[INFO] 初始化字体管理器...")
font_success = font_manager.register_chinese_font()
if font_success:
    print("[OK] 字体管理器初始化成功")
else:
    print("[WARNING] 字体管理器初始化失败，将使用默认字体")


class DataToPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data to PDF Print - Excel转PDF工具")
        self.root.geometry("600x500")
        self.root.resizable(True, True)

        # 设置窗口居中
        self.center_window()

        self.setup_ui()
        self.setup_file_selection()

    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 标题
        title_label = ttk.Label(
            main_frame, text="Excel数据到PDF转换工具", font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 20))

        # 文件选择区域
        self.select_frame = tk.Frame(
            main_frame, bg="#f0f0f0", relief="ridge", bd=2, height=120
        )
        self.select_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        self.select_frame.grid_propagate(False)

        # 文件选择提示
        self.select_label = tk.Label(
            self.select_frame,
            text="📁 点击此区域选择Excel文件\n\n支持 .xlsx 和 .xls 格式",
            bg="#f0f0f0",
            font=("Arial", 11),
            fg="#666666",
            cursor="hand2",
        )
        self.select_label.place(relx=0.5, rely=0.5, anchor="center")

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=(0, 20))

        # 选择文件按钮
        select_btn = ttk.Button(
            button_frame, text="📂 选择Excel文件", command=self.select_file
        )
        select_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 生成PDF按钮
        self.generate_btn = ttk.Button(
            button_frame,
            text="🔄 选择模板并生成PDF",
            command=self.start_generation_workflow,
            state="disabled",
        )
        self.generate_btn.pack(side=tk.LEFT)

        # 文件信息显示
        info_frame = ttk.LabelFrame(main_frame, text="文件信息", padding="10")
        info_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))

        self.info_text = tk.Text(info_frame, height=10, width=70, font=("Consolas", 10))
        self.info_text.pack(fill=tk.BOTH, expand=True)

        # 滚动条
        scrollbar = ttk.Scrollbar(
            info_frame, orient="vertical", command=self.info_text.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.info_text.configure(yscrollcommand=scrollbar.set)

        # 状态栏
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

        self.status_var = tk.StringVar()
        self.status_var.set("📋 准备就绪 - 请选择Excel文件")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT)

        # 配置网格权重
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        self.current_file = None
        self.current_data = None
        self.packaging_params = None


    def center_window(self):
        """窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def setup_file_selection(self):
        """设置文件选择功能"""
        # 点击区域打开文件选择
        self.select_frame.bind("<Button-1>", self.on_click_select)
        self.select_label.bind("<Button-1>", self.on_click_select)
        self.root.bind("<Control-o>", lambda e: self.select_file())  # Ctrl+O快捷键

    def on_click_select(self, event):
        """点击选择区域打开文件选择"""
        self.select_file()

    def select_file(self):
        """选择文件对话框"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if file_path:
            self.process_file(file_path)

    def process_file(self, file_path):
        """处理Excel文件"""
        try:
            self.status_var.set("🔄 正在处理文件...")
            self.info_text.delete(1.0, tk.END)
            self.info_text.insert(tk.END, "正在读取Excel文件...\n")
            self.root.update()

            # 检查文件格式
            if not file_path.lower().endswith((".xlsx", ".xls")):
                messagebox.showerror("格式错误", "请选择Excel文件(.xlsx或.xls)")
                self.status_var.set("❌ 文件格式错误")
                return

            # 使用统一的Excel数据提取器
            extractor = ExcelDataExtractor(file_path)
            
            # 先尝试获取统一标准数据（仅Excel数据）
            self.current_data = extractor.get_unified_standard_data()
            
            # 检查是否有缺失字段需要用户补充
            missing_fields = [field for field, value in self.current_data.items() if value is None]
            
            if missing_fields:
                # 有数据缺失，显示数据补充对话框
                self.status_var.set("⚠️ 数据不完整，请补充...")
                self.info_text.insert(tk.END, f"检测到{len(missing_fields)}个数据缺失，请补充：{', '.join(missing_fields)}\n")
                self.root.update()
                
                # 准备对话框显示用的数据（转换为字符串格式）
                display_data = {}
                for field, value in self.current_data.items():
                    display_data[field] = str(value) if value is not None else ''
                
                # 调用数据补充对话框
                supplemented_data = show_data_input_dialog(self.root, display_data)
                
                if supplemented_data is None:
                    # 用户取消了补充操作
                    self.status_var.set("❌ 已取消数据补充")
                    self.info_text.insert(tk.END, "用户取消了数据补充操作\n")
                    return
                
                # 使用统一数据处理方法合并Excel数据和用户输入数据
                self.current_data = extractor.get_unified_standard_data(supplemented_data)
                self.info_text.insert(tk.END, "✅ 数据补充完成\n")
            else:
                # 数据完整，已经通过统一方法处理
                self.info_text.insert(tk.END, "✅ Excel数据完整，无需补充\n")

            # 显示提取的信息
            info_text = f"文件: {Path(file_path).name}\n"
            info_text += f"文件大小: {Path(file_path).stat().st_size} 字节\n\n"
            info_text += "提取的数据:\n"
            info_text += "-" * 40 + "\n"

            for key, value in self.current_data.items():
                info_text += f"{key}: {value}\n"

            self.info_text.insert(tk.END, info_text)

            self.current_file = file_path
            self.generate_btn.config(state="normal")
            self.status_var.set("✅ 文件处理完成")

            # 更新选择区域显示 - 不再显示总张数（避免重复）
            display_text = (
                f"✅ 已选择文件: {Path(file_path).name}"
                f"\n\n点击生成多级标签PDF按钮继续"
            )
            self.select_label.config(text=display_text, fg="green")

        except Exception as e:
            error_msg = f"处理文件失败: {str(e)}"
            messagebox.showerror("处理错误", error_msg)
            self.status_var.set("❌ 处理失败")
            self.info_text.insert(tk.END, f"\n错误: {error_msg}\n")





    def _auto_resize_and_center_dialog(self, dialog, content_frame):
        """自动调整对话框大小并居中显示，完全基于内容自适应"""
        try:
            # 多次更新确保所有组件都已完全渲染
            for _ in range(3):
                dialog.update_idletasks()
                content_frame.update_idletasks()
            
            # 获取内容的实际所需尺寸
            content_width = content_frame.winfo_reqwidth()
            content_height = content_frame.winfo_reqheight()
            
            # 添加必要的边距：滚动条、对话框边框、标题栏等
            padding_width = 60   # 减少左右边距
            padding_height = 80   # 减少上下边距，让对话框更紧凑
            
            # 计算对话框所需的实际尺寸
            required_width = content_width + padding_width
            required_height = content_height + padding_height
            
            # 获取屏幕尺寸，确保不会超出屏幕
            screen_width = dialog.winfo_screenwidth()
            screen_height = dialog.winfo_screenheight()
            
            # 最终尺寸：完全基于内容，但不超过屏幕90%
            final_width = min(required_width, int(screen_width * 0.9))
            final_height = min(required_height, int(screen_height * 0.9))
            
            # 计算居中位置
            x = (screen_width - final_width) // 2
            y = (screen_height - final_height) // 2
            
            # 设置对话框几何形状
            dialog.geometry(f"{final_width}x{final_height}+{x}+{y}")
            
            print(f"✅ 完全自适应调整: {final_width}x{final_height}")
            print(f"   内容尺寸: {content_width}x{content_height}")
            print(f"   边距: {padding_width}x{padding_height}")
            
        except Exception as e:
            print(f"⚠️ 自适应调整失败: {e}")
            # 备用方案：让系统自动计算
            dialog.update_idletasks()
            dialog.geometry("")  # 清空几何设置，让Tkinter自动调整
            dialog.update_idletasks()
            
            # 获取自动调整后的尺寸并居中
            width = dialog.winfo_width()
            height = dialog.winfo_height()
            x = (dialog.winfo_screenwidth() - width) // 2
            y = (dialog.winfo_screenheight() - height) // 2
            dialog.geometry(f"{width}x{height}+{x}+{y}")

    def show_template_selection_dialog(self):
        """显示模板选择对话框"""
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("选择标签模板")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")

        # 主框架
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(
            main_frame, text="选择标签模板", font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 20))

        # 模板选择变量
        self.template_choice = tk.StringVar(value="常规")

        # 模板选择框架
        template_frame = ttk.LabelFrame(main_frame, text="模板类型", padding="15")
        template_frame.pack(fill=tk.X, pady=(0, 20))

        # 三个模板选项
        templates = [
            ("常规", "适用于普通包装标签"),
            ("分盒", "适用于分盒包装标签"),
            ("套盒", "适用于套盒包装标签")
        ]

        for i, (template_name, description) in enumerate(templates):
            radio = ttk.Radiobutton(
                template_frame, 
                text=f"{template_name} - {description}",
                variable=self.template_choice,
                value=template_name
            )
            radio.pack(anchor=tk.W, pady=5)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))

        self.selected_template = None

        def confirm_template():
            self.selected_template = self.template_choice.get()
            dialog.destroy()

        def cancel_template():
            self.selected_template = None
            dialog.destroy()

        # 确认和取消按钮
        ttk.Button(button_frame, text="确认", command=confirm_template).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(button_frame, text="取消", command=cancel_template).pack(side=tk.RIGHT)

        # 等待对话框关闭
        dialog.wait_window()
        return self.selected_template

    def start_generation_workflow(self):
        """开始生成工作流：先选择模板，再设置参数"""
        if not self.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        # 步骤1: 选择模板
        template_choice = self.show_template_selection_dialog()
        if not template_choice:
            return  # 用户取消选择

        # 保存模板选择
        self.selected_main_template = template_choice

        # 步骤2: 设置参数（根据模板调整参数界面）
        self.show_parameters_dialog_for_template(template_choice)

    def show_parameters_dialog_for_template(self, template_type):
        """根据模板类型显示对应的参数设置对话框"""
        if template_type == "常规":
            get_regular_ui_dialog(self).show_parameters_dialog()
        elif template_type == "分盒": 
            get_split_box_ui_dialog(self).show_parameters_dialog()
        elif template_type == "套盒":
            get_nested_box_ui_dialog(self).show_parameters_dialog()

    def generate_multi_level_pdfs(self):
        """生成多级标签PDF"""
        if not self.current_data or not self.packaging_params:
            messagebox.showwarning("警告", "缺少必要数据或参数")
            return

        # 使用已选择的模板
        template_choice = getattr(self, 'selected_main_template', '常规')

        try:
            self.status_var.set(f"🔄 正在生成{template_choice}模板PDF...")
            self.info_text.insert(tk.END, f"\n开始生成{template_choice}模板PDF...\n")
            self.root.update()

            # 选择输出目录
            output_dir = filedialog.askdirectory(
                title="选择输出目录", initialdir=os.path.expanduser("~/Desktop")
            )

            if output_dir:
                # 创建PDF生成器
                generator = PDFGenerator()
                
                # 根据模板选择调用不同的生成方法
                if template_choice == "常规":
                    generated_files = generator.create_multi_level_pdfs(
                        self.current_data, self.packaging_params, output_dir, self.current_file
                    )
                elif template_choice == "分盒":
                    generated_files = generator.create_split_box_multi_level_pdfs(
                        self.current_data, self.packaging_params, output_dir, self.current_file
                    )
                elif template_choice == "套盒":
                    generated_files = generator.create_nested_box_multi_level_pdfs(
                        self.current_data, self.packaging_params, output_dir, self.current_file
                    )

                self.status_var.set(f"✅ {template_choice}模板PDF生成成功!")

                # 显示生成结果
                result_text = "\n✅ 生成完成! 文件列表:\n"
                for label_type, file_path in generated_files.items():
                    result_text += f"  - {label_type}: {Path(file_path).name}\n"

                # 使用和模板中完全相同的主题清理逻辑
                clean_theme = self.current_data['标签名称'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
                folder_name = f"{self.current_data['客户名称编码']}+{clean_theme}+标签"
                result_text += (
                    f"\n📁 保存目录: {os.path.join(output_dir, folder_name)}\n"
                )

                self.info_text.insert(tk.END, result_text)

                # 询问是否打开文件夹
                if messagebox.askyesno(
                    "生成成功",
                    f"多级标签PDF已生成!\n\n保存目录: {folder_name}\n\n是否打开文件夹？",
                ):
                    import subprocess
                    import platform

                    folder_path = os.path.join(output_dir, folder_name)
                    try:
                        if platform.system() == "Darwin":  # macOS
                            result = subprocess.run(["open", folder_path], capture_output=True, text=True)
                            if result.returncode != 0:
                                messagebox.showerror("错误", f"无法打开文件夹: {result.stderr}")
                        elif platform.system() == "Windows":  # Windows
                            os.startfile(folder_path)
                        else:
                            messagebox.showinfo("提示", f"请手动打开文件夹: {folder_path}")
                    except Exception as e:
                        messagebox.showerror("错误", f"打开文件夹失败: {str(e)}")

            else:
                self.status_var.set("📋 PDF生成已取消")

        except Exception as e:
            error_msg = f"生成PDF失败: {str(e)}"
            messagebox.showerror("生成错误", error_msg)
            self.status_var.set("❌ PDF生成失败")
            self.info_text.insert(tk.END, f"\n错误: {error_msg}\n")


def main():
    """启动GUI应用"""
    root = tk.Tk()
    app = DataToPDFApp(root)

    # 检查命令行参数，支持文件关联
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if file_path.lower().endswith((".xlsx", ".xls")):
            # 延迟处理文件，等GUI完全加载
            root.after(500, lambda: app.process_file(file_path))

    root.mainloop()


if __name__ == "__main__":
    main()
