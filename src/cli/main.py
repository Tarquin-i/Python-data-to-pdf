"""
命令行主程序入口

处理命令行参数解析和主要流程控制
"""

import click
from pathlib import Path


@click.command()
@click.argument("input_files", nargs=-1, type=click.Path(exists=True))
@click.option(
    "--input",
    "-i",
    "input_file",
    help="输入Excel文件路径",
    type=click.Path(exists=True),
)
@click.option("--template", "-t", default="basic", help="使用的模板名称 (默认: basic)")
@click.option("--output", "-o", "output_dir", help="输出目录路径", type=click.Path())
@click.option(
    "--multi-level", "-m", is_flag=True, help="生成多级标签PDF (盒标、小箱标、大箱标)"
)
@click.option("--pieces-per-box", default=50, help="张/盒 (默认: 50)", type=int)
@click.option("--boxes-per-small-box", default=5, help="盒/小箱 (默认: 5)", type=int)
@click.option(
    "--small-boxes-per-large-box", default=10, help="小箱/大箱 (默认: 10)", type=int
)
@click.option(
    "--appearance",
    default="外观一",
    help="盒标外观选择 (外观一/外观二)",
    type=click.Choice(["外观一", "外观二"]),
)
@click.version_option(version="0.1.0", prog_name="data-to-pdf")
def main(
    input_files,
    input_file,
    template,
    output_dir,
    multi_level,
    pieces_per_box,
    boxes_per_small_box,
    small_boxes_per_large_box,
    appearance,
):
    """
    Excel数据到PDF标签打印工具

    支持拖拽Excel文件: data-to-pdf file1.xlsx file2.xlsx
    或使用选项: data-to-pdf --input file.xlsx
    """
    # 处理拖拽的文件
    target_file = None
    if input_files:
        target_file = input_files[0]  # 使用第一个拖拽的文件
    elif input_file:
        target_file = input_file

    click.echo("🚀 欢迎使用 Data to PDF Print 工具!")

    # 显示当前配置
    click.echo("✨ 当前配置:")
    click.echo(f"   输入文件: {target_file or '未指定'}")
    click.echo(f"   使用模板: {template}")
    click.echo(f"   输出目录: {output_dir or '默认输出目录'}")
    if multi_level:
        click.echo("   多级模式: 开启")
        click.echo(f"   张/盒: {pieces_per_box}")
        click.echo(f"   盒/小箱: {boxes_per_small_box}")
        click.echo(f"   小箱/大箱: {small_boxes_per_large_box}")
        click.echo(f"   选择外观: {appearance}")

    # 检查输入文件
    if target_file:
        file_path = Path(target_file)
        click.echo("📂 文件信息:")
        click.echo(f"   文件名: {file_path.name}")
        click.echo(f"   文件大小: {file_path.stat().st_size} 字节")

        # 检查文件扩展名
        if file_path.suffix.lower() in [".xlsx", ".xls"]:
            click.echo("✅ Excel文件格式正确")
        else:
            click.echo("⚠️  警告: 文件可能不是Excel格式")
    else:
        click.echo("💡 提示: 使用 --input 参数指定Excel文件")
        click.echo("   例如: data-to-pdf --input data.xlsx")

    if target_file:
        try:
            # 导入处理模块
            import sys
            import os

            # 项目根目录已在setup.py中配置，直接导入即可

            from src.pdf.generator import PDFGenerator
            import pandas as pd

            click.echo("🔄 正在处理Excel文件...")

            # 读取Excel数据
            df = pd.read_excel(target_file, header=None)
            total_count = df.iloc[3, 5]

            # 提取数据
            pdf_data = {
                "客户编码": df.iloc[3, 0],
                "主题": df.iloc[3, 1],
                "排列要求": df.iloc[3, 2],
                "订单数量": df.iloc[3, 3],
                "张/盒": df.iloc[3, 4],
                "总张数": total_count,
            }

            click.echo("📊 提取的数据:")
            for key, value in pdf_data.items():
                click.echo(f"   {key}: {value}")

            # 生成PDF
            generator = PDFGenerator()
            output_path = output_dir or "output"

            if multi_level:
                # 多级标签模式
                packaging_params = {
                    "张/盒": pieces_per_box,
                    "盒/小箱": boxes_per_small_box,
                    "小箱/大箱": small_boxes_per_large_box,
                    "选择外观": appearance,
                }

                click.echo("🔄 正在生成多级标签PDF...")
                generated_files = generator.create_multi_level_pdfs(
                    pdf_data, packaging_params, output_path
                )

                click.echo("✅ 多级标签PDF生成成功!")
                click.echo("📁 生成的文件:")
                for label_type, file_path in generated_files.items():
                    click.echo(f"   - {label_type}: {Path(file_path).name}")

                folder_name = f"{pdf_data['客户编码']}+{pdf_data['主题']}+标签"
                click.echo(f"📂 保存目录: {output_path}/{folder_name}")
            else:
                # 单一标签模式 (兼容旧版本)
                pdf_file = f"{output_path}/label_{pdf_data['客户编码']}.pdf"
                os.makedirs(output_path, exist_ok=True)
                generator.create_label_pdf(pdf_data, pdf_file)

                click.echo(f"✅ PDF生成成功: {pdf_file}")
                click.echo(f"📄 总张数: {total_count}")

        except Exception as e:
            click.echo(f"❌ 处理失败: {e}")
            return

    click.echo("👋 感谢使用!")


if __name__ == "__main__":
    main()
