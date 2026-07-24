"""
PPT 表格自动插入工具
功能：在指定位置新增 slide 并插入原生表格
用法：python insert_tables.py <input.pptx> <output.pptx>
"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# 深蓝青色科技配色（与 vnerp 品牌一致）
COLOR_HEADER_BG = RGBColor(0x1A, 0x23, 0x32)  # 深蓝 #1a2332
COLOR_HEADER_FG = RGBColor(0xFF, 0xFF, 0xFF)  # 白
COLOR_ROW_ALT = RGBColor(0xF0, 0xF4, 0xF8)    # 浅灰蓝
COLOR_ROW_FG = RGBColor(0x1A, 0x23, 0x32)
COLOR_ACCENT = RGBColor(0x00, 0xB4, 0xD8)     # 青色 #00b4d8


def _style_table(table, header_bg=COLOR_HEADER_BG, header_fg=COLOR_HEADER_FG):
    """统一表格样式：深蓝表头 + 斑马纹行"""
    # 表头行
    for cell in table.rows[0].cells:
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_bg
        for para in cell.text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER
            for run in para.runs:
                run.font.bold = True
                run.font.color.rgb = header_fg
                run.font.size = Pt(12)
                run.font.name = "微软雅黑"
    # 数据行（斑马纹）
    for row_idx in range(1, len(table.rows)):
        row = table.rows[row_idx]
        for cell in row.cells:
            if row_idx % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = COLOR_ROW_ALT
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.color.rgb = COLOR_ROW_FG
                    run.font.size = Pt(11)
                    run.font.name = "微软雅黑"


def _set_cell_text(cell, text, bold=False, color=None, align=PP_ALIGN.LEFT):
    """设置单元格文本"""
    cell.text = str(text)
    for para in cell.text_frame.paragraphs:
        para.alignment = align
        for run in para.runs:
            run.font.bold = bold
            run.font.size = Pt(11)
            run.font.name = "微软雅黑"
            if color:
                run.font.color.rgb = color


def _add_title(slide, text, left, top, width):
    """在 slide 顶部添加标题文本框"""
    txBox = slide.shapes.add_textbox(left, top, width, Inches(0.6))
    tf = txBox.text_frame
    tf.text = text
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    for run in para.runs:
        run.font.size = Pt(24)
        run.font.bold = True
        run.font.color.rgb = COLOR_HEADER_BG
        run.font.name = "微软雅黑"


def _get_blank_layout(prs):
    """获取空白版式，索引 6 不存在时回退到第一个"""
    for i in [6, 5, 4, 3, 2, 1, 0]:
        try:
            return prs.slide_layouts[i]
        except IndexError:
            continue
    return prs.slide_layouts[0]


def insert_comparison_table(prs, title_text, headers, rows):
    """新增一页插入对比表格"""
    blank_layout = _get_blank_layout(prs)
    slide = prs.slides.add_slide(blank_layout)

    # 标题
    _add_title(slide, title_text, Inches(0.6), Inches(0.4), Inches(12))

    # 表格（3 列 N 行）
    rows_count = len(rows) + 1
    cols_count = len(headers)
    table_left = Inches(0.6)
    table_top = Inches(1.4)
    table_width = Inches(12.1)
    table_height = Inches(0.5) * rows_count

    table_shape = slide.shapes.add_table(rows_count, cols_count,
                                          table_left, table_top,
                                          table_width, table_height)
    table = table_shape.table

    # 设置列宽
    if cols_count == 3:
        table.columns[0].width = Inches(3.5)
        table.columns[1].width = Inches(4.3)
        table.columns[2].width = Inches(4.3)

    # 表头
    for i, h in enumerate(headers):
        _set_cell_text(table.cell(0, i), h, bold=True,
                       color=COLOR_HEADER_FG, align=PP_ALIGN.CENTER)
    # 数据行
    for r_idx, row_data in enumerate(rows, start=1):
        for c_idx, val in enumerate(row_data):
            align = PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT
            _set_cell_text(table.cell(r_idx, c_idx), val, align=align)

    _style_table(table)
    return slide


def insert_pricing_table(prs, title_text, headers, rows, note=None):
    """新增一页插入报价表格"""
    blank_layout = _get_blank_layout(prs)
    slide = prs.slides.add_slide(blank_layout)

    _add_title(slide, title_text, Inches(0.6), Inches(0.4), Inches(12))

    rows_count = len(rows) + 1
    cols_count = len(headers)
    table_left = Inches(0.6)
    table_top = Inches(1.4)
    table_width = Inches(12.1)
    table_height = Inches(0.55) * rows_count

    table_shape = slide.shapes.add_table(rows_count, cols_count,
                                          table_left, table_top,
                                          table_width, table_height)
    table = table_shape.table

    # 列宽
    if cols_count == 3:
        table.columns[0].width = Inches(3.0)
        table.columns[1].width = Inches(7.1)
        table.columns[2].width = Inches(2.0)

    # 表头
    for i, h in enumerate(headers):
        _set_cell_text(table.cell(0, i), h, bold=True,
                       color=COLOR_HEADER_FG, align=PP_ALIGN.CENTER)
    # 数据行
    for r_idx, row_data in enumerate(rows, start=1):
        for c_idx, val in enumerate(row_data):
            align = PP_ALIGN.CENTER if c_idx == 0 else PP_ALIGN.LEFT
            if c_idx == 2:
                align = PP_ALIGN.CENTER
            _set_cell_text(table.cell(r_idx, c_idx), val, align=align)

    _style_table(table)

    # 备注框
    if note:
        txBox = slide.shapes.add_textbox(Inches(0.6), Inches(5.5), Inches(12), Inches(1.2))
        tf = txBox.text_frame
        tf.text = note
        tf.word_wrap = True
        for para in tf.paragraphs:
            for run in para.runs:
                run.font.size = Pt(12)
                run.font.italic = True
                run.font.color.rgb = COLOR_ACCENT
                run.font.name = "微软雅黑"

    return slide


def move_slide(prs, old_index, new_index):
    """移动 slide 顺序（0-based）"""
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    slide_to_move = slides_list[old_index]
    xml_slides.remove(slide_to_move)
    xml_slides.insert(new_index, slide_to_move)


def main():
    if len(sys.argv) < 3:
        print("用法: python insert_tables.py <input.pptx> <output.pptx>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    prs = Presentation(input_path)
    total_before = len(prs.slides)
    print(f"原始页数: {total_before}")

    # 表格1：传统ERP vs vnerp 对比表
    comp_headers = ["对比维度", "传统商业印刷ERP", "vnerp 解决方案"]
    comp_rows = [
        ["软件授权费", "几十万 ~ 上百万", "0 元（开源）"],
        ["实施服务费", "10万 ~ 50万", "5.6万 ~ 15.5万"],
        ["年维保费",   "5万 ~ 20万",   "0.8万 ~ 1.5万"],
        ["定制灵活性", "受限，需原厂支持", "完全可控"],
        ["技术先进性", "老旧技术栈",   "Next.js + React 19 + TS"],
        ["行业适配",   "通用ERP改版",   "专为丝网印刷设计"],
    ]
    insert_comparison_table(prs, "对比传统方案，优势明显", comp_headers, comp_rows)
    print("✅ 已插入对比表（P10 之后）")

    # 表格2：报价方案表
    price_headers = ["服务项目", "内容", "参考报价"]
    price_rows = [
        ["基础部署", "服务器环境搭建、系统部署、数据库初始化、基础数据配置", "1.5万~3万元"],
        ["业务定制", "根据客户流程调整功能模块、报表定制、字段调整",       "2万~8万元"],
        ["数据迁移", "从旧系统/Excel导入历史数据",                         "0.5万~1.5万元"],
        ["员工培训", "现场/远程操作培训（2-3天）",                          "0.8万~1.5万元"],
        ["年度运维", "技术支持、bug修复、安全更新",                         "0.8万~1.5万元/年"],
    ]
    price_note = "💡 合计：5.6万~15.5万元\n💡 SaaS订阅制可按 200~500元/人/月 或 1万~15万/年 报价，根据企业规模灵活调整。"
    insert_pricing_table(prs, "透明报价，按需选择", price_headers, price_rows, price_note)
    print("✅ 已插入报价表（P13 之后）")

    # 把新增的 2 页（当前在末尾）移动到结束页之前
    # 新增页索引为 total_before 和 total_before+1（0-based）
    # 结束页索引为 total_before - 1（0-based，原最后一页）
    # 目标：新增 2 页插到结束页之前
    end_idx = total_before - 1  # 结束页原索引
    # 新增的第1页（对比表）当前在 total_before，移到 end_idx 位置
    move_slide(prs, total_before, end_idx)
    # 新增的第2页（报价表）现在也在末尾（total_before+1 的位置，但移除一个后索引变化）
    # 移动后，报价表在 total_before 位置（末尾），结束页在 total_before+1
    # 需要把报价表移到 end_idx+1 位置
    move_slide(prs, total_before, end_idx + 1)

    prs.save(output_path)
    total_after = len(prs.slides)
    print(f"最终页数: {total_after}")
    print(f"✅ 已保存: {output_path}")


if __name__ == "__main__":
    main()
