"""
5 类核心页面标准布局（T004）

页面类型：
    1. cover        封面页：上下分割版式
    2. catalog      目录页：左标题 + 右编号列表
    3. divider      章节分隔页：左右分栏大号序号
    4. numbered_list 编号列表页：条目卡片化布局
    5. kpi          KPI 页：等宽横向卡片矩阵

设计规范：
    - 严格执行安全边距、栅格布局、字体层级、卡片圆角/阴影、间距令牌、文本自适应
    - 所有样式从主题令牌读取，禁止硬编码
    - 元素自动标记 role + shape_id（供动画模块匹配）
    - 渲染输出为原生可编辑 PPTX 元素

字段约定（outline.json page 对象）：
    - cover: title, subtitle
    - catalog: title, items[]
    - divider: title, section_no
    - numbered_list: title, items[]
    - kpi: title, kpi_items[{label, value, trend}]
"""
from __future__ import annotations

from typing import Any

from pptx.slide import Slide
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from aippt.logger import logger
from aippt.layout.ppt_auto_layout import (
    GridRect, LayoutContext, ElementMeta,
    safe_area, column_x, row_y,
    split_horizontal, split_vertical, grid_matrix,
    get_token, hex_to_rgb,
    add_text_box, add_shape, add_line, add_card, render_number_badge,
    register_layout,
)
from aippt.theme_loader import resolve_color


# ==================== cover 封面页：上下分割版式 ====================
@register_layout("cover")
def layout_cover(slide: Slide, page_data: dict, theme: dict, ctx: LayoutContext) -> None:
    """封面页布局：上下分割

    版式：
        - 上半部分（45%）：主题色背景，大号标题居中
        - 下半部分（55%）：白色背景，副标题居中
    字段：
        - title: 主标题（必须）
        - subtitle: 副标题（可选）
    元素角色：
        - cover_bg: 背景矩形
        - title: 主标题
        - subtitle: 副标题
    """
    area = safe_area(get_token(theme, "spacing.safe_margin_inch", 0.5))
    split_ratio = get_token(theme, "layout.cover_split_ratio", 0.45)
    top_rect, bottom_rect = split_vertical(area, ratio=split_ratio, gutter=0)

    # 上半部分主题色背景（铺满整宽，超出安全区到画布顶部）
    cover_bg_rect = GridRect(
        left=0, top=0,
        width=13.333,
        height=split_ratio * area.height + get_token(theme, "spacing.safe_margin_inch", 0.5),
    )
    cover_bg = add_shape(
        slide, cover_bg_rect, 1, "cover_bg", ctx, theme,  # 1 = MSO_SHAPE.RECTANGLE
        fill_color=resolve_color(theme, get_token(theme, "color.cover_bg", "#1A56DB")),
        line_color=None,
        line_width_pt=0,
    )
    # 去掉边框
    cover_bg.line.fill.background()

    # 主标题（上半部分居中）
    title_text = page_data.get("title", "")
    title_rect = GridRect(
        left=area.left, top=cover_bg_rect.top + cover_bg_rect.height * 0.3,
        width=area.width, height=cover_bg_rect.height * 0.4,
    )
    add_text_box(
        slide, title_rect, title_text, "cover_title", ctx, theme,
        font_size_pt=get_token(theme, "font.cover_title.size_pt", 44),
        bold=get_token(theme, "font.cover_title.bold", True),
        color=resolve_color(theme, get_token(theme, "color.cover_title", "#FFFFFF")),
        align="center", anchor="middle",
    )

    # 副标题（下半部分居中）
    subtitle_text = page_data.get("subtitle", "")
    if subtitle_text:
        sub_rect = GridRect(
            left=area.left, top=cover_bg_rect.height + 0.5,
            width=area.width, height=1.0,
        )
        add_text_box(
            slide, sub_rect, subtitle_text, "cover_subtitle", ctx, theme,
            font_size_pt=get_token(theme, "font.cover_subtitle.size_pt", 22),
            bold=get_token(theme, "font.cover_subtitle.bold", False),
            color=resolve_color(theme, get_token(theme, "color.cover_subtitle", "#DBEAFE")),
            align="center", anchor="middle",
        )


# ==================== catalog 目录页：左标题 + 右编号列表 ====================
@register_layout("catalog")
def layout_catalog(slide: Slide, page_data: dict, theme: dict, ctx: LayoutContext) -> None:
    """目录页布局：左标题 + 右编号列表

    版式：
        - 左侧（35%）：大号"目录"标题 + 副标题
        - 右侧（65%）：编号列表，每项含序号 + 标题
    字段：
        - title: 目录标题（如"目录"或"CONTENTS"）
        - items[]: 目录条目列表
    元素角色：
        - title: 左侧大标题
        - body: 右侧每个目录条目
        - number: 每个条目的序号
    """
    area = safe_area(get_token(theme, "spacing.safe_margin_inch", 0.5))
    left_rect, right_rect = split_horizontal(area, ratio=0.35)

    # 左侧大标题
    title_text = page_data.get("title", "目录")
    add_text_box(
        slide, GridRect(left_rect.left, left_rect.top + 1.5, left_rect.width, 1.5),
        title_text, "title", ctx, theme,
        font_size_pt=get_token(theme, "font.title.size_pt", 36),
        bold=True,
        color=resolve_color(theme, get_token(theme, "color.text_primary", "#1F2937")),
        align="left", anchor="middle",
    )

    # 左侧装饰线
    add_line(
        slide, left_rect.left, left_rect.top + 3.2,
        left_rect.left + 1.5, left_rect.top + 3.2,
        "divider", ctx, theme,
        color=resolve_color(theme, get_token(theme, "color.primary", "#1A56DB")),
        width_pt=2.0,
    )

    # 右侧编号列表
    items = page_data.get("items", [])
    if not items:
        return

    item_height = 0.7
    gap = get_token(theme, "spacing.element_gap_inch", 0.15)
    start_y = right_rect.top + 0.3
    for idx, item in enumerate(items):
        item_y = start_y + idx * (item_height + gap)
        if item_y + item_height > area.top + area.height:
            break  # 防止溢出

        # 序号徽章
        badge_size = 0.5
        badge_rect = GridRect(
            left=right_rect.left, top=item_y,
            width=badge_size, height=badge_size,
        )
        render_number_badge(
            slide, badge_rect, f"{idx + 1:02d}", "number", ctx, theme,
        )

        # 条目文本
        text_rect = GridRect(
            left=right_rect.left + badge_size + 0.2, top=item_y,
            width=right_rect.width - badge_size - 0.2, height=item_height,
        )
        add_text_box(
            slide, text_rect, str(item), "body", ctx, theme,
            font_size_pt=get_token(theme, "font.body.size_pt", 16),
            bold=False,
            color=resolve_color(theme, get_token(theme, "color.text_primary", "#1F2937")),
            align="left", anchor="middle",
        )


# ==================== divider 章节分隔页：左右分栏大号序号 ====================
@register_layout("divider")
def layout_divider(slide: Slide, page_data: dict, theme: dict, ctx: LayoutContext) -> None:
    """章节分隔页布局：左右分栏大号序号

    版式：
        - 左侧（35%）：超大号章节序号（如 01）
        - 右侧（65%）：章节标题
    字段：
        - section_no: 章节序号（如 "01" / "PART ONE"）
        - title: 章节标题
    元素角色：
        - number: 左侧大号序号
        - title: 右侧章节标题
        - divider_bg: 全页背景
    """
    area = safe_area(get_token(theme, "spacing.safe_margin_inch", 0.5))

    # 全页背景
    bg_rect = GridRect(left=0, top=0, width=13.333, height=7.5)
    bg = add_shape(
        slide, bg_rect, 1, "divider_bg", ctx, theme,
        fill_color=resolve_color(theme, get_token(theme, "color.divider_bg", "#1E40AF")),
        line_color=None,
    )
    bg.line.fill.background()

    # 左右分栏
    left_rect, right_rect = split_horizontal(area, ratio=0.35)

    # 左侧大号序号
    section_no = page_data.get("section_no", "")
    if not section_no:
        # 从 title 前缀提取（如 "01. 项目背景" → "01"）
        title = page_data.get("title", "")
        if title and len(title) >= 2 and title[:2].isdigit():
            section_no = title[:2]
        else:
            section_no = ""

    if section_no:
        add_text_box(
            slide, GridRect(left_rect.left, left_rect.top + 1.5, left_rect.width, 3.5),
            section_no, "divider_number", ctx, theme,
            font_size_pt=get_token(theme, "font.divider_number.size_pt", 72),
            bold=get_token(theme, "font.divider_number.bold", True),
            color=resolve_color(theme, get_token(theme, "color.divider_number", "#FFFFFF")),
            align="center", anchor="middle",
        )

    # 右侧章节标题
    title_text = page_data.get("title", "")
    # 去掉前缀序号
    if title_text and section_no and title_text.startswith(section_no):
        title_text = title_text[len(section_no):].lstrip(".、 ")

    add_text_box(
        slide, GridRect(right_rect.left, right_rect.top + 2.5, right_rect.width, 2.0),
        title_text, "divider_title", ctx, theme,
        font_size_pt=get_token(theme, "font.divider_title.size_pt", 40),
        bold=get_token(theme, "font.divider_title.bold", True),
        color=resolve_color(theme, get_token(theme, "color.divider_title", "#FFFFFF")),
        align="left", anchor="middle",
    )


# ==================== numbered_list 编号列表页：条目卡片化布局 ====================
@register_layout("numbered_list")
def layout_numbered_list(slide: Slide, page_data: dict, theme: dict, ctx: LayoutContext) -> None:
    """编号列表页布局：条目卡片化

    版式：
        - 顶部标题区（占 20%）
        - 下方卡片网格（2 列布局，每卡片含序号 + 标题）
    字段：
        - title: 页面标题
        - items[]: 条目列表
    元素角色：
        - title: 页面标题
        - card: 每个条目卡片
        - number: 卡片序号
        - body: 卡片文本
    """
    area = safe_area(get_token(theme, "spacing.safe_margin_inch", 0.5))

    # 顶部标题区
    title_rect = GridRect(area.left, area.top, area.width, 1.0)
    add_text_box(
        slide, title_rect, page_data.get("title", ""), "title", ctx, theme,
        font_size_pt=get_token(theme, "font.title.size_pt", 36),
        bold=True,
        color=resolve_color(theme, get_token(theme, "color.text_primary", "#1F2937")),
        align="left", anchor="middle",
    )

    # 装饰下划线
    add_line(
        slide, area.left, area.top + 1.1,
        area.left + 1.5, area.top + 1.1,
        "divider", ctx, theme,
        color=resolve_color(theme, get_token(theme, "color.primary", "#1A56DB")),
        width_pt=2.0,
    )

    # 卡片网格区域
    items = page_data.get("items", [])
    if not items:
        return

    cards_area = GridRect(
        left=area.left, top=area.top + 1.5,
        width=area.width, height=area.height - 1.5,
    )
    per_row = get_token(theme, "layout.numbered_list_card_per_row", 2)
    rows = (len(items) + per_row - 1) // per_row
    matrix = grid_matrix(cards_area, rows=rows, cols=per_row,
                         gutter=get_token(theme, "spacing.grid_gutter_inch", 0.25))

    for idx, item in enumerate(items):
        r = idx // per_row
        c = idx % per_row
        if r >= len(matrix):
            break
        cell = matrix[r][c]

        # 卡片背景
        card = add_card(slide, cell, "card", ctx, theme,
                        radius=get_token(theme, "effect.card_radius", True))

        # 序号徽章（左上角）
        badge_size = 0.5
        badge_rect = GridRect(
            left=cell.left + 0.15, top=cell.top + 0.15,
            width=badge_size, height=badge_size,
        )
        render_number_badge(
            slide, badge_rect, f"{idx + 1:02d}", "number", ctx, theme,
        )

        # 卡片文本
        text_rect = GridRect(
            left=cell.left + 0.15 + badge_size + 0.15, top=cell.top + 0.15,
            width=cell.width - badge_size - 0.45, height=badge_size,
        )
        add_text_box(
            slide, text_rect, str(item), "body", ctx, theme,
            font_size_pt=get_token(theme, "font.body.size_pt", 16),
            bold=False,
            color=resolve_color(theme, get_token(theme, "color.text_primary", "#1F2937")),
            align="left", anchor="middle",
        )


# ==================== kpi KPI 页：等宽横向卡片矩阵 ====================
@register_layout("kpi")
def layout_kpi(slide: Slide, page_data: dict, theme: dict, ctx: LayoutContext) -> None:
    """KPI 页布局：等宽横向卡片矩阵

    版式：
        - 顶部标题区（占 20%）
        - 下方 KPI 卡片矩阵（1 行 N 列，N 由 kpi_items 数量决定，最多 4 列）
    字段：
        - title: 页面标题
        - kpi_items[{label, value, trend}]: KPI 数据
    元素角色：
        - title: 页面标题
        - card: 每个 KPI 卡片
        - number: KPI 数值
        - desc: KPI 标签
    """
    area = safe_area(get_token(theme, "spacing.safe_margin_inch", 0.5))

    # 顶部标题区
    title_rect = GridRect(area.left, area.top, area.width, 1.0)
    add_text_box(
        slide, title_rect, page_data.get("title", ""), "title", ctx, theme,
        font_size_pt=get_token(theme, "font.title.size_pt", 36),
        bold=True,
        color=resolve_color(theme, get_token(theme, "color.text_primary", "#1F2937")),
        align="left", anchor="middle",
    )

    # 装饰下划线
    add_line(
        slide, area.left, area.top + 1.1,
        area.left + 1.5, area.top + 1.1,
        "divider", ctx, theme,
        color=resolve_color(theme, get_token(theme, "color.primary", "#1A56DB")),
        width_pt=2.0,
    )

    # KPI 卡片矩阵
    kpi_items = page_data.get("kpi_items", [])
    if not kpi_items:
        return

    cards_area = GridRect(
        left=area.left, top=area.top + 1.5,
        width=area.width, height=area.height - 1.5,
    )
    # KPI 卡片固定 1 行，列数 = KPI 数量（最多 4）
    cols = min(len(kpi_items), 4)
    matrix = grid_matrix(cards_area, rows=1, cols=cols,
                         gutter=get_token(theme, "spacing.grid_gutter_inch", 0.25))

    for idx, item in enumerate(kpi_items):
        if idx >= cols:
            break
        cell = matrix[0][idx]

        # KPI 卡片背景
        add_card(slide, cell, "card", ctx, theme,
                 radius=get_token(theme, "effect.card_radius", True))

        # KPI 数值（大号居中，上半部分）
        value_text = str(item.get("value", ""))
        value_rect = GridRect(
            left=cell.left, top=cell.top + cell.height * 0.2,
            width=cell.width, height=cell.height * 0.5,
        )
        add_text_box(
            slide, value_rect, value_text, "kpi_number", ctx, theme,
            font_size_pt=get_token(theme, "font.kpi_number.size_pt", 48),
            bold=get_token(theme, "font.kpi_number.bold", True),
            color=resolve_color(theme, get_token(theme, "color.kpi_number", "#1A56DB")),
            align="center", anchor="middle",
        )

        # KPI 标签（小号居中，下半部分）
        label_text = str(item.get("label", ""))
        label_rect = GridRect(
            left=cell.left, top=cell.top + cell.height * 0.7,
            width=cell.width, height=cell.height * 0.25,
        )
        add_text_box(
            slide, label_rect, label_text, "desc", ctx, theme,
            font_size_pt=get_token(theme, "font.kpi_label.size_pt", 14),
            bold=False,
            color=resolve_color(theme, get_token(theme, "color.text_secondary", "#6B7280")),
            align="center", anchor="middle",
        )
