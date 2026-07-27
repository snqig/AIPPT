"""
aippt.layout 子包：自动布局引擎

模块说明：
    - ppt_auto_layout.py：栅格坐标计算、原子绘制组件、页面分发入口
    - page_layouts.py：T004 注册 5 类核心页面布局（cover/catalog/divider/numbered_list/kpi）

设计目标：
    1. 单位统一英寸，对齐 python-pptx 原生坐标体系
    2. 所有样式从主题 Design Token 读取，禁止硬编码
    3. 自动生成元素附加 role + shape_id，供动画模块匹配
    4. 渲染输出为原生可编辑 PPTX 元素，禁止图片嵌入
"""
from aippt.layout.ppt_auto_layout import (
    # 常量
    CANVAS_WIDTH_INCH, CANVAS_HEIGHT_INCH,
    GRID_COLUMNS, GRID_GUTTER_INCH, SAFE_MARGIN_INCH,
    # 数据结构
    GridRect, ElementMeta, LayoutContext,
    # 栅格工具
    safe_area, column_x, row_y,
    split_horizontal, split_vertical, grid_matrix,
    # 主题令牌工具
    get_token, hex_to_rgb,
    # 原子组件
    add_text_box, add_shape, add_line, add_card, render_number_badge,
    # 页面分发
    PAGE_LAYOUT_REGISTRY, register_layout, dispatch_page_layout,
    # Presentation 工具
    create_presentation, add_blank_slide,
)

# 导入 page_layouts 触发 5 类页面布局函数注册（cover/catalog/divider/numbered_list/kpi）
from aippt.layout import page_layouts  # noqa: F401

__all__ = [
    "CANVAS_WIDTH_INCH", "CANVAS_HEIGHT_INCH",
    "GRID_COLUMNS", "GRID_GUTTER_INCH", "SAFE_MARGIN_INCH",
    "GridRect", "ElementMeta", "LayoutContext",
    "safe_area", "column_x", "row_y",
    "split_horizontal", "split_vertical", "grid_matrix",
    "get_token", "hex_to_rgb",
    "add_text_box", "add_shape", "add_line", "add_card", "render_number_badge",
    "PAGE_LAYOUT_REGISTRY", "register_layout", "dispatch_page_layout",
    "create_presentation", "add_blank_slide",
]
