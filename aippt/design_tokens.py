"""
guizang-ppt-skill 设计令牌直接抽取（T501 核心模块）

设计理念：
    guizang-ppt-skill 项目的设计令牌已明确文档化（CSS 变量形式），
    无需用 CV 分析图片，直接 1:1 翻译为 Python 常量更准确、更高效。
    本模块提供"令牌直接抽取"作为主要方案，CV 解析（design_parser.py）作为辅助。

来源（仅参考规范文档，不复制代码，避免 AGPL 传染）：
    - references/themes-swiss.md：Style B 4 套锚点色
    - references/themes.md：Style A 5 套墨水/纸张配色
    - references/components.md：字体层级、间距令牌、栅格规范
    - references/layouts-swiss.md：22 个锁定版式

许可边界：
    本模块仅参考"设计规范思想与数值"（hex 值、字号、间距），
    不复制 HTML/CSS 代码，避免 AGPL-3.0 传染。
"""
from __future__ import annotations

from typing import Any


# ==================== Style B 瑞士风 4 套锚点色 ====================
# 每套含 accent（主色）+ accent_on（前景色）+ accent_rgb（RGB 元组）
# 来源：references/themes-swiss.md
SWISS_THEMES: dict[str, dict[str, Any]] = {
    "克莱因蓝": {
        "name": "克莱因蓝 IKB",
        "name_en": "klein_blue",
        "accent": "#002FA7",
        "accent_rgb": (0, 47, 167),
        "accent_on": "#FFFFFF",
        "description": "通用/AI/科技场景，瑞士风首选锚点色",
    },
    "柠檬黄": {
        "name": "柠檬黄",
        "name_en": "lemon_yellow",
        "accent": "#FFD500",
        "accent_rgb": (255, 213, 0),
        "accent_on": "#0A0A0A",
        "description": "年轻/消费场景，高对比活力色",
    },
    "柠檬绿": {
        "name": "柠檬绿",
        "name_en": "lemon_green",
        "accent": "#C5E803",
        "accent_rgb": (197, 232, 3),
        "accent_on": "#0A0A0A",
        "description": "生态/Z 世代场景，鲜亮生机色",
    },
    "安全橙": {
        "name": "安全橙",
        "name_en": "safety_orange",
        "accent": "#FF6B35",
        "accent_rgb": (255, 107, 53),
        "accent_on": "#FFFFFF",
        "description": "工业/警示场景，强视觉冲击色",
    },
}


# ==================== Style B 瑞士风灰阶基座（跨主题统一）====================
# 来源：references/themes-swiss.md
SWISS_GRAY_SCALE: dict[str, str] = {
    "paper": "#FAFAF8",          # 纸张背景色
    "paper_rgb": (250, 250, 248),
    "ink": "#0A0A0A",            # 主文字色
    "ink_rgb": (10, 10, 10),
    "grey_1": "#F0F0EE",         # 浅灰背景
    "grey_2": "#D4D4D2",         # 中灰边框
    "grey_3": "#737373",         # 次要文字
    "border_subtle": "#E0E0E0",  # 细边框
}


# ==================== Style A 电子杂志 5 套配色 ====================
# 每套含 ink（墨水色）+ paper（纸张色）
# 来源：references/themes.md
MAGAZINE_THEMES: dict[str, dict[str, Any]] = {
    "墨水经典": {
        "name": "墨水经典",
        "name_en": "ink_classic",
        "ink": "#0A0A0B",
        "ink_rgb": (10, 10, 11),
        "paper": "#F1EFEA",
        "paper_rgb": (241, 239, 234),
        "description": "黑白经典，杂志感最强",
    },
    "靛蓝瓷": {
        "name": "靛蓝瓷",
        "name_en": "indigo_porcelain",
        "ink": "#0A1F3D",
        "ink_rgb": (10, 31, 61),
        "paper": "#F1F3F5",
        "paper_rgb": (241, 243, 245),
        "description": "深蓝 + 浅灰，科技商务感",
    },
    "森林墨": {
        "name": "森林墨",
        "name_en": "forest_ink",
        "ink": "#1A2E1F",
        "ink_rgb": (26, 46, 31),
        "paper": "#F5F1E8",
        "paper_rgb": (245, 241, 232),
        "description": "墨绿 + 米黄，自然沉稳",
    },
    "牛皮纸": {
        "name": "牛皮纸",
        "name_en": "kraft_paper",
        "ink": "#2A1E13",
        "ink_rgb": (42, 30, 19),
        "paper": "#EEDFC7",
        "paper_rgb": (238, 223, 199),
        "description": "深棕 + 牛皮纸黄，复古质感",
    },
    "沙丘": {
        "name": "沙丘",
        "name_en": "dune",
        "ink": "#1F1A14",
        "ink_rgb": (31, 26, 20),
        "paper": "#F0E6D2",
        "paper_rgb": (240, 230, 210),
        "description": "墨褐 + 沙黄，沙漠暖调",
    },
}


# ==================== 间距令牌（px → EMU 转换）====================
# 来源：references/components.md（--sp-3 到 --sp-13）
# 1px = 9525 EMU
SPACING_TOKENS_PX: dict[str, int] = {
    "sp_3": 8,     # 最小间距
    "sp_4": 12,
    "sp_5": 16,    # 基础间距（栅格 gap）
    "sp_6": 24,
    "sp_7": 32,
    "sp_8": 40,
    "sp_9": 48,
    "sp_10": 64,
    "sp_11": 80,
    "sp_12": 96,
    "sp_13": 160,  # 最大间距
}

#: 1px = 9525 EMU（python-pptx 内部单位）
PX_PER_EMU: int = 9525


def px_to_emu(px: int) -> int:
    """像素转 EMU

    :param px: 像素值
    :return: EMU 值
    """
    return px * PX_PER_EMU


def px_to_inch(px: int, dpi: int = 96) -> float:
    """像素转英寸（默认 96 DPI，对应 Web 标准 DPI）

    :param px: 像素值
    :param dpi: DPI，默认 96
    :return: 英寸值
    """
    return px / dpi


def spacing_to_inch(token: str) -> float:
    """间距令牌转英寸

    :param token: 间距令牌名（如 "sp_5"）
    :return: 英寸值
    :raises KeyError: 未知令牌
    """
    px = SPACING_TOKENS_PX[token]
    return px_to_inch(px)


# ==================== 字体层级规范（Style B 瑞士风）====================
# 来源：references/components.md
# "越大越细，越小越粗"反直觉规则
FONT_LADDER_SWISS: dict[str, dict[str, Any]] = {
    "display": {
        "size_vw": 8.0,        # 封面大字
        "weight": 200,         # ExtraLight
        "weight_name": "ExtraLight",
        "use_for": "封面大字、h-statement",
    },
    "h1": {
        "size_vw": 6.0,        # 章节标题
        "weight": 300,         # Light
        "weight_name": "Light",
        "use_for": "章节标题、页面主标题",
    },
    "h2": {
        "size_vw": 4.0,        # 中型标题
        "weight": 300,         # Light
        "weight_name": "Light",
        "use_for": "二级标题",
    },
    "h3": {
        "size_vw": 2.5,        # 小标题
        "weight": 400,         # Regular
        "weight_name": "Regular",
        "use_for": "卡片标题、小节标题",
    },
    "body": {
        "size_vw": 1.3,        # 正文
        "weight": 400,         # Regular
        "weight_name": "Regular",
        "use_for": "正文内容",
    },
    "meta": {
        "size_vw": 1.0,        # 元信息
        "weight": 500,         # Medium
        "weight_name": "Medium",
        "use_for": "kicker、meta 信息",
    },
    "kicker": {
        "size_vw": 0.9,        # 最小标签
        "weight": 600,         # SemiBold
        "weight_name": "SemiBold",
        "use_for": "uppercase 标签",
    },
}

#: 演示最小字号下限（投屏不可读下限）
MIN_FONT_PX: dict[str, int] = {
    "body": 18,        # 正文最小 18px
    "card_desc": 16,   # 卡片描述最小 16px
    "meta": 14,        # meta 最小 14px
}


# ==================== 字体族规范 ====================
FONT_FAMILY_SWISS: dict[str, str] = {
    "sans": "Inter, Helvetica Neue, Noto Sans SC, Microsoft YaHei",
    "sans_fallback": "Microsoft YaHei",  # Windows 中文回退
    "mono": "JetBrains Mono, Consolas, monospace",
}

FONT_FAMILY_MAGAZINE: dict[str, str] = {
    "serif": "Playfair Display, Noto Serif SC, SimSun",
    "serif_fallback": "SimSun",  # Windows 中文衬线回退
    "sans": "Noto Sans SC, Microsoft YaHei",
    "mono": "IBM Plex Mono, Consolas, monospace",
}


# ==================== 栅格规范 ====================
# Style B：16 列 grid
SWISS_GRID: dict[str, Any] = {
    "columns": 16,
    "gap_px": 16,
    "gap_inch": px_to_inch(16),
    "canvas_padding_vh": 5.6,   # 上下内边距
    "canvas_padding_vw": 5.0,   # 左右内边距
    "nav_safe_bottom_vh": 8.0,  # 底部导航安全区
}

# Style A：非对称杂志网格
MAGAZINE_GRIDS: dict[str, list[int]] = {
    "grid_2_7_5": [2, 7, 5],   # 左标签 + 中主内容 + 右辅助
    "grid_2_6_6": [2, 6, 6],   # 左标签 + 双等宽
    "grid_2_8_4": [2, 8, 4],   # 左标签 + 大主内容 + 小辅助
    "grid_3_3": [3, 3],        # 双等宽
    "grid_6": [6],             # 单列 6 等分
    "grid_4": [4],             # 单列 4 等分
}


# ==================== 单位换算工具（vw/vh → pt/inch）====================
def vw_to_pt(vw: float, slide_width_inch: float = 13.333,
             dpi: int = 96) -> float:
    """vw（视口宽度百分比）转 pt

    python-pptx 用 pt 作为字号单位，1pt = 1/72 inch。
    vw 是相对视口宽度的百分比，1vw = 视口宽度 × 1%。

    :param vw: vw 值（如 8.0 表示 8vw）
    :param slide_width_inch: slide 宽度（英寸），默认 13.333（16:9）
    :param dpi: DPI，默认 96
    :return: pt 值
    """
    width_px = slide_width_inch * dpi
    font_px = width_px * vw / 100
    font_inch = font_px / dpi
    return font_inch * 72  # inch → pt


def vh_to_pt(vh: float, slide_height_inch: float = 7.5,
             dpi: int = 96) -> float:
    """vh（视口高度百分比）转 pt

    :param vh: vh 值
    :param slide_height_inch: slide 高度（英寸），默认 7.5（16:9）
    :param dpi: DPI，默认 96
    :return: pt 值
    """
    height_px = slide_height_inch * dpi
    font_px = height_px * vh / 100
    font_inch = font_px / dpi
    return font_inch * 72


def min_vw_vh_to_pt(vw: float, vh: float,
                    slide_width_inch: float = 13.333,
                    slide_height_inch: float = 7.5) -> float:
    """处理 min(Xvw, Yvh) 双约束，取较小值转 pt

    guizang-ppt-skill 的字号常用 min(Xvw, Yvh) 双约束，
    避免在 16:9 屏高度截断。

    :param vw: vw 值
    :param vh: vh 值
    :param slide_width_inch: slide 宽度（英寸）
    :param slide_height_inch: slide 高度（英寸）
    :return: pt 值（取较小者）
    """
    return min(vw_to_pt(vw, slide_width_inch), vh_to_pt(vh, slide_height_inch))


def font_ladder_to_pt(slide_width_inch: float = 13.333) -> dict[str, float]:
    """将瑞士风字体阶梯转为 pt 值（基于 16:9 标准画布）

    :param slide_width_inch: slide 宽度（英寸）
    :return: {角色: pt 值} 字典
    """
    result: dict[str, float] = {}
    for role, spec in FONT_LADDER_SWISS.items():
        pt = vw_to_pt(spec["size_vw"], slide_width_inch)
        # 应用最小字号下限保护
        if role in MIN_FONT_PX:
            min_pt = MIN_FONT_PX[role] / 96 * 72  # px → pt
            pt = max(pt, min_pt)
        result[role] = round(pt, 1)
    return result
