"""
主题生成器（T502）

功能：
    将 design_parser.py 的解析结果（或 design_tokens.py 的明文抽取结果）
    转换为标准主题 JSON 配置（与 themes/商务蓝.json schema 一致），
    支持手动微调参数与提取结果优化。

核心 API：
    generate_theme(parser_output, theme_name, ...) → 标准主题 dict
    generate_theme_from_image(image_path, theme_name, ...) → 一步到位
    generate_theme_from_preset(preset_name, theme_name, ...) → 从 guizang 预设生成
    save_theme(theme, output_path) → 写入 themes/ 目录

设计约束：
    - 输出 schema 与 themes/商务蓝.json 完全一致，可直接被 theme_loader.load_theme 加载
    - 所有颜色推导使用色彩学公式（HSL/亮度公式），不依赖外部库
    - overrides 参数支持任意字段覆盖，优先级最高
    - 缺失字段使用兜底默认值，保证输出完整性
"""
from __future__ import annotations

import colorsys
import json
import re
from pathlib import Path
from typing import Any, Optional

from aippt.logger import logger


# ==================== 兜底默认值 ====================
#: 当解析器未提供时的兜底主色（中性蓝灰，避免黑白单调）
_FALLBACK_PRIMARY = "#1A56DB"
#: 兜底背景色
_FALLBACK_BG = "#FFFFFF"
#: 兜底文字色
_FALLBACK_TEXT = "#1F2937"
#: 兜底次要文字色
_FALLBACK_TEXT_SECONDARY = "#6B7280"


# ==================== 颜色工具 ====================
def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """hex → RGB 元组

    :param hex_color: hex 字符串（如 "#1A56DB" 或 "1A56DB"）
    :return: (r, g, b) 0-255
    """
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    """RGB → hex 字符串

    :param rgb: (r, g, b) 0-255
    :return: hex 字符串（如 "#1A56DB"）
    """
    return "#{:02X}{:02X}{:02X}".format(*rgb)


def _clamp(value: int, low: int = 0, high: int = 255) -> int:
    """限制到 [low, high] 区间"""
    return max(low, min(high, value))


def _adjust_lightness(hex_color: str, delta_l: float) -> str:
    """调整颜色亮度（HSL 空间）

    :param hex_color: 输入 hex
    :param delta_l: 亮度增量（-1.0 到 1.0，负值变暗，正值变亮）
    :return: 调整后的 hex
    """
    r, g, b = _hex_to_rgb(hex_color)
    # RGB 0-255 → HLS 0-1
    h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    new_l = max(0.0, min(1.0, l + delta_l))
    nr, ng, nb = colorsys.hls_to_rgb(h, new_l, s)
    return _rgb_to_hex((_clamp(int(nr * 255)), _clamp(int(ng * 255)), _clamp(int(nb * 255))))


def darken(hex_color: str, amount: float = 0.15) -> str:
    """变暗颜色（默认 -15% 亮度）"""
    return _adjust_lightness(hex_color, -abs(amount))


def lighten(hex_color: str, amount: float = 0.20) -> str:
    """变亮颜色（默认 +20% 亮度）"""
    return _adjust_lightness(hex_color, abs(amount))


def relative_luminance(hex_color: str) -> float:
    """计算相对亮度（WCAG 2.1 公式）

    用于判断文字在背景上的可读性。
    返回 0-1，0 表示最暗（黑），1 表示最亮（白）。

    :param hex_color: hex 颜色
    :return: 亮度值 0-1
    """
    r, g, b = _hex_to_rgb(hex_color)

    def _channel(c: int) -> float:
        cs = c / 255
        return cs / 12.92 if cs <= 0.03928 else ((cs + 0.055) / 1.055) ** 2.4

    return 0.2126 * _channel(r) + 0.7152 * _channel(g) + 0.0722 * _channel(b)


def contrast_ratio(hex1: str, hex2: str) -> float:
    """计算两色对比度（WCAG 2.1）

    :return: 对比度 1-21，4.5+ 满足正文可读，3+ 满足大字可读
    """
    l1 = relative_luminance(hex1)
    l2 = relative_luminance(hex2)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def pick_readable_fg(bg_hex: str, dark_choice: str = "#0A0A0A",
                     light_choice: str = "#FFFFFF") -> str:
    """为背景选择可读的前景色（黑或白）

    :param bg_hex: 背景色
    :param dark_choice: 深色选项
    :param light_choice: 浅色选项
    :return: 对比度更高的前景色
    """
    dark_cr = contrast_ratio(bg_hex, dark_choice)
    light_cr = contrast_ratio(bg_hex, light_choice)
    return dark_choice if dark_cr >= light_cr else light_choice


def is_hex_color(s: Any) -> bool:
    """判断字符串是否为合法 hex 颜色"""
    if not isinstance(s, str):
        return False
    return bool(re.match(r"^#?[0-9A-Fa-f]{6}$", s))


# ==================== 颜色块构建 ====================
def _build_color_block(
    parser_output: dict[str, Any],
    overrides: Optional[dict[str, str]] = None,
) -> dict[str, str]:
    """从解析结果构建标准 color 块

    推导链：
        primary / secondary / background / text_primary / text_secondary → 解析器输出
        primary_dark / primary_light → 由 primary 亮度推导
        text_on_primary → 由 primary 亮度自动选择黑/白
        bg_section / card_bg / card_border / divider → 由 background 推导
        number_bg / number_fg → primary + 前景自动选择
        kpi_bg / kpi_border / kpi_number → card_bg / card_border / primary
        cover_bg / cover_title / cover_subtitle → primary_dark + 前景
        divider_bg / divider_number / divider_title → primary_dark + 前景

    :param parser_output: design_parser.parse_design_image 返回值
    :param overrides: 手动覆盖的 color 字段 {key: hex}
    :return: 标准 color 块
    """
    overrides = overrides or {}
    colors_in = parser_output.get("colors", {}) or {}

    # 基础五色（解析器输出 → 兜底）
    primary = overrides.get("primary") or colors_in.get("primary") or _FALLBACK_PRIMARY
    secondary = overrides.get("secondary") or colors_in.get("secondary") or lighten(primary, 0.15)
    background = overrides.get("bg_page") or colors_in.get("background") or _FALLBACK_BG
    text_primary = overrides.get("text_primary") or colors_in.get("text_primary") or _FALLBACK_TEXT
    text_secondary = (overrides.get("text_secondary")
                      or colors_in.get("text_secondary") or _FALLBACK_TEXT_SECONDARY)

    # 规范化 hex（确保以 # 开头）
    def _norm(c: str) -> str:
        return c if c.startswith("#") else f"#{c}"

    primary = _norm(primary)
    secondary = _norm(secondary)
    background = _norm(background)
    text_primary = _norm(text_primary)
    text_secondary = _norm(text_secondary)

    # 推导变体
    primary_dark = overrides.get("primary_dark") or darken(primary, 0.15)
    primary_light = overrides.get("primary_light") or lighten(primary, 0.20)
    accent = overrides.get("accent") or secondary

    # 前景色选择
    text_on_primary = overrides.get("text_on_primary") or pick_readable_fg(primary)
    text_on_dark = overrides.get("text_on_dark") or pick_readable_fg(primary_dark)

    # 背景系列
    bg_page = background
    bg_section = overrides.get("bg_section") or _adjust_lightness(background, -0.03)
    card_bg = overrides.get("card_bg") or background
    card_border = overrides.get("card_border") or _adjust_lightness(background, -0.10)
    divider = overrides.get("divider") or card_border

    # number 元素
    number_bg = overrides.get("number_bg") or primary
    number_fg = overrides.get("number_fg") or pick_readable_fg(number_bg)

    # KPI 卡片
    kpi_bg = overrides.get("kpi_bg") or card_bg
    kpi_border = overrides.get("kpi_border") or card_border
    kpi_number = overrides.get("kpi_number") or primary

    # 封面
    cover_bg = overrides.get("cover_bg") or primary_dark
    cover_title = overrides.get("cover_title") or pick_readable_fg(cover_bg)
    cover_subtitle = overrides.get("cover_subtitle") or lighten(primary_light, 0.10)

    # 章节分隔页
    divider_bg = overrides.get("divider_bg") or primary_dark
    divider_number = overrides.get("divider_number") or pick_readable_fg(divider_bg)
    divider_title = overrides.get("divider_title") or pick_readable_fg(divider_bg)

    return {
        "primary": primary,
        "primary_dark": primary_dark,
        "primary_light": primary_light,
        "secondary": secondary,
        "accent": accent,
        "text_primary": text_primary,
        "text_secondary": text_secondary,
        "text_on_primary": text_on_primary,
        "text_on_dark": text_on_dark,
        "bg_page": bg_page,
        "bg_section": bg_section,
        "card_bg": card_bg,
        "card_border": card_border,
        "divider": divider,
        "number_bg": number_bg,
        "number_fg": number_fg,
        "kpi_bg": kpi_bg,
        "kpi_border": kpi_border,
        "kpi_number": kpi_number,
        "cover_bg": cover_bg,
        "cover_title": cover_title,
        "cover_subtitle": cover_subtitle,
        "divider_bg": divider_bg,
        "divider_number": divider_number,
        "divider_title": divider_title,
    }


# ==================== 字体块构建 ====================
#: 字体角色 → 解析器字号层级 role_guess 映射
_ROLE_TO_THEME_ROLE: dict[str, str] = {
    "display": "cover_title",
    "h1": "divider_title",
    "h2": "title",
    "h3": "subtitle",
    "body": "body",
    "meta": "desc",
    "kicker": "kpi_label",
}


def _pick_font_size_by_role(
    levels: list[dict[str, Any]],
    role: str,
    default_pt: float,
) -> float:
    """从解析器字号层级中按角色挑选 pt

    :param levels: 解析器返回的 fonts.levels
    :param role: 字体角色（display/h1/h2/h3/body/meta/kicker）
    :param default_pt: 找不到时的兜底字号
    :return: 字号 pt
    """
    for lv in levels:
        if lv.get("role_guess") == role:
            return float(lv.get("height_pt", default_pt))
    return default_pt


def _build_font_block(
    parser_output: dict[str, Any],
    style: str = "auto",
    overrides: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """从解析结果构建标准 font 块

    字号来源优先级：
        1. overrides 中显式指定的 size_pt
        2. 解析器 fonts.levels 中对应 role_guess
        3. guizang 瑞士风字体阶梯（design_tokens.FONT_LADDER_SWISS）转 pt
        4. 兜底默认值

    :param parser_output: 解析器输出
    :param style: 风格预设（auto/swiss/magazine）
    :param overrides: 手动覆盖字段
    :return: 标准 font 块
    """
    overrides = overrides or {}
    fonts_in = parser_output.get("fonts", {}) or {}
    levels = fonts_in.get("levels", []) if isinstance(fonts_in, dict) else []

    # 字体族选择
    if style == "magazine":
        family = overrides.get("family", "Noto Serif SC, SimSun, Microsoft YaHei")
        family_en = overrides.get("family_en", "Playfair Display, Calibri")
    else:
        # swiss / auto 默认无衬线
        family = overrides.get("family", "Microsoft YaHei")
        family_en = overrides.get("family_en", "Calibri")

    # 兜底字号（瑞士风阶梯转 pt 后的值，13.333 英寸宽画布）
    # 这些值与 themes/商务蓝.json 大致一致，保证可用性
    _DEFAULTS = {
        "title": 36, "subtitle": 20, "body": 16, "desc": 14,
        "number": 28, "year": 24,
        "kpi_number": 48, "kpi_label": 14,
        "cover_title": 44, "cover_subtitle": 22,
        "divider_number": 72, "divider_title": 40,
    }

    def _size(role: str, parser_role: str) -> int:
        """优先 overrides → 解析器 → 默认"""
        ov = overrides.get(role, {})
        if isinstance(ov, dict) and "size_pt" in ov:
            return int(ov["size_pt"])
        pt = _pick_font_size_by_role(levels, parser_role, _DEFAULTS.get(role, 16))
        return int(pt)

    def _bold(role: str, default_bold: bool) -> bool:
        ov = overrides.get(role, {})
        if isinstance(ov, dict) and "bold" in ov:
            return bool(ov["bold"])
        return default_bold

    def _color(role: str, default_ref: str) -> str:
        """优先 overrides.color → 默认引用"""
        ov = overrides.get(role, {})
        if isinstance(ov, dict) and "color" in ov:
            return ov["color"]
        return default_ref

    return {
        "family": family,
        "family_en": family_en,
        "title": {"size_pt": _size("title", "h2"), "bold": _bold("title", True), "color": _color("title", "text_primary")},
        "subtitle": {"size_pt": _size("subtitle", "h3"), "bold": _bold("subtitle", False), "color": _color("subtitle", "text_secondary")},
        "body": {"size_pt": _size("body", "body"), "bold": _bold("body", False), "color": _color("body", "text_primary")},
        "desc": {"size_pt": _size("desc", "meta"), "bold": _bold("desc", False), "color": _color("desc", "text_secondary")},
        "number": {"size_pt": _size("number", "h2"), "bold": _bold("number", True), "color": _color("number", "number_fg")},
        "year": {"size_pt": _size("year", "h2"), "bold": _bold("year", True), "color": _color("year", "primary")},
        "kpi_number": {"size_pt": _size("kpi_number", "display"), "bold": _bold("kpi_number", True), "color": _color("kpi_number", "kpi_number")},
        "kpi_label": {"size_pt": _size("kpi_label", "kicker"), "bold": _bold("kpi_label", False), "color": _color("kpi_label", "text_secondary")},
        "cover_title": {"size_pt": _size("cover_title", "display"), "bold": _bold("cover_title", True), "color": _color("cover_title", "cover_title")},
        "cover_subtitle": {"size_pt": _size("cover_subtitle", "body"), "bold": _bold("cover_subtitle", False), "color": _color("cover_subtitle", "cover_subtitle")},
        "divider_number": {"size_pt": _size("divider_number", "display"), "bold": _bold("divider_number", True), "color": _color("divider_number", "divider_number")},
        "divider_title": {"size_pt": _size("divider_title", "h1"), "bold": _bold("divider_title", True), "color": _color("divider_title", "divider_title")},
    }


# ==================== 间距块构建 ====================
def _build_spacing_block(
    parser_output: dict[str, Any],
    overrides: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """从解析结果构建标准 spacing 块

    推导：
        safe_margin_inch → 解析器 spacing.safe_margin_inch（兜底 0.5）
        grid_gutter_inch → dominant_gap_inch（兜底 0.25）
        card_padding_inch → grid_gutter * 0.8
        section_gap_inch → safe_margin * 0.8
        element_gap_inch → grid_gutter * 0.6

    :param parser_output: 解析器输出
    :param overrides: 手动覆盖
    :return: 标准 spacing 块
    """
    overrides = overrides or {}
    spacing_in = parser_output.get("spacing", {}) or {}

    safe_margin = overrides.get("safe_margin_inch") or spacing_in.get("safe_margin_inch") or 0.5
    safe_margin = float(safe_margin)

    dominant_gap = overrides.get("grid_gutter_inch") or spacing_in.get("dominant_gap_inch") or 0.25
    dominant_gap = float(dominant_gap)

    return {
        "safe_margin_inch": round(safe_margin, 3),
        "grid_gutter_inch": round(dominant_gap, 3),
        "card_padding_inch": round(overrides.get("card_padding_inch", dominant_gap * 0.8), 3),
        "section_gap_inch": round(overrides.get("section_gap_inch", safe_margin * 0.8), 3),
        "element_gap_inch": round(overrides.get("element_gap_inch", dominant_gap * 0.6), 3),
    }


# ==================== 效果与布局块（无解析器来源，使用默认 + overrides）====================
def _build_effect_block(
    overrides: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """构建 effect 块（默认无阴影、圆角）"""
    overrides = overrides or {}
    return {
        "card_shadow": overrides.get("card_shadow", False),
        "card_radius": overrides.get("card_radius", True),
        "card_radius_inch": overrides.get("card_radius_inch", 0.08),
    }


def _build_layout_block(
    overrides: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """构建 layout 块（默认 4 KPI / 2 列表卡片）"""
    overrides = overrides or {}
    return {
        "cover_split_ratio": overrides.get("cover_split_ratio", 0.45),
        "divider_number_ratio": overrides.get("divider_number_ratio", 0.35),
        "kpi_card_per_row": overrides.get("kpi_card_per_row", 4),
        "numbered_list_card_per_row": overrides.get("numbered_list_card_per_row", 2),
    }


# ==================== 主题生成主入口 ====================
def generate_theme(
    parser_output: dict[str, Any],
    theme_name: str,
    name_en: Optional[str] = None,
    description: str = "",
    style: str = "auto",
    overrides: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """从解析器输出生成标准主题 JSON

    :param parser_output: design_parser.parse_design_image 返回的字典
    :param theme_name: 主题名（如 "guizang-商务科技风"）
    :param name_en: 英文名，缺省时从 theme_name 拼音化或留空
    :param description: 主题描述
    :param style: 风格预设（auto/swiss/magazine），影响字体族与字号推导
    :param overrides: 手动微调字段，支持结构：
        {
          "color": {"primary": "#002FA7", "text_primary": "#0A0A0A", ...},
          "font": {"title": {"size_pt": 40, "bold": true}, "family": "Inter", ...},
          "spacing": {"safe_margin_inch": 0.6, ...},
          "effect": {"card_shadow": true, ...},
          "layout": {"kpi_card_per_row": 3, ...}
        }
    :return: 标准主题字典，schema 与 themes/商务蓝.json 一致
    """
    overrides = overrides or {}
    color_overrides = overrides.get("color", {})
    font_overrides = overrides.get("font", {})
    spacing_overrides = overrides.get("spacing", {})
    effect_overrides = overrides.get("effect", {})
    layout_overrides = overrides.get("layout", {})

    # 风格自动检测：解析器无明确信号时默认 swiss
    if style == "auto":
        style = "swiss"

    color_block = _build_color_block(parser_output, color_overrides)
    font_block = _build_font_block(parser_output, style, font_overrides)
    spacing_block = _build_spacing_block(parser_output, spacing_overrides)
    effect_block = _build_effect_block(effect_overrides)
    layout_block = _build_layout_block(layout_overrides)

    theme = {
        "name": theme_name,
        "name_en": name_en or "",
        "description": description or f"由设计稿解析自动生成（风格: {style}）",
        "color": color_block,
        "font": font_block,
        "spacing": spacing_block,
        "effect": effect_block,
        "layout": layout_block,
    }

    # 一致性校验：font.color 引用必须在 color 块中存在
    _validate_color_refs(theme)

    # 可读性校验：警告低对比度组合
    _warn_low_contrast(theme)

    return theme


def _validate_color_refs(theme: dict[str, Any]) -> None:
    """校验 font.*.color 引用是否在 color 块中存在

    :raises ValueError: 引用无法解析
    """
    color_block = theme.get("color", {})
    font_block = theme.get("font", {})
    for role, spec in font_block.items():
        if role in ("family", "family_en"):
            continue
        if not isinstance(spec, dict):
            continue
        ref = spec.get("color")
        if ref and not ref.startswith("#") and ref not in color_block:
            raise ValueError(f"font.{role}.color 引用 '{ref}' 在 color 块中不存在")


def _warn_low_contrast(theme: dict[str, Any]) -> None:
    """警告低对比度组合（不抛异常，仅记录）

    检查项：
        - text_primary vs bg_page
        - text_on_primary vs primary
        - cover_title vs cover_bg
    """
    color_block = theme.get("color", {})
    font_block = theme.get("font", {})

    def _resolve(ref: str) -> str:
        if ref.startswith("#"):
            return ref
        return color_block.get(ref, "#000000")

    pairs = [
        ("text_primary vs bg_page", _resolve("text_primary"), color_block.get("bg_page", "#FFFFFF")),
        ("text_on_primary vs primary", color_block.get("text_on_primary", "#FFFFFF"), color_block.get("primary", "#000000")),
        ("cover_title vs cover_bg", color_block.get("cover_title", "#FFFFFF"), color_block.get("cover_bg", "#000000")),
    ]
    for name, fg, bg in pairs:
        try:
            cr = contrast_ratio(fg, bg)
            if cr < 3.0:
                logger.warning("对比度不足: %s = %.2f（fg=%s, bg=%s，建议 ≥3.0）",
                               name, cr, fg, bg)
        except Exception as e:
            logger.debug("对比度计算失败 %s: %s", name, e)


# ==================== 便捷入口 ====================
def generate_theme_from_image(
    image_path: str,
    theme_name: str,
    name_en: Optional[str] = None,
    description: str = "",
    style: str = "auto",
    overrides: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """一步到位：图片 → 解析 → 主题 JSON

    :param image_path: 设计稿图片路径
    :param theme_name: 主题名
    :param name_en: 英文名
    :param description: 描述
    :param style: 风格预设
    :param overrides: 手动微调
    :return: 标准主题字典
    """
    from aippt.design_parser import parse_design_image
    parser_output = parse_design_image(image_path)
    return generate_theme(
        parser_output, theme_name, name_en, description, style, overrides,
    )


def generate_theme_from_preset(
    preset_name: str,
    theme_name: Optional[str] = None,
    description: str = "",
    overrides: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """从 guizang 预设生成主题（无需图片，直接用 design_tokens 明文令牌）

    支持的预设：
        - 瑞士风-克莱因蓝 / 瑞士风-柠檬黄 / 瑞士风-柠檬绿 / 瑞士风-安全橙
        - 杂志风-墨水经典 / 杂志风-靛蓝瓷 / 杂志风-森林墨 / 杂志风-牛皮纸 / 杂志风-沙丘

    :param preset_name: 预设名
    :param theme_name: 输出主题名（缺省时用预设名）
    :param description: 描述
    :param overrides: 手动微调
    :return: 标准主题字典
    """
    from aippt import design_tokens

    # 解析预设名
    if preset_name.startswith("瑞士风-"):
        accent_key = preset_name[4:]
        if accent_key not in design_tokens.SWISS_THEMES:
            raise ValueError(f"未知瑞士风预设: {accent_key}，可选: {list(design_tokens.SWISS_THEMES.keys())}")
        theme_data = design_tokens.SWISS_THEMES[accent_key]
        gray = design_tokens.SWISS_GRAY_SCALE
        parser_output = {
            "colors": {
                "primary": theme_data["accent"],
                "secondary": gray["grey_3"],
                "background": gray["paper"],
                "text_primary": gray["ink"],
                "text_secondary": gray["grey_3"],
            },
            "fonts": {"levels": []},  # 由 _build_font_block 走瑞士风默认阶梯
            "spacing": {
                "safe_margin_inch": design_tokens.px_to_inch(design_tokens.SPACING_TOKENS_PX["sp_7"]),
                "dominant_gap_inch": design_tokens.px_to_inch(design_tokens.SPACING_TOKENS_PX["sp_5"]),
            },
        }
        style = "swiss"
        default_name = f"guizang-瑞士风-{accent_key}"
        default_desc = f"guizang-ppt-skill 瑞士风预设：{theme_data['name']}（{theme_data.get('description', '')}）"
    elif preset_name.startswith("杂志风-"):
        accent_key = preset_name[4:]
        if accent_key not in design_tokens.MAGAZINE_THEMES:
            raise ValueError(f"未知杂志风预设: {accent_key}，可选: {list(design_tokens.MAGAZINE_THEMES.keys())}")
        theme_data = design_tokens.MAGAZINE_THEMES[accent_key]
        parser_output = {
            "colors": {
                "primary": theme_data["ink"],
                "secondary": design_tokens.SWISS_GRAY_SCALE["grey_3"],
                "background": theme_data["paper"],
                "text_primary": theme_data["ink"],
                "text_secondary": design_tokens.SWISS_GRAY_SCALE["grey_3"],
            },
            "fonts": {"levels": []},
            "spacing": {
                "safe_margin_inch": design_tokens.px_to_inch(design_tokens.SPACING_TOKENS_PX["sp_7"]),
                "dominant_gap_inch": design_tokens.px_to_inch(design_tokens.SPACING_TOKENS_PX["sp_5"]),
            },
        }
        style = "magazine"
        default_name = f"guizang-杂志风-{accent_key}"
        default_desc = f"guizang-ppt-skill 杂志风预设：{theme_data['name']}（{theme_data.get('description', '')}）"
    else:
        raise ValueError(
            f"未知预设: {preset_name}，支持的预设：\n"
            f"  瑞士风: {[f'瑞士风-{k}' for k in design_tokens.SWISS_THEMES]}\n"
            f"  杂志风: {[f'杂志风-{k}' for k in design_tokens.MAGAZINE_THEMES]}"
        )

    return generate_theme(
        parser_output=parser_output,
        theme_name=theme_name or default_name,
        name_en=theme_data.get("name_en", ""),
        description=description or default_desc,
        style=style,
        overrides=overrides,
    )


def save_theme(theme: dict[str, Any], output_path: str) -> str:
    """保存主题到 JSON 文件

    :param theme: 主题字典
    :param output_path: 输出路径（如 "themes/guizang-瑞士风-克莱因蓝.json"）
    :return: 实际写入的绝对路径
    """
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(theme, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    logger.info("主题已保存: %s", path.resolve())
    return str(path.resolve())


def list_presets() -> list[str]:
    """列出所有可用的 guizang 预设名

    :return: 预设名列表
    """
    from aippt import design_tokens
    presets: list[str] = []
    presets.extend(f"瑞士风-{k}" for k in design_tokens.SWISS_THEMES)
    presets.extend(f"杂志风-{k}" for k in design_tokens.MAGAZINE_THEMES)
    return presets
