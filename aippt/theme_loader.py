"""
主题配置加载器（T003）

功能：
    1. 根据 theme_name 加载 themes/ 目录下的 JSON 主题配置
    2. 提供统一样式读取接口 get_style(theme, role, ...)
    3. 内置默认兜底主题，主题不存在自动降级
    4. 支持列出所有可用主题名

设计约束：
    - 主题不存在时降级到 DEFAULT_THEME，不抛异常（保证渲染不中断）
    - 所有样式读取走 get_token 点分路径，禁止硬编码
    - 主题切换不影响业务逻辑，仅影响视觉呈现
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Optional

from aippt.logger import logger


# ==================== 默认兜底主题 ====================
#: 主题不存在或加载失败时使用的兜底主题（保证渲染不中断）
DEFAULT_THEME: dict[str, Any] = {
    "name": "默认主题",
    "name_en": "default",
    "description": "兜底主题，主题加载失败时使用",
    "color": {
        "primary": "#1A56DB",
        "primary_dark": "#1E40AF",
        "primary_light": "#3B82F6",
        "secondary": "#64748B",
        "accent": "#F59E0B",
        "text_primary": "#1F2937",
        "text_secondary": "#6B7280",
        "text_on_primary": "#FFFFFF",
        "text_on_dark": "#F9FAFB",
        "bg_page": "#FFFFFF",
        "bg_section": "#F3F4F6",
        "card_bg": "#FFFFFF",
        "card_border": "#E5E7EB",
        "divider": "#E5E7EB",
        "number_bg": "#1A56DB",
        "number_fg": "#FFFFFF",
        "kpi_bg": "#FFFFFF",
        "kpi_border": "#E5E7EB",
        "kpi_number": "#1A56DB",
        "cover_bg": "#1A56DB",
        "cover_title": "#FFFFFF",
        "cover_subtitle": "#DBEAFE",
        "divider_bg": "#1E40AF",
        "divider_number": "#FFFFFF",
        "divider_title": "#FFFFFF",
    },
    "font": {
        "family": "Microsoft YaHei",
        "family_en": "Calibri",
        "title": {"size_pt": 36, "bold": True, "color": "text_primary"},
        "subtitle": {"size_pt": 20, "bold": False, "color": "text_secondary"},
        "body": {"size_pt": 16, "bold": False, "color": "text_primary"},
        "desc": {"size_pt": 14, "bold": False, "color": "text_secondary"},
        "number": {"size_pt": 28, "bold": True, "color": "number_fg"},
        "year": {"size_pt": 24, "bold": True, "color": "primary"},
        "kpi_number": {"size_pt": 48, "bold": True, "color": "kpi_number"},
        "kpi_label": {"size_pt": 14, "bold": False, "color": "text_secondary"},
        "cover_title": {"size_pt": 44, "bold": True, "color": "cover_title"},
        "cover_subtitle": {"size_pt": 22, "bold": False, "color": "cover_subtitle"},
        "divider_number": {"size_pt": 72, "bold": True, "color": "divider_number"},
        "divider_title": {"size_pt": 40, "bold": True, "color": "divider_title"},
    },
    "spacing": {
        "safe_margin_inch": 0.5,
        "grid_gutter_inch": 0.25,
        "card_padding_inch": 0.2,
        "section_gap_inch": 0.4,
        "element_gap_inch": 0.15,
    },
    "effect": {
        "card_shadow": False,
        "card_radius": True,
        "card_radius_inch": 0.08,
    },
    "layout": {
        "cover_split_ratio": 0.45,
        "divider_number_ratio": 0.35,
        "kpi_card_per_row": 4,
        "numbered_list_card_per_row": 2,
    },
}


#: 主题文件所在目录（默认项目根/themes）
#: __file__ 位于 Ppt_work/aippt/theme_loader.py，向上 parent 一次到 Ppt_work/
_THEMES_DIR: Path = Path(__file__).resolve().parent.parent / "themes"

#: 主题缓存 {theme_name: theme_dict}
_THEME_CACHE: dict[str, dict[str, Any]] = {}


def set_themes_dir(themes_dir: str | Path) -> None:
    """设置主题文件目录（支持自定义路径）

    :param themes_dir: 主题目录路径
    """
    global _THEMES_DIR
    _THEMES_DIR = Path(themes_dir)
    _THEME_CACHE.clear()  # 清空缓存，强制重新加载
    logger.debug("主题目录已切换: %s", _THEMES_DIR)


def get_themes_dir() -> Path:
    """获取当前主题目录

    :return: 主题目录 Path
    """
    return _THEMES_DIR


def list_themes() -> list[str]:
    """列出所有可用主题名

    扫描主题目录下所有 .json 文件，返回主题名列表。
    主题名优先使用 JSON 内的 name 字段，缺省时使用文件名 stem。

    :return: 主题名列表（如 ["商务蓝", "极简灰", "科技青"]）
    """
    if not _THEMES_DIR.exists():
        return []
    names: list[str] = []
    for f in sorted(_THEMES_DIR.glob("*.json")):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            name = data.get("name") or f.stem
            names.append(name)
        except Exception as e:
            logger.warning("主题文件 %s 解析失败: %s", f, e)
            names.append(f.stem)
    return names


def load_theme(theme_name: str) -> dict[str, Any]:
    """加载指定主题

    查找顺序：
        1. 缓存命中直接返回
        2. themes/ 目录下匹配 name 字段或文件名的 JSON
        3. 未找到时降级到 DEFAULT_THEME，记录 warning

    :param theme_name: 主题名（如 "商务蓝"）或文件名 stem
    :return: 主题字典（非空，最坏情况返回 DEFAULT_THEME）
    """
    if not theme_name:
        return DEFAULT_THEME

    # 缓存命中
    if theme_name in _THEME_CACHE:
        return _THEME_CACHE[theme_name]

    # 扫描主题目录
    if not _THEMES_DIR.exists():
        logger.warning("主题目录不存在: %s，使用默认主题", _THEMES_DIR)
        return DEFAULT_THEME

    for f in sorted(_THEMES_DIR.glob("*.json")):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            name = data.get("name")
            if name == theme_name or f.stem == theme_name:
                _THEME_CACHE[theme_name] = data
                logger.debug("主题 %s 已加载: %s", theme_name, f)
                return data
        except Exception as e:
            logger.warning("主题文件 %s 解析失败: %s", f, e)
            continue

    logger.warning("主题 %s 未找到，降级到默认主题", theme_name)
    return DEFAULT_THEME


def get_style(
    theme: dict[str, Any],
    role: str,
    attr: str,
    default: Any = None,
) -> Any:
    """统一样式读取接口

    根据 role + attr 从主题读取样式值，路径规则：
        - 字体：font.{role}.{attr}（如 font.title.size_pt）
        - 颜色：color.{role}_{attr}（如 color.title_bg）或 color.{attr}
        - 间距：spacing.{attr}
        - 布局：layout.{attr}
        - 效果：effect.{attr}

    本函数为高层封装，底层仍走 get_token 点分路径。

    :param theme: 主题字典
    :param role: 元素角色（title/subtitle/desc/number/year/card/...）
    :param attr: 样式属性（size_pt/bold/color/bg/border/...）
    :param default: 未找到时的默认值
    :return: 样式值或 default
    """
    if not theme:
        return default

    # 尝试多种路径，按优先级返回
    paths: list[str] = []
    if attr in ("size_pt", "bold"):
        paths.append(f"font.{role}.{attr}")
    elif attr in ("family", "family_en"):
        paths.append(f"font.{attr}")
    elif attr == "color":
        # 颜色：font.{role}.color 优先（可能引用 color.xxx），其次 color.{role}
        paths.append(f"font.{role}.color")
        paths.append(f"color.{role}")
    elif attr.endswith("_bg") or attr.endswith("_border") or attr.endswith("_fg"):
        paths.append(f"color.{attr}")
        paths.append(f"color.{role}_{attr.split('_')[-1]}")
    else:
        # 通用：spacing.{attr} / layout.{attr} / effect.{attr}
        paths.append(f"spacing.{attr}")
        paths.append(f"layout.{attr}")
        paths.append(f"effect.{attr}")
        paths.append(f"color.{attr}")

    from aippt.layout.ppt_auto_layout import get_token
    for path in paths:
        val = get_token(theme, path)
        if val is not None:
            # 颜色引用解析（如 font.title.color = "text_primary" → color.text_primary）
            if (attr == "color" or attr.endswith("_bg") or attr.endswith("_border")
                    or attr.endswith("_fg")) and isinstance(val, str) and not val.startswith("#"):
                resolved = get_token(theme, f"color.{val}")
                if resolved is not None:
                    return resolved
            return val
    return default


def resolve_color(theme: dict[str, Any], color_ref: str) -> str:
    """解析颜色引用（支持主题令牌引用与直接 hex）

    示例：
        resolve_color(theme, "text_primary") → "#1F2937"（从 color.text_primary 读取）
        resolve_color(theme, "#1A56DB")      → "#1A56DB"（直接返回）

    :param theme: 主题字典
    :param color_ref: 颜色引用（令牌名或 hex 字符串）
    :return: hex 颜色字符串
    :raises ValueError: 引用无法解析
    """
    if not color_ref:
        raise ValueError("color_ref 为空")
    if color_ref.startswith("#"):
        return color_ref
    from aippt.layout.ppt_auto_layout import get_token
    val = get_token(theme, f"color.{color_ref}")
    if val is None:
        raise ValueError(f"颜色引用无法解析: {color_ref}")
    if isinstance(val, str) and not val.startswith("#"):
        return resolve_color(theme, val)  # 多级引用
    return val


def clear_cache() -> None:
    """清空主题缓存（用于热更新或测试）"""
    _THEME_CACHE.clear()
