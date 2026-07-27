from typing import Any, Optional

from aippt.logger import logger

PAGE_TYPE_LAYOUT_MAP: dict[str, str] = {
    "cover": "hero",
    "numbered_list": "list_with_images",
    "catalog": "list_with_images",
    "kpi": "card_icons",
    "two_column": "dual_images",
    "timeline": "timeline_icons",
    "ending": "hero",
    "divider": "none",
    "table": "none",
}

PHOTO_PAGE_TYPES = {"cover", "numbered_list", "two_column", "ending"}
ICON_PAGE_TYPES = {"numbered_list", "catalog", "kpi", "timeline", "two_column"}


def build_asset_plan(page: dict[str, Any], scene: str = "", style: str = "") -> Optional[list[dict]]:
    page_type = page.get("page_type", "numbered_list")
    title = page.get("title", "")
    items = page.get("items", []) or []

    layout_mode = PAGE_TYPE_LAYOUT_MAP.get(page_type, "none")
    if layout_mode == "none":
        return None

    assets: list[dict] = []
    slot_id = 0

    if page_type in PHOTO_PAGE_TYPES:
        query = _build_photo_query(title, items, scene, style)
        if query:
            assets.append({
                "slot": f"img_{slot_id}",
                "query": query,
                "type": "photo",
                "orientation": "landscape" if page_type in ("cover", "ending") else "square",
                "count": 1,
                "role": "hero" if page_type in ("cover", "ending") else "card",
            })
            slot_id += 1

    if page_type in ICON_PAGE_TYPES and items:
        icon_count = min(len(items), 8)
        from aippt.assets.icon_mapping import map_to_icon
        for idx in range(icon_count):
            item = items[idx] if isinstance(items[idx], dict) else {"title": str(items[idx])}
            text = item.get("title", "") + " " + item.get("desc", "")
            icon_name = map_to_icon(text)
            if icon_name:
                assets.append({
                    "slot": f"icon_{idx}",
                    "query": icon_name,
                    "type": "icon",
                    "set": "lucide",
                    "size": 128,
                })

    if not assets:
        return None
    return assets


def _build_photo_query(title: str, items: list, scene: str, style: str) -> str:
    parts = []
    if scene:
        scene_keywords = {
            "年终总结": "business year-end review",
            "工作总结": "office work summary",
            "工作汇报": "business presentation",
            "工作计划": "business planning",
            "公司简介": "corporate office building",
            "述职报告": "business meeting",
            "职业规划": "career growth",
            "安全教育": "safety training",
            "产品发布": "product launch event",
        }
        kw = scene_keywords.get(scene, "")
        if kw:
            parts.append(kw)

    for item in items[:3]:
        if isinstance(item, dict):
            t = item.get("title", "")
            d = item.get("desc", "")
            if t:
                parts.append(t)
            if d and len(d) > 5:
                parts.append(d[:40])

    if style:
        style_keywords = {
            "商务蓝": "corporate professional",
            "极简灰": "minimal clean",
            "科技青": "modern technology",
            "柠檬黄": "creative vibrant",
            "安全橙": "industrial warning",
        }
        skw = style_keywords.get(style, "")
        if skw:
            parts.append(skw)
    else:
        parts.append("professional")

    parts.append("high-quality")
    return " ".join(parts[:6])
