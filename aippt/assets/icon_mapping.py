from typing import Optional

ICON_MAPPING: dict[str, str] = {
    "增长": "lucide:trending-up",
    "上升": "lucide:trending-up",
    "下降": "lucide:trending-down",
    "数据": "lucide:bar-chart-2",
    "图表": "lucide:bar-chart-3",
    "分析": "lucide:pie-chart",
    "目标": "lucide:target",
    "完成": "lucide:check-circle",
    "进度": "lucide:activity",
    "团队": "lucide:users",
    "协作": "lucide:users",
    "人员": "lucide:user",
    "领导": "lucide:user-cog",
    "客户": "lucide:smile",
    "安全": "lucide:shield-check",
    "风险": "lucide:alert-triangle",
    "警告": "lucide:alert-circle",
    "保护": "lucide:shield",
    "成功": "lucide:trophy",
    "荣誉": "lucide:award",
    "创新": "lucide:lightbulb",
    "技术": "lucide:cpu",
    "产品": "lucide:package",
    "服务": "lucide:headphones",
    "财务": "lucide:dollar-sign",
    "营收": "lucide:wallet",
    "计划": "lucide:calendar",
    "时间": "lucide:clock",
    "地点": "lucide:map-pin",
    "文档": "lucide:file-text",
    "报告": "lucide:clipboard-list",
    "设置": "lucide:settings",
    "搜索": "lucide:search",
    "添加": "lucide:plus-circle",
    "删除": "lucide:trash-2",
    "编辑": "lucide:edit-3",
    "链接": "lucide:link",
    "云": "lucide:cloud",
    "手机": "lucide:smartphone",
    "电脑": "lucide:monitor",
    "年终": "lucide:party-popper",
    "总结": "lucide:book-open",
    "未来": "lucide:rocket",
    "挑战": "lucide:mountain",
    "机会": "lucide:key",
    # 英文同义词
    "growth": "lucide:trending-up",
    "chart": "lucide:bar-chart-3",
    "team": "lucide:users",
    "user": "lucide:user",
    "success": "lucide:trophy",
    "innovation": "lucide:lightbulb",
    "target": "lucide:target",
    "time": "lucide:clock",
    "location": "lucide:map-pin",
    "document": "lucide:file-text",
    "report": "lucide:clipboard-list",
}

ENGLISH_KEYWORDS = {
    "growth", "chart", "data", "team", "user", "success",
    "innovation", "target", "time", "location", "document",
    "report", "analysis", "goal", "risk", "security", "cloud",
    "mobile", "service", "product", "plan", "link", "search",
}


def map_to_icon(text: str, default_set: str = "lucide") -> str:
    if not text:
        return ""
    lower = text.lower().strip()
    for key, icon in ICON_MAPPING.items():
        if key in lower:
            return icon
    words = lower.replace("-", " ").replace("_", " ").split()
    for w in words:
        if w in ENGLISH_KEYWORDS:
            return f"{default_set}:{w}"
    return ""


def enrich_icon_mapping(page: dict) -> dict:
    page = dict(page)
    items = page.get("items", [])
    if not items:
        return page
    enriched = []
    for item in items:
        if isinstance(item, dict):
            item = dict(item)
            text = item.get("title", "") + " " + item.get("desc", "")
            icon_name = map_to_icon(text)
            if icon_name and not item.get("icon"):
                item["icon"] = icon_name
        enriched.append(item)
    page["items"] = enriched
    return page
