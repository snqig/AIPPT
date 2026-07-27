"""
PPT 动画预设主题模块
功能：提供 6 套预设动画/转场主题（business/tech/formal/cinematic/dynamic-impact/minimal-plus），
      根据页面类型自动选择合适的转场与动画效果，实现一键风格切换。
依赖：ppt_animations.py 的 ANIMATION_CATALOG、ppt_transitions.py 的 TRANSITION_CATALOG

主题优先级（高 → 低）：
    1. outline.json 单页显式 transition/animations 字段
    2. --animation-theme 主题的 page_overrides
    3. --animation-theme 主题的 global_transition
    4. --transitions / --animations 全局参数

设计约束：100% 向后兼容，theme 未传入时行为与原版完全一致。

超强方案新增（P0）：
    - cinematic：电影感叙事，morph 转场为主，fly_in from_bottom + 较长 bullet_delay
    - dynamic-impact：高视觉冲击，vortex/switch 谨慎使用，bounce + 强 emphasis
    - minimal-plus：极简增强，cut/fade + appear + by_bullet 保留控制力
    - 主题支持 dir / bullet_delay_ms / intensity / rhythm 透传字段
"""
from typing import Optional

from aippt.logger import logger


# ==================== 页面类型 → shape 角色映射 ====================
# 用于将主题的简化 animations 配置 ({entry, exit, emphasis, by_bullet})
# 展开为 inject_animations 所需的 spec 列表
# 每项：(shape 角色, 触发类型, 默认时长 ms)
_PAGE_TYPE_SHAPES = {
    "cover": [("title", "on_load", 1000), ("subtitle", "after_prev", 800)],
    "catalog": [("title", "on_load", 600), ("body", "on_click", 400)],
    "divider": [("title", "on_load", 800)],
    "numbered_list": [("title", "on_load", 600), ("body", "on_click", 400)],
    "kpi": [("title", "on_load", 600), ("number", "after_prev", 800)],
    "timeline": [("year", "on_load", 600), ("title", "with_prev", 500), ("desc", "after_prev", 500)],
    "two_column": [("title", "on_load", 600), ("body", "on_click", 400)],
    "skill_percent": [("title", "on_load", 600), ("body", "on_click", 400)],
    "preset_titles": [("title", "on_load", 600), ("body", "on_click", 400)],
    "chart": [("title", "on_load", 600)],
    "table": [("title", "on_load", 600)],
    "ending": [("title", "on_load", 1000)],
}


# 渲染器 slide_type（大写）→ outline page_type（小写）映射
# 用于渲染引擎在无 outline 上下文时根据 meta 推断的 slide_type 反查 page_type
SLIDE_TYPE_TO_PAGE_TYPE = {
    "COVER": "cover",
    "CHAPTER": "divider",
    "CONTENT": "numbered_list",
    "KPI": "kpi",
    "TIMELINE": "timeline",
    "CHART": "chart",
    "TABLE": "table",
    "END": "ending",
}


# ==================== 3 套预设主题 ====================
# 每套主题包含：
#   - description: 主题风格说明
#   - global_transition: 全局默认转场名（page_overrides 未覆盖时使用）
#   - page_overrides: {page_type: {transition, animations: {entry, exit, emphasis, by_bullet}}}
ANIMATION_THEMES = {
    # ---------- 简约商务 ----------
    "business": {
        "description": "简约商务：以 fade/push 为主，克制柔和，列表页用 fly_in + by_bullet，KPI 用 zoom，封面/分隔用 fade。"
                       "适合工作汇报、述职报告、年终总结等正式商务场景。",
        "global_transition": "fade",
        "page_overrides": {
            "cover":         {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "catalog":       {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "divider":       {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "numbered_list": {"transition": "push", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,      "by_bullet": True}},
            "two_column":    {"transition": "push", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,      "by_bullet": True}},
            "skill_percent": {"transition": "push", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,      "by_bullet": True}},
            "preset_titles": {"transition": "push", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,      "by_bullet": True}},
            "kpi":           {"transition": "fade", "animations": {"entry": "zoom",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "timeline":      {"transition": "push", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,      "by_bullet": False}},
            "chart":         {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "table":         {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "ending":        {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
        },
    },
    # ---------- 活力科技 ----------
    "tech": {
        "description": "活力科技：以 zoom/flip 为主，动感明快，列表页用 wipe + by_bullet，KPI 用 zoom + emphasis:pulse，封面用 zoom。"
                       "适合产品发布、技术分享、创意提案等活力场景。",
        "global_transition": "zoom",
        "page_overrides": {
            "cover":         {"transition": "zoom", "animations": {"entry": "zoom",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "catalog":       {"transition": "zoom", "animations": {"entry": "zoom",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "divider":       {"transition": "flip", "animations": {"entry": "zoom",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "numbered_list": {"transition": "flip", "animations": {"entry": "wipe",   "exit": None, "emphasis": None,      "by_bullet": True}},
            "two_column":    {"transition": "flip", "animations": {"entry": "wipe",   "exit": None, "emphasis": None,      "by_bullet": True}},
            "skill_percent": {"transition": "flip", "animations": {"entry": "wipe",   "exit": None, "emphasis": None,      "by_bullet": True}},
            "preset_titles": {"transition": "flip", "animations": {"entry": "wipe",   "exit": None, "emphasis": None,      "by_bullet": True}},
            "kpi":           {"transition": "zoom", "animations": {"entry": "zoom",   "exit": None, "emphasis": "pulse",   "by_bullet": False}},
            "timeline":      {"transition": "flip", "animations": {"entry": "zoom",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "chart":         {"transition": "zoom", "animations": {"entry": "zoom",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "table":         {"transition": "zoom", "animations": {"entry": "wipe",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "ending":        {"transition": "zoom", "animations": {"entry": "zoom",   "exit": None, "emphasis": None,      "by_bullet": False}},
        },
    },
    # ---------- 沉稳正式 ----------
    "formal": {
        "description": "沉稳正式：以 fade/wipe 为主，庄重克制，列表页用 fade + by_bullet，KPI 用 fade，封面用 fade，几乎不用 emphasis。"
                       "适合政府汇报、金融报告、学术答辩等庄重场景。",
        "global_transition": "fade",
        "page_overrides": {
            "cover":         {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "catalog":       {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "divider":       {"transition": "wipe", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "numbered_list": {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": True}},
            "two_column":    {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": True}},
            "skill_percent": {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": True}},
            "preset_titles": {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": True}},
            "kpi":           {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "timeline":      {"transition": "wipe", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "chart":         {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "table":         {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
            "ending":        {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None,      "by_bullet": False}},
        },
    },
    # ---------- 电影感叙事（超强方案 P0 新增）----------
    "cinematic": {
        "description": "电影感叙事：优先 morph 转场（byObject），列表 fly_in from_bottom + by_bullet + 较长段间 delay，"
                       "封面 zoom + 标题延迟强调，KPI 大数字 zoom + pulse。"
                       "适合故事型汇报、融资路演、年终回顾等叙事性强的场景。",
        "global_transition": "morph",
        "global_transition_option": "byObject",  # morph 选项：byObject/byWord/byChar
        "rhythm": "slow-build",                   # 节奏曲线：slow-build 段间 delay 较长
        "page_overrides": {
            "cover":         {"transition": "morph", "animations": {"entry": "zoom",   "exit": None, "emphasis": "pulse",  "by_bullet": False, "delay_ms": 200}},
            "catalog":       {"transition": "morph", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,     "by_bullet": True,  "dir": "from_bottom", "bullet_delay_ms": 500}},
            "divider":       {"transition": "morph", "animations": {"entry": "zoom",   "exit": None, "emphasis": "spin",   "by_bullet": False}},
            "numbered_list": {"transition": "morph", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,     "by_bullet": True,  "dir": "from_bottom", "bullet_delay_ms": 600}},
            "two_column":    {"transition": "morph", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,     "by_bullet": True,  "dir": "from_bottom", "bullet_delay_ms": 500}},
            "skill_percent": {"transition": "morph", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,     "by_bullet": True,  "dir": "from_bottom", "bullet_delay_ms": 500}},
            "preset_titles": {"transition": "morph", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,     "by_bullet": True,  "dir": "from_bottom", "bullet_delay_ms": 500}},
            "kpi":           {"transition": "fade",  "animations": {"entry": "zoom",   "exit": None, "emphasis": "pulse",  "by_bullet": False, "delay_ms": 300}},
            "timeline":      {"transition": "morph", "animations": {"entry": "fly_in", "exit": None, "emphasis": None,     "by_bullet": True,  "dir": "from_left",   "bullet_delay_ms": 500}},
            "chart":         {"transition": "morph", "animations": {"entry": "zoom",   "exit": None, "emphasis": None,     "by_bullet": False}},
            "table":         {"transition": "morph", "animations": {"entry": "wipe",   "exit": None, "emphasis": None,     "by_bullet": False}},
            "ending":        {"transition": "morph", "animations": {"entry": "zoom",   "exit": None, "emphasis": "pulse",  "by_bullet": False, "delay_ms": 200}},
        },
    },
    # ---------- 高视觉冲击（超强方案 P0 新增）----------
    "dynamic-impact": {
        "description": "高视觉冲击：全局以 zoom 为主，列表页 bounce + by_bullet 快速段间，KPI zoom + 高频 pulse，"
                       "分隔页 flip，封面/结尾 zoom + grow_shrink 强调。"
                       "适合产品发布、内部动员会、年轻团队等高冲击场景。谨慎使用 vortex/switch。",
        "global_transition": "zoom",
        "rhythm": "fast-start",                   # 节奏曲线：fast-start 段间 delay 较短
        "page_overrides": {
            "cover":         {"transition": "zoom",   "animations": {"entry": "bounce",   "exit": None, "emphasis": "grow_shrink", "by_bullet": False}},
            "catalog":       {"transition": "zoom",   "animations": {"entry": "fly_in",   "exit": None, "emphasis": None,          "by_bullet": True,  "bullet_delay_ms": 200}},
            "divider":       {"transition": "flip",   "animations": {"entry": "zoom",     "exit": None, "emphasis": "spin",         "by_bullet": False}},
            "numbered_list": {"transition": "zoom",   "animations": {"entry": "bounce",   "exit": None, "emphasis": None,          "by_bullet": True,  "bullet_delay_ms": 200}},
            "two_column":    {"transition": "zoom",   "animations": {"entry": "fly_in",   "exit": None, "emphasis": None,          "by_bullet": True,  "bullet_delay_ms": 200}},
            "skill_percent": {"transition": "zoom",   "animations": {"entry": "fly_in",   "exit": None, "emphasis": "pulse",       "by_bullet": True,  "bullet_delay_ms": 200}},
            "preset_titles": {"transition": "zoom",   "animations": {"entry": "bounce",   "exit": None, "emphasis": None,          "by_bullet": True,  "bullet_delay_ms": 200}},
            "kpi":           {"transition": "zoom",   "animations": {"entry": "zoom",     "exit": None, "emphasis": "pulse",       "by_bullet": False, "delay_ms": 200}},
            "timeline":      {"transition": "flip",   "animations": {"entry": "fly_in",   "exit": None, "emphasis": None,          "by_bullet": True,  "dir": "from_left", "bullet_delay_ms": 200}},
            "chart":         {"transition": "zoom",   "animations": {"entry": "zoom",     "exit": None, "emphasis": "pulse",       "by_bullet": False}},
            "table":         {"transition": "zoom",   "animations": {"entry": "wipe",     "exit": None, "emphasis": None,          "by_bullet": False}},
            "ending":        {"transition": "zoom",   "animations": {"entry": "zoom",     "exit": None, "emphasis": "grow_shrink", "by_bullet": False}},
        },
    },
    # ---------- 极简增强（超强方案 P0 新增）----------
    "minimal-plus": {
        "description": "极简增强：以 cut/fade 为主，列表页 appear + by_bullet（几乎无动画感但保留逐条控制力），"
                       "几乎无 emphasis，封面/结尾 fade。适合极简汇报、纯内容输出、追求干净的场景。",
        "global_transition": "fade",
        "rhythm": "steady",                       # 节奏曲线：steady 稳定
        "page_overrides": {
            "cover":         {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None, "by_bullet": False}},
            "catalog":       {"transition": "cut",  "animations": {"entry": "appear", "exit": None, "emphasis": None, "by_bullet": True,  "bullet_delay_ms": 100}},
            "divider":       {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None, "by_bullet": False}},
            "numbered_list": {"transition": "fade", "animations": {"entry": "appear", "exit": None, "emphasis": None, "by_bullet": True,  "bullet_delay_ms": 100}},
            "two_column":    {"transition": "fade", "animations": {"entry": "appear", "exit": None, "emphasis": None, "by_bullet": True,  "bullet_delay_ms": 100}},
            "skill_percent": {"transition": "fade", "animations": {"entry": "appear", "exit": None, "emphasis": None, "by_bullet": True,  "bullet_delay_ms": 100}},
            "preset_titles": {"transition": "fade", "animations": {"entry": "appear", "exit": None, "emphasis": None, "by_bullet": True,  "bullet_delay_ms": 100}},
            "kpi":           {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None, "by_bullet": False}},
            "timeline":      {"transition": "fade", "animations": {"entry": "appear", "exit": None, "emphasis": None, "by_bullet": True,  "bullet_delay_ms": 150}},
            "chart":         {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None, "by_bullet": False}},
            "table":         {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None, "by_bullet": False}},
            "ending":        {"transition": "fade", "animations": {"entry": "fade",   "exit": None, "emphasis": None, "by_bullet": False}},
        },
    },
}


def get_theme(name: str) -> dict:
    """获取指定主题配置

    :param name: 主题名（business/tech/formal）
    :return: 主题配置 dict（含 description/global_transition/page_overrides）
    :raises KeyError: 主题不存在时抛出，包含可用主题列表
    """
    if name not in ANIMATION_THEMES:
        raise KeyError(
            f"未知动画主题: {name}，可用主题: {list(ANIMATION_THEMES.keys())}"
        )
    return ANIMATION_THEMES[name]


def list_themes() -> list[str]:
    """返回所有可用主题名列表"""
    return list(ANIMATION_THEMES.keys())


def build_page_transition_spec(transition_name: str, speed: str = "med",
                               option: Optional[str] = None) -> dict:
    """将主题的 transition 名转换为 inject_transition 所需的 spec dict

    :param transition_name: 转场名（如 fade/push/zoom/flip/morph），见 TRANSITION_CATALOG
    :param speed: 速度（slow/med/fast），默认 med
    :param option: 转场选项，目前仅 morph 支持 byObject/byWord/byChar（超强方案 P0 新增）
    :return: {"type": transition_name, "speed": speed, "option": option}；
             transition_name 为空或 none 时返回 None
    """
    if not transition_name or transition_name == "none":
        return None
    spec = {"type": transition_name, "speed": speed}
    # morph 转场支持 byObject/byWord/byChar 选项
    if transition_name == "morph" and option:
        spec["option"] = option
    return spec


def build_page_animations_spec(page_type: str, anims_cfg: dict) -> list[dict]:
    """将主题的 animations 配置展开为 inject_animations 所需的 spec 列表

    主题配置格式（与 outline.json 的 animations 字段一致）：
        {"entry": "fade", "exit": None, "emphasis": "pulse", "by_bullet": True,
         "dir": "from_bottom", "bullet_delay_ms": 500, "delay_ms": 200}
    展开后的 spec 列表格式（供 inject_animations 消费）：
        [{"shape": "title", "effect": "fade", "trigger": "on_load", "duration_ms": 600}, ...]

    超强方案 P0 新增透传字段：
        - dir: 入场方向（如 from_bottom/from_left/from_right），透传到 inject_animations
        - bullet_delay_ms: by_bullet 段间延迟（毫秒），透传到 text_build 配置
        - delay_ms: 单页延迟（毫秒），覆盖默认 delay

    :param page_type: 页面类型（小写，如 cover/kpi/numbered_list）
    :param anims_cfg: 主题 animations 配置 dict
    :return: anim spec 列表；anims_cfg 为空或无入场效果时返回空列表
    """
    if not anims_cfg:
        return []

    entry = anims_cfg.get("entry")
    exit_eff = anims_cfg.get("exit")
    emphasis = anims_cfg.get("emphasis")
    by_bullet = anims_cfg.get("by_bullet", False)
    # 超强方案 P0 新增字段透传
    direction = anims_cfg.get("dir")              # 入场方向
    bullet_delay_ms = anims_cfg.get("bullet_delay_ms")  # by_bullet 段间延迟
    page_delay_ms = anims_cfg.get("delay_ms")     # 单页延迟

    # 无入场效果则不生成动画（退场/强调依附于入场之后）
    if not entry:
        return []

    shapes = _PAGE_TYPE_SHAPES.get(page_type, [("title", "on_load", 600)])
    spec: list[dict] = []
    for shape, trigger, duration in shapes:
        item = {
            "shape": shape,
            "effect": entry,
            "trigger": trigger,
            "duration_ms": duration,
        }
        # 入场方向透传（仅对 fly_in/wipe 等方向敏感效果有意义）
        if direction:
            item["dir"] = direction
        # 单页延迟透传（覆盖默认 group delay）
        if page_delay_ms is not None:
            item["delay_ms"] = page_delay_ms
        # 列表类页面的 body shape 启用按段落逐步显示
        if by_bullet and shape == "body":
            item["text_build"] = "by_bullet"
            # by_bullet 段间延迟透传
            if bullet_delay_ms is not None:
                item["bullet_delay_ms"] = bullet_delay_ms
        spec.append(item)

    # 退场动画（附加在入场之后，作用于 title）
    if exit_eff:
        spec.append({
            "shape": "title",
            "effect": exit_eff,
            "trigger": "after_prev",
            "duration_ms": 600,
        })

    # 强调动画（附加在最后，作用于 title）
    if emphasis:
        spec.append({
            "shape": "title",
            "effect": emphasis,
            "trigger": "after_prev",
            "duration_ms": 600,
        })

    return spec


def resolve_page_transition(page_type: str, theme: dict, page_explicit=None) -> str:
    """解析单页应使用的 transition 名

    优先级：单页显式 > 主题 page_overrides > 主题 global_transition

    :param page_type: 页面类型（小写）
    :param theme: 主题配置 dict
    :param page_explicit: outline 中该页显式设置的 transition（None 表示未设置）
    :return: transition 名（如 "fade"），可为 None
    """
    # 1. 单页显式配置优先
    if page_explicit is not None:
        return page_explicit
    # 2. 主题 page_overrides
    page_cfg = theme.get("page_overrides", {}).get(page_type)
    if page_cfg and page_cfg.get("transition"):
        return page_cfg["transition"]
    # 3. 主题 global_transition
    return theme.get("global_transition")


def resolve_page_animations(page_type: str, theme: dict, page_explicit=None) -> dict:
    """解析单页应使用的 animations 配置

    优先级：单页显式 > 主题 page_overrides

    :param page_type: 页面类型（小写）
    :param theme: 主题配置 dict
    :param page_explicit: outline 中该页显式设置的 animations dict（None 表示未设置）
    :return: animations 配置 dict，可为 None
    """
    # 1. 单页显式配置优先
    if page_explicit is not None:
        return page_explicit
    # 2. 主题 page_overrides
    page_cfg = theme.get("page_overrides", {}).get(page_type)
    if page_cfg and page_cfg.get("animations"):
        return page_cfg["animations"]
    return None


if __name__ == "__main__":
    # 自测：列出主题概览
    print(f"可用主题: {list_themes()}")
    for name in list_themes():
        theme = get_theme(name)
        print(f"\n【{name}】{theme['description']}")
        print(f"  全局转场: {theme['global_transition']}")
        print(f"  页面覆盖: {len(theme['page_overrides'])} 种页面类型")
        # 抽样展开一个 numbered_list 的 animations spec
        nl_cfg = theme["page_overrides"].get("numbered_list", {}).get("animations")
        if nl_cfg:
            spec = build_page_animations_spec("numbered_list", nl_cfg)
            print(f"  numbered_list 动画 spec ({len(spec)} 项): {spec}")
