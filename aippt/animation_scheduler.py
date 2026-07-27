"""
动画注入调度层
功能：按「识别角色 → 匹配规则 → 计算时序 → 逐层注入」流程执行
       支持延迟、时长、触发方式精细控制，支持 by_bullet 逐段动画

设计约束：
    - 入参使用纯 Python 类型（不绑定 pptx 对象），便于测试
    - 三层次序：entry → emphasis → exit，每层内部按角色优先级排列
    - 100% 向后兼容，不修改现有 inject_animations / inject_transition 签名
"""
from typing import Any, Optional

from aippt.animation_themes import (
    get_theme, build_page_transition_spec, build_page_animations_spec,
    ANIMATION_THEMES, _PAGE_TYPE_SHAPES as THEME_PAGE_TYPE_SHAPES,
    SLIDE_TYPE_TO_PAGE_TYPE,
)
from ppt_animations import inject_animations, RECOMMENDED_ANIMATIONS
from aippt.logger import logger


# ==================== 角色优先级常量 ====================

ROLE_PRIORITY_MAP = {
    "title":     0,
    "subtitle":  1,
    "number":    2,     # 章节编号（divider/catalog）
    "kpi_value": 3,     # KPI 数值
    "body":      4,
    "desc":      5,
    "year":      6,
}

ROLE_PRIORITY_ORDER = sorted(ROLE_PRIORITY_MAP.keys(), key=lambda r: ROLE_PRIORITY_MAP[r])

# 默认时长（毫秒）per role
ROLE_DEFAULT_DURATION_MS = {
    "title":     800,
    "subtitle":  600,
    "number":    500,
    "kpi_value": 800,
    "body":      400,
    "desc":      500,
    "year":      400,
}

# 角色触发类型（首角色 on_load，后续 after_prev）
ROLE_TRIGGER = {
    "title":     "on_load",
    "subtitle":  "after_prev",
    "number":    "with_prev",
    "kpi_value": "after_prev",
    "body":      "on_click",
    "desc":      "after_prev",
    "year":      "on_load",
}

# 支持 by_bullet 的页面类型
# 约束：by_bullet 仅允许列表类、目录类、时间轴类页面（cover/divider/kpi 禁止）
BY_BULLET_PAGE_TYPES = {"numbered_list", "two_column", "skill_percent", "preset_titles",
                        "catalog", "timeline"}

# 注入时使用的 slide_type → 角色列表映射
# 与 animation_themes._PAGE_TYPE_SHAPES 同步，但使用统一的角色名
PAGE_TYPE_ROLES = {
    "cover":         [("title",), ("subtitle",)],
    "catalog":       [("title",), ("number", "body")],
    "divider":       [("title",), ("number",)],
    "numbered_list": [("title",), ("body",)],
    "kpi":           [("title",), ("kpi_value",)],
    "timeline":      [("year",), ("title",), ("desc",)],
    "two_column":    [("title",), ("body",)],
    "skill_percent": [("title",), ("body",)],
    "preset_titles": [("title",), ("body",)],
    "chart":         [("title",)],
    "table":         [("title",)],
    "ending":        [("title",)],
}


# ==================== 时序计算 ====================

def _get_role_priority(role: str) -> int:
    return ROLE_PRIORITY_MAP.get(role, 99)


def _calculate_role_delay(role: str, base_delay_ms: int = 200) -> int:
    """基于角色优先级计算延迟时间

    :param role: 角色名
    :param base_delay_ms: 优先级增量步长
    :return: 延迟毫秒数
    """
    priority = _get_role_priority(role)
    return priority * base_delay_ms


def _calculate_role_duration(role: str, effect: str, fallback_ms: int = 600) -> int:
    """获取角色默认动画时长

    :param role: 角色名
    :param effect: 效果名（用于 zoom 等特殊效果加长）
    :param fallback_ms: 兜底时长
    :return: 时长毫秒数
    """
    base = ROLE_DEFAULT_DURATION_MS.get(role, fallback_ms)
    if effect in ("zoom", "float_in"):
        return base + 200
    return base


# ==================== 三层层规范构建 ====================

def build_entry_specs(page_type: str, entry_effect: str, by_bullet: bool,
                      direction: Optional[str] = None,
                      bullet_delay_ms: Optional[int] = None,
                      page_delay_ms: Optional[int] = None,
                      intensity_factor: float = 1.0,
                      sequence: str = "sequential") -> list[dict]:
    """构建入场动画 spec 列表（按角色优先级排序）

    超强方案 P0 新增透传字段：
        - direction: 入场方向（如 from_bottom/from_left），透传到 inject_animations
        - bullet_delay_ms: by_bullet 段间延迟（毫秒）
        - page_delay_ms: 单页延迟（毫秒），覆盖默认角色延迟

    超强方案 P1 新增透传字段：
        - intensity_factor: 强度倍率，应用到 duration（1.0=默认，0.7=low，1.4=high）
        - sequence: by_bullet 播放顺序（sequential/staggered/custom），透传到 inject_animations

    :param page_type: 页面类型
    :param entry_effect: 入场效果名
    :param by_bullet: 是否启用 by_bullet
    :param direction: 入场方向（可选）
    :param bullet_delay_ms: by_bullet 段间延迟（可选）
    :param page_delay_ms: 单页延迟（可选，覆盖默认）
    :param intensity_factor: 强度倍率（默认 1.0）
    :param sequence: by_bullet 顺序（默认 sequential）
    :return: spec 列表，每项含 shape/effect/trigger/duration_ms/delay_ms
    """
    roles = PAGE_TYPE_ROLES.get(page_type, [("title",)])
    specs: list[dict] = []
    group_delay = 0

    for group in roles:
        group_roles: list[str] = list(group)
        for idx, role in enumerate(group_roles):
            # intensity_factor 应用到 duration（超强方案 P1）
            duration = int(_calculate_role_duration(role, entry_effect) * intensity_factor)
            # page_delay_ms 覆盖默认角色延迟
            delay = page_delay_ms if page_delay_ms is not None else _calculate_role_delay(role)
            # group 内首角色用 ROLE_TRIGGER 配置，后续角色用 with_prev 同时触发
            role_trigger = ROLE_TRIGGER.get(role, "after_prev")
            trigger = role_trigger if idx == 0 else "with_prev"

            spec = {
                "shape": role,
                "effect": entry_effect,
                "trigger": trigger,
                "duration_ms": duration,
                "delay_ms": delay,
            }

            # 入场方向透传（仅对 fly_in/wipe 等方向敏感效果有意义）
            if direction:
                spec["dir"] = direction

            enable_by_bullet = by_bullet and page_type in BY_BULLET_PAGE_TYPES and role in ("body", "desc")
            if enable_by_bullet:
                spec["text_build"] = "by_bullet"
                # by_bullet 段间延迟透传
                if bullet_delay_ms is not None:
                    spec["bullet_delay_ms"] = bullet_delay_ms
                # sequence 透传（超强方案 P1）：inject_animations 据此调整段落 trigger
                if sequence and sequence != "sequential":
                    spec["sequence"] = sequence

            specs.append(spec)
            group_delay = max(group_delay, delay + duration)

    return specs


def build_emphasis_specs(page_type: str, emphasis_effect: str) -> list[dict]:
    """构建强调动画 spec 列表

    强调动画作用于 title/kpi_value 角色，在入场完成后自动播放。

    :param page_type: 页面类型
    :param emphasis_effect: 强调效果名
    :return: spec 列表
    """
    roles = PAGE_TYPE_ROLES.get(page_type, [("title",)])
    target_roles = {"title", "kpi_value"}
    specs: list[dict] = []

    for group in roles:
        for role in group:
            if role in target_roles:
                specs.append({
                    "shape": role,
                    "effect": emphasis_effect,
                    "trigger": "after_prev",
                    "duration_ms": 600,
                    "delay_ms": 200,
                })

    return specs


def build_exit_specs(page_type: str, exit_effect: str) -> list[dict]:
    """构建退场动画 spec 列表（按角色优先级逆序）

    退场时次要角色先退，主要角色后退。

    :param page_type: 页面类型
    :param exit_effect: 退场效果名
    :return: spec 列表
    """
    roles = PAGE_TYPE_ROLES.get(page_type, [("title",)])
    specs: list[dict] = []

    reversed_roles: list[str] = []
    for group in reversed(roles):
        reversed_roles.extend(reversed(list(group)))

    for role in reversed_roles:
        specs.append({
            "shape": role,
            "effect": exit_effect,
            "trigger": "after_prev",
            "duration_ms": 500,
            "delay_ms": 100,
        })

    return specs


# ==================== 主调度函数 ====================

def schedule_slide_animations(
    page_type: str,
    theme_name: Optional[str] = None,
    page_animations: Optional[dict] = None,
) -> list[dict]:
    """为单页调度生成完整的动画 spec 列表

    执行流程：
        1. 解析动画配置（优先级：显式 > 主题覆盖 > 默认）
        2. 构建三层动画 spec（entry → emphasis → exit）
        3. 合并为一个有序 spec 列表

    :param page_type: 页面类型（小写，如 cover/kpi/numbered_list）
    :param theme_name: 动画主题名（business/tech/formal），None 不使用主题
    :param page_animations: 该页显式 animations 配置 dict，None 表示未设置
    :return: 合并后的动画 spec 列表，供 inject_animations 使用
    """
    # Step 1: 解析动画配置
    entry_effect = None
    exit_effect = None
    emphasis_effect = None
    by_bullet = False
    # 超强方案 P0 新增：方向、段间延迟、单页延迟
    direction = None
    bullet_delay_ms = None
    page_delay_ms = None
    # 超强方案 P1 新增：强度、顺序
    intensity = None
    sequence = None

    if page_animations is not None and isinstance(page_animations, dict):
        entry_effect = page_animations.get("entry")
        exit_effect = page_animations.get("exit")
        emphasis_effect = page_animations.get("emphasis")
        by_bullet = page_animations.get("by_bullet", False)
        direction = page_animations.get("dir")
        bullet_delay_ms = page_animations.get("bullet_delay_ms")
        page_delay_ms = page_animations.get("delay_ms")
        intensity = page_animations.get("intensity")
        sequence = page_animations.get("sequence")

    if entry_effect is None and theme_name:
        try:
            theme = get_theme(theme_name)
            page_cfg = theme.get("page_overrides", {}).get(page_type, {})
            a_cfg = page_cfg.get("animations")
            if a_cfg:
                entry_effect = entry_effect or a_cfg.get("entry")
                exit_effect = exit_effect or a_cfg.get("exit")
                emphasis_effect = emphasis_effect or a_cfg.get("emphasis")
                by_bullet = by_bullet if page_animations is not None else a_cfg.get("by_bullet", False)
                # 主题的 dir/bullet_delay_ms/delay_ms 作为兜底默认（单页显式优先）
                direction = direction or a_cfg.get("dir")
                bullet_delay_ms = bullet_delay_ms or a_cfg.get("bullet_delay_ms")
                page_delay_ms = page_delay_ms if page_delay_ms is not None else a_cfg.get("delay_ms")
                # 主题不提供 intensity/sequence，仅单页显式可设置
        except KeyError:
            pass

    if entry_effect is None:
        slide_type = next(
            (st for st, pt in SLIDE_TYPE_TO_PAGE_TYPE.items() if pt == page_type),
            "CONTENT"
        )
        rec = RECOMMENDED_ANIMATIONS.get(slide_type, [])
        if rec:
            entry_effect = rec[0].get("effect", "fade")
            by_bullet = any(s.get("text_build") == "by_bullet" for s in rec)
        else:
            entry_effect = "fade"

    by_bullet = bool(by_bullet) if isinstance(by_bullet, bool) else by_bullet in (True, "true", "True", 1)

    # A005: 不支持 by_bullet 的页面类型自动关闭
    if by_bullet and page_type not in BY_BULLET_PAGE_TYPES:
        logger.warning("A005: 页面类型 %s 不支持 by_bullet，已自动关闭", page_type)
        by_bullet = False

    # 超强方案 P1：intensity → duration 倍率 / bullet_delay 默认值
    # low=0.7x  med=1.0x  high=1.4x；同时为未设置 bullet_delay_ms 提供默认
    intensity_factor = 1.0
    intensity_default_bullet_delay = None
    if intensity == "low":
        intensity_factor = 0.7
        intensity_default_bullet_delay = 150
    elif intensity == "high":
        intensity_factor = 1.4
        intensity_default_bullet_delay = 400
    elif intensity == "med":
        intensity_factor = 1.0
        intensity_default_bullet_delay = 250

    # intensity 应用到 bullet_delay_ms（仅在未显式设置时使用默认）
    if bullet_delay_ms is None and intensity_default_bullet_delay is not None:
        bullet_delay_ms = intensity_default_bullet_delay

    # 超强方案 P1：sequence → by_bullet 顺序
    # sequential（默认）= 段落依次入场
    # staggered = 段落交错入场（更动感，通过更短 bullet_delay + with_prev 实现）
    # custom = 同 sequential，由用户通过 bullet_delay_ms 自定义
    # 这里仅记录到 spec，由 inject_animations 在 by_bullet 节点树构建时读取
    sequence_mode = sequence or "sequential"

    # Step 2: 构建三层动画 spec
    all_specs: list[dict] = []

    entry_specs = build_entry_specs(
        page_type, entry_effect, by_bullet,
        direction=direction,
        bullet_delay_ms=bullet_delay_ms,
        page_delay_ms=page_delay_ms,
        intensity_factor=intensity_factor,
        sequence=sequence_mode,
    )
    all_specs.extend(entry_specs)

    if emphasis_effect:
        emph_specs = build_emphasis_specs(page_type, emphasis_effect)
        all_specs.extend(emph_specs)

    if exit_effect:
        exit_specs = build_exit_specs(page_type, exit_effect)
        all_specs.extend(exit_specs)

    return all_specs


def schedule_transition(
    page_type: str,
    theme_name: Optional[str] = None,
    page_transition: Optional[str] = None,
    speed: str = "med",
    default_fallback: Optional[str] = "fade",
    page_transition_option: Optional[str] = None,
) -> Optional[dict]:
    """调度单页转场配置

    优先级：单页显式 > 主题 page_overrides > 主题 global_transition

    :param page_type: 页面类型
    :param theme_name: 动画主题名
    :param page_transition: 单页显式转场名
    :param speed: 转场速度
    :param default_fallback: 无任何配置时的默认转场（默认 fade），None 表示不使用
    :param page_transition_option: 转场选项（仅 morph 支持 byObject/byWord/byChar），
                                   超强方案 P0 新增
    :return: transition spec dict，或 None
    """
    t_name = page_transition
    t_option = page_transition_option

    if t_name is None and theme_name:
        try:
            theme = get_theme(theme_name)
            page_cfg = theme.get("page_overrides", {}).get(page_type, {})
            t_name = page_cfg.get("transition") or theme.get("global_transition")
            # morph 选项优先级：单页显式 > page_overrides > global_transition_option
            if t_name == "morph" and not t_option:
                t_option = page_cfg.get("transition_option") or theme.get("global_transition_option")
        except KeyError:
            pass

    if t_name is None and default_fallback:
        t_name = default_fallback

    if t_name and t_name != "none":
        return build_page_transition_spec(t_name, speed, option=t_option)

    return None


# ==================== 面向渲染器的高层接口 ====================

def inject_page_effects(
    slide: Any,
    page_type: str,
    theme_name: Optional[str] = None,
    page_animations: Optional[dict] = None,
    page_transition: Optional[str] = None,
    page_transition_option: Optional[str] = None,
) -> dict[str, Any]:
    """为单页幻灯片注入动画 + 转场（调度层一站式入口）

    支持两种 page_animations 格式：
        - dict: 动画配置 {"entry": "fade", "by_bullet": True, "dir": "from_bottom"} → 调度层解析
        - list: 预构建的 inject spec 列表 → 直接注入（来自 _apply_animation_theme）

    超强方案 P0 新增：
        - page_animations 支持 dir/bullet_delay_ms/delay_ms 字段透传
        - page_transition_option 支持 morph 的 byObject/byWord/byChar 选项

    :param slide: python-pptx Slide 对象
    :param page_type: 页面类型（小写）
    :param theme_name: 动画主题名
    :param page_animations: 单页动画配置（dict）或预构建 spec 列表（list）
    :param page_transition: 单页显式转场名
    :param page_transition_option: 转场选项（仅 morph 支持），超强方案 P0 新增
    :return: {"transition": bool, "animations": bool} 注入结果
    """
    result = {"transition": False, "animations": False}

    t_spec = schedule_transition(page_type, theme_name, page_transition,
                                 page_transition_option=page_transition_option)
    if t_spec:
        try:
            from ppt_transitions import inject_transition as do_transition
            do_transition(slide, t_spec)
            result["transition"] = True
        except Exception as e:
            logger.warning("转场注入失败 (page=%s): %s", page_type, e)

    # 兼容两种 page_animations 格式
    a_specs: Optional[list] = None
    if isinstance(page_animations, list):
        a_specs = page_animations
    elif isinstance(page_animations, dict):
        a_specs = schedule_slide_animations(page_type, theme_name, page_animations)
    else:
        a_specs = schedule_slide_animations(page_type, theme_name, None)

    if a_specs:
        try:
            slide_type = next(
                (st for st, pt in SLIDE_TYPE_TO_PAGE_TYPE.items() if pt == page_type),
                "CONTENT"
            )
            inject_animations(slide, a_specs, slide_type)
            result["animations"] = True
        except Exception as e:
            logger.warning("动画注入失败 (page=%s): %s", page_type, e)

    return result


def batch_inject_effects(
    prs: Any,
    page_types: dict[int, str],
    theme_name: Optional[str] = None,
    page_animations_map: Optional[dict[int, dict]] = None,
    page_transition_map: Optional[dict[int, str]] = None,
) -> dict[str, Any]:
    """批量注入多页的动画 + 转场

    :param prs: Presentation 对象
    :param page_types: {page_num: page_type} 映射
    :param theme_name: 动画主题名
    :param page_animations_map: {page_num: animations_dict} 显式动画覆盖
    :param page_transition_map: {page_num: transition_name} 显式转场覆盖
    :return: {"injected": int, "details": {page_num: result}}
    """
    if page_animations_map is None:
        page_animations_map = {}
    if page_transition_map is None:
        page_transition_map = {}

    total = 0
    details: dict[int, dict] = {}

    for page_num, slide in enumerate(prs.slides, 1):
        pt = page_types.get(page_num, "numbered_list")
        pa = page_animations_map.get(page_num)
        ptc = page_transition_map.get(page_num)
        r = inject_page_effects(slide, pt, theme_name, pa, ptc)
        details[page_num] = r
        if r["transition"] or r["animations"]:
            total += 1

    return {"injected": total, "details": details}


def list_animation_themes() -> list[str]:
    """列出所有可用动画主题"""
    return list(ANIMATION_THEMES.keys())
