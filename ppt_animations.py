"""
PPT 动画效果注入模块
功能：在 slide 上写入 <p:timing> 子元素，支持入场/退场/强调三类动画
依赖：lxml（python-pptx 自带依赖）

参考：ECMA-376 第 4 版 §19.5.1（timing）
ECMA-376 约束：p:childTnLst 必须是 p:cTn 的子元素（不能颠倒顺序）
"""
from typing import Any, Optional

from lxml import etree

from aippt.logger import logger


# ==================== 命名空间常量 ====================
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


# ==================== 动画效果目录 ====================
# 每种效果包含：
#   preset_class: animEffect 类别（entr/exit/emph）
#   preset_id: PowerPoint presetID 整数（ECMA-376 §18.56-18.58）
#   preset_subtype: 子类型整数（0 表示无）
#   filter: animEffect 的 filter 属性字符串（None 表示无 animEffect，仅用 anim）
#   transition: animEffect 的 transition 属性（in/out）
#   anim_elem: 行为元素类型（animEffect/anim/set）
#   dir_map: 方向别名 → preset_subtype 映射
#   description: 中文描述
#   best_for: 适用场景建议
ANIMATION_CATALOG = {
    # ---------- 入场动画（entrance）----------
    "appear": {
        "preset_class": "entr",
        "preset_id": 1,
        "preset_subtype": 0,
        "filter": None,
        "transition": None,
        "anim_elem": "set",
        "dir_map": {},
        "description": "出现（无动画，直接显示）",
        "best_for": "正文段落快速入场",
    },
    "fade": {
        "preset_class": "entr",
        "preset_id": 10,
        "preset_subtype": 0,
        "filter": "fade",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "淡入（柔和入场）",
        "best_for": "正文、卡片、数据块",
    },
    "fly_in": {
        "preset_class": "entr",
        "preset_id": 2,
        "preset_subtype": 4,
        "filter": "wipe(fromLeft)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {
            "from_left": 4, "from_right": 8, "from_top": 1, "from_bottom": 2,
            "from_top_left": 16, "from_top_right": 32,
            "from_bottom_left": 64, "from_bottom_right": 128,
        },
        "description": "飞入（按方向飞入）",
        "best_for": "重点条目、章节标题",
    },
    "float_in": {
        "preset_class": "entr",
        "preset_id": 12,
        "preset_subtype": 0,
        "filter": "fade",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {"up": 1, "down": 2},
        "description": "浮入（柔和上升入场）",
        "best_for": "副标题、说明文字",
    },
    "wipe": {
        "preset_class": "entr",
        "preset_id": 23,
        "preset_subtype": 8,
        "filter": "wipe(fromLeft)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {
            "from_left": 8, "from_right": 4, "from_top": 1, "from_bottom": 2,
        },
        "description": "擦除（按方向擦出）",
        "best_for": "进度条、时间轴",
    },
    "zoom": {
        "preset_class": "entr",
        "preset_id": 23,
        "preset_subtype": 0,
        "filter": "zoom(in)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "缩放（放大入场）",
        "best_for": "强调数据、KPI 数字",
    },
    "split": {
        "preset_class": "entr",
        "preset_id": 30,
        "preset_subtype": 0,
        "filter": "wipe(fromLeft)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "分裂（按方向分裂展开）",
        "best_for": "标题、章节分隔",
    },
    "box": {
        "preset_class": "entr",
        "preset_id": 22,
        "preset_subtype": 0,
        "filter": "box(in)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {"in": 0, "out": 1},
        "description": "方框（方框展开）",
        "best_for": "卡片式内容",
    },
    "dissolve": {
        "preset_class": "entr",
        "preset_id": 9,
        "preset_subtype": 0,
        "filter": "dissolve",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "溶解（像素化入场）",
        "best_for": "图像、装饰",
    },
    "swivel": {
        "preset_class": "entr",
        "preset_id": 17,
        "preset_subtype": 0,
        "filter": "wipe(fromLeft)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "旋转（旋转入场）",
        "best_for": "图标、装饰元素",
    },
    "bounce": {
        "preset_class": "entr",
        "preset_id": 26,
        "preset_subtype": 0,
        "filter": "fade",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "弹跳（弹跳入场）",
        "best_for": "趣味元素、强调",
    },
    "wheel": {
        "preset_class": "entr",
        "preset_id": 21,
        "preset_subtype": 1,
        "filter": "wheel(1)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "轮辐（辐射展开）",
        "best_for": "封面、特殊强调",
    },
    "blinds": {
        "preset_class": "entr",
        "preset_id": 7,
        "preset_subtype": 0,
        "filter": "blinds(horizontal)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "百叶窗（按方向展开）",
        "best_for": "图像、图片",
    },
    "checker": {
        "preset_class": "entr",
        "preset_id": 6,
        "preset_subtype": 0,
        "filter": "checkerboard(across)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "棋盘格（按方向展开）",
        "best_for": "图像、装饰",
    },
    "wedge": {
        "preset_class": "entr",
        "preset_id": 24,
        "preset_subtype": 0,
        "filter": "wedge",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "楔形（双向楔形展开）",
        "best_for": "特殊转场",
    },
    "random_bars": {
        "preset_class": "entr",
        "preset_id": 5,
        "preset_subtype": 0,
        "filter": "randombar(horizontal)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "随机条（随机条纹展开）",
        "best_for": "图像、装饰",
    },
    "strips": {
        "preset_class": "entr",
        "preset_id": 8,
        "preset_subtype": 0,
        "filter": "strips(downLeft)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "条带（按对角方向展开）",
        "best_for": "图像、装饰",
    },
    "plus": {
        "preset_class": "entr",
        "preset_id": 25,
        "preset_subtype": 0,
        "filter": "plus",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "十字展开（十字扩散）",
        "best_for": "特殊转场",
    },
    "circle": {
        "preset_class": "entr",
        "preset_id": 23,
        "preset_subtype": 0,
        "filter": "circle(out)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "圆形展开（圆形扩散）",
        "best_for": "图标、强调",
    },
    "diamond": {
        "preset_class": "entr",
        "preset_id": 27,
        "preset_subtype": 0,
        "filter": "diamond(out)",
        "transition": "in",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "菱形展开（菱形扩散）",
        "best_for": "图标、强调",
    },

    # ---------- 退场动画（exit）----------
    "disappear": {
        "preset_class": "exit",
        "preset_id": 1,
        "preset_subtype": 0,
        "filter": None,
        "transition": None,
        "anim_elem": "set",
        "dir_map": {},
        "description": "消失（无动画，直接隐藏）",
        "best_for": "快速退场",
    },
    "fade_out": {
        "preset_class": "exit",
        "preset_id": 10,
        "preset_subtype": 0,
        "filter": "fade",
        "transition": "out",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "淡出（柔和退场）",
        "best_for": "卡片、数据块",
    },
    "fly_out": {
        "preset_class": "exit",
        "preset_id": 2,
        "preset_subtype": 4,
        "filter": "wipe(fromLeft)",
        "transition": "out",
        "anim_elem": "animEffect",
        "dir_map": {
            "from_left": 4, "from_right": 8, "from_top": 1, "from_bottom": 2,
        },
        "description": "飞出（按方向飞出）",
        "best_for": "重点条目",
    },
    "wipe_out": {
        "preset_class": "exit",
        "preset_id": 23,
        "preset_subtype": 8,
        "filter": "wipe(fromLeft)",
        "transition": "out",
        "anim_elem": "animEffect",
        "dir_map": {
            "from_left": 8, "from_right": 4, "from_top": 1, "from_bottom": 2,
        },
        "description": "擦除退场",
        "best_for": "进度条、时间轴",
    },
    "zoom_out": {
        "preset_class": "exit",
        "preset_id": 23,
        "preset_subtype": 0,
        "filter": "zoom(out)",
        "transition": "out",
        "anim_elem": "animEffect",
        "dir_map": {},
        "description": "缩放退场",
        "best_for": "强调数据",
    },

    # ---------- 强调动画（emphasis）----------
    "pulse": {
        "preset_class": "emph",
        "preset_id": 1,
        "preset_subtype": 0,
        "filter": None,
        "transition": None,
        "anim_elem": "anim",
        "dir_map": {},
        "description": "脉冲（缩放强调）",
        "best_for": "强调元素",
    },
    "spin": {
        "preset_class": "emph",
        "preset_id": 5,
        "preset_subtype": 0,
        "filter": None,
        "transition": None,
        "anim_elem": "anim",
        "dir_map": {},
        "description": "旋转（360度旋转）",
        "best_for": "图标强调",
    },
    "grow_shrink": {
        "preset_class": "emph",
        "preset_id": 3,
        "preset_subtype": 0,
        "filter": None,
        "transition": None,
        "anim_elem": "anim",
        "dir_map": {},
        "description": "放大缩小",
        "best_for": "数据强调",
    },
    "bold_flash": {
        "preset_class": "emph",
        "preset_id": 6,
        "preset_subtype": 0,
        "filter": None,
        "transition": None,
        "anim_elem": "anim",
        "dir_map": {},
        "description": "加粗闪烁",
        "best_for": "文字强调",
    },
}


# ==================== 按页面类型推荐的动画配置 ====================
# 每个值为动画规范列表，符合 inject_animations 的 animations_spec 参数格式
RECOMMENDED_ANIMATIONS = {
    "COVER": [
        {"shape": "title", "effect": "fade", "trigger": "on_load", "duration_ms": 1000},
        {"shape": "subtitle", "effect": "fade", "trigger": "after_prev", "duration_ms": 800, "delay_ms": 300},
    ],
    "CHAPTER": [
        {"shape": "title", "effect": "fly_in", "trigger": "on_load", "duration_ms": 800, "dir": "from_left"},
        {"shape": "number", "effect": "fade", "trigger": "with_prev", "duration_ms": 600},
    ],
    "CONTENT": [
        {"shape": "title", "effect": "fade", "trigger": "on_load", "duration_ms": 600},
        {"shape": "body", "effect": "wipe", "trigger": "on_click",
         "text_build": "by_bullet", "dir": "from_left", "duration_ms": 400},
    ],
    "KPI": [
        {"shape": "title", "effect": "fade", "trigger": "on_load", "duration_ms": 600},
        {"shape": "number", "effect": "zoom", "trigger": "after_prev", "duration_ms": 800, "delay_ms": 200},
    ],
    "TIMELINE": [
        {"shape": "year", "effect": "fly_in", "trigger": "on_load", "duration_ms": 600, "dir": "from_left"},
        {"shape": "title", "effect": "fade", "trigger": "with_prev", "duration_ms": 500},
        {"shape": "desc", "effect": "fade", "trigger": "after_prev", "duration_ms": 500},
    ],
    "CHART": [
        {"shape": "title", "effect": "fade", "trigger": "on_load", "duration_ms": 600},
        {"shape": "title", "effect": "fade", "trigger": "after_prev", "duration_ms": 500, "delay_ms": 200},
    ],
    "TABLE": [
        {"shape": "title", "effect": "fade", "trigger": "on_load", "duration_ms": 600},
    ],
    "END": [
        {"shape": "title", "effect": "zoom", "trigger": "on_load", "duration_ms": 1000},
    ],
}


# ==================== 触发类型映射 ====================
# trigger → (nodeType, delay_value)
# delay_value: "indefinite" 表示点击触发，"0" 表示立即触发
TRIGGER_MAP = {
    "on_load": ("afterEffect", "0"),       # 进入页面后自动触发
    "on_click": ("clickEffect", "indefinite"),  # 点击触发
    "after_prev": ("afterEffect", "0"),    # 上一动画完成后触发
    "with_prev": ("withEffect", "0"),      # 与上一动画同时触发
}


# ==================== 核心实现 ====================
def _p(tag: str) -> str:
    """构造 p: 命名空间的 Element 标签"""
    return f"{{{NS_P}}}{tag}"


def _find_shape_by_role(slide: Any, role: str) -> Optional[int]:
    """
    根据 role（title/desc/number 等）定位 slide 中的 shape
    优先匹配 shape name 中包含 role 的，其次匹配文本特征

    :param slide: python-pptx Slide 对象
    :param role: 角色名（title/desc/number/year 等）
    :return: shape_id（int），找不到返回 None
    """
    candidates = []
    for shape in slide.shapes:
        name = (shape.name or "").lower()
        # 精确包含匹配
        if role == "title" and ("title" in name or "标题" in name):
            candidates.append(shape)
        elif role == "subtitle" and ("subtitle" in name or "副标题" in name):
            candidates.append(shape)
        elif role == "body" and ("body" in name or "正文" in name or "内容" in name
                                  or "content" in name or "desc" in name or "描述" in name):
            candidates.append(shape)
        elif role == "desc" and ("desc" in name or "描述" in name or "content" in name or "内容" in name):
            candidates.append(shape)
        elif role == "number" and ("number" in name or "数字" in name):
            candidates.append(shape)
        elif role == "year" and ("year" in name or "时间" in name):
            candidates.append(shape)

    if not candidates:
        # 兜底：取第一个有文本的 shape
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                candidates.append(shape)
                break

    if not candidates:
        return None

    # 取 shape_id
    try:
        return candidates[0].shape_id
    except Exception:
        return None


def _build_paragraph_tgt(parent_cBhvr: etree._Element,
                         shape_id: int,
                         paragraph_idx: int) -> None:
    """
    构建按段落动画的目标引用：p:tgtEl > p:spTgt + p:txEl > p:p（指定段落索引）

    :param parent_cBhvr: 父 cBhvr 元素
    :param shape_id: shape ID
    :param paragraph_idx: 段落索引（0-based）
    """
    tgtEl = etree.SubElement(parent_cBhvr, _p("tgtEl"))
    spTgt = etree.SubElement(tgtEl, _p("spTgt"))
    spTgt.set("spid", str(shape_id))
    txEl = etree.SubElement(spTgt, _p("txEl"))
    p_elem = etree.SubElement(txEl, _p("p"))
    p_elem.set("id", str(paragraph_idx))


def _build_by_bullet_nodes(spec: dict[str, Any],
                           shape_id: int,
                           paragraph_count: int,
                           duration_ms: int,
                           delay_ms: int,
                           trigger: str,
                           cTn_id_start: int,
                           bullet_delay_ms: int = 500,
                           bullet_order: Optional[list[int]] = None,
                           staggered: bool = False) -> list[etree._Element]:
    """
    构建按段落（bullet）逐步显示的动画节点列表

    每个段落生成一个独立的 <p:par> 节点，引用对应段落索引
    第一个段落按 trigger 触发，后续段落自动 with_prev/after_prev

    :param spec: ANIMATION_CATALOG 中的规范
    :param shape_id: 目标 shape 的 ID
    :param paragraph_count: 段落总数
    :param duration_ms: 单段时长（毫秒）
    :param delay_ms: 段间延迟（毫秒）
    :param trigger: 第一个段落的触发类型
    :param cTn_id_start: cTn ID 起始值
    :param bullet_delay_ms: 段间延迟（毫秒），默认 500ms；用于在每段 cTn 上设置 delay
    :param bullet_order: 段落播放顺序，默认 [0,1,2,...] 顺序播放；
                         传入 [2,0,1] 表示先播第3段再第1段再第2段；
                         长度必须等于 paragraph_count，否则忽略并警告
    :param staggered: 超强方案 P1 新增。True 时后续段落用 with_prev 触发（段落间重叠，
                      更动感）；False 时（默认）后续段落用 after_prev（依次衔接）
    :return: list of <p:par> Element
    """
    # 校验 bullet_order：长度必须等于 paragraph_count，否则忽略并警告
    if bullet_order is not None:
        if len(bullet_order) != paragraph_count:
            logger.warning(
                "bullet_order 长度 %d 与段落数 %d 不符，忽略 bullet_order 按顺序播放",
                len(bullet_order), paragraph_count,
            )
            bullet_order = None

    # 确定播放序列：bullet_order 指定顺序，否则 0..paragraph_count-1
    play_sequence = bullet_order if bullet_order is not None else list(range(paragraph_count))

    # 超强方案 P1：staggered 模式下后续段落使用 with_prev（段落重叠）
    subsequent_trigger = "with_prev" if staggered else "after_prev"

    nodes = []
    cTn_id = cTn_id_start

    for seq_idx, para_idx in enumerate(play_sequence):
        # 第一个段落按用户指定 trigger，后续段落按 staggered/sequential 决定
        para_trigger = trigger if seq_idx == 0 else subsequent_trigger
        node_type, delay_value = TRIGGER_MAP.get(para_trigger, TRIGGER_MAP["on_load"])

        outer_par = etree.Element(_p("par"))
        outer_cTn = etree.SubElement(outer_par, _p("cTn"))
        outer_cTn.set("id", str(cTn_id))
        outer_cTn.set("fill", "hold")
        outer_stCondLst = etree.SubElement(outer_cTn, _p("stCondLst"))
        outer_cond = etree.SubElement(outer_stCondLst, _p("cond"))
        outer_cond.set("delay", delay_value)
        outer_childTnLst = etree.SubElement(outer_cTn, _p("childTnLst"))

        # 中层 par
        mid_par = etree.SubElement(outer_childTnLst, _p("par"))
        mid_cTn = etree.SubElement(mid_par, _p("cTn"))
        mid_cTn.set("id", str(cTn_id + 1))
        mid_cTn.set("fill", "hold")
        mid_stCondLst = etree.SubElement(mid_cTn, _p("stCondLst"))
        mid_cond = etree.SubElement(mid_stCondLst, _p("cond"))
        mid_cond.set("delay", "0")
        mid_childTnLst = etree.SubElement(mid_cTn, _p("childTnLst"))

        # 内层 par
        inner_par = etree.SubElement(mid_childTnLst, _p("par"))
        inner_cTn = etree.SubElement(inner_par, _p("cTn"))
        inner_cTn.set("id", str(cTn_id + 2))
        inner_cTn.set("presetID", str(spec["preset_id"]))
        inner_cTn.set("presetClass", spec["preset_class"])
        inner_cTn.set("presetSubtype", str(spec["preset_subtype"]))
        inner_cTn.set("fill", "hold")
        inner_cTn.set("grpId", "0")
        inner_cTn.set("nodeType", node_type)
        inner_stCondLst = etree.SubElement(inner_cTn, _p("stCondLst"))
        inner_cond = etree.SubElement(inner_stCondLst, _p("cond"))
        # 段间延迟：第一段 delay=0
        # - sequential 模式：后续段落 delay=bullet_delay_ms * (seq_idx+1)（在前一段结束后额外延迟）
        # - staggered 模式（超强方案 P1）：后续段落 delay=bullet_delay_ms * seq_idx（从启动算起的绝对延迟，交错更紧凑）
        if seq_idx == 0:
            para_delay = 0
        elif staggered:
            para_delay = bullet_delay_ms * seq_idx
        else:
            para_delay = bullet_delay_ms * (seq_idx + 1)
        inner_cond.set("delay", str(para_delay))
        inner_childTnLst = etree.SubElement(inner_cTn, _p("childTnLst"))

        # 入场：先 set visible 再 animEffect，按段落引用
        if spec["preset_class"] == "entr":
            # <p:set> visibility visible
            set_elem = etree.SubElement(inner_childTnLst, _p("set"))
            set_cBhvr = etree.SubElement(set_elem, _p("cBhvr"))
            set_cTn = etree.SubElement(set_cBhvr, _p("cTn"))
            set_cTn.set("id", str(cTn_id + 3))
            set_cTn.set("dur", "1")
            set_cTn.set("fill", "hold")
            set_stCondLst = etree.SubElement(set_cTn, _p("stCondLst"))
            set_cond = etree.SubElement(set_stCondLst, _p("cond"))
            set_cond.set("delay", "0")
            _build_paragraph_tgt(set_cBhvr, shape_id, para_idx)
            set_attrNameLst = etree.SubElement(set_cBhvr, _p("attrNameLst"))
            set_attrName = etree.SubElement(set_attrNameLst, _p("attrName"))
            set_attrName.text = "style.visibility"
            set_to = etree.SubElement(set_elem, _p("to"))
            set_strVal = etree.SubElement(set_to, _p("strVal"))
            set_strVal.set("val", "visible")

            # animEffect
            if spec.get("filter"):
                animEffect = etree.SubElement(inner_childTnLst, _p("animEffect"))
                animEffect.set("transition", spec.get("transition", "in"))
                animEffect.set("filter", spec["filter"])
                ae_cBhvr = etree.SubElement(animEffect, _p("cBhvr"))
                ae_cTn = etree.SubElement(ae_cBhvr, _p("cTn"))
                ae_cTn.set("id", str(cTn_id + 4))
                ae_cTn.set("dur", str(duration_ms))
                _build_paragraph_tgt(ae_cBhvr, shape_id, para_idx)
        # 退场：animEffect 再 set hidden
        elif spec["preset_class"] == "exit":
            if spec.get("filter"):
                animEffect = etree.SubElement(inner_childTnLst, _p("animEffect"))
                animEffect.set("transition", spec.get("transition", "out"))
                animEffect.set("filter", spec["filter"])
                ae_cBhvr = etree.SubElement(animEffect, _p("cBhvr"))
                ae_cTn = etree.SubElement(ae_cBhvr, _p("cTn"))
                ae_cTn.set("id", str(cTn_id + 3))
                ae_cTn.set("dur", str(duration_ms))
                _build_paragraph_tgt(ae_cBhvr, shape_id, para_idx)
            set_elem = etree.SubElement(inner_childTnLst, _p("set"))
            set_cBhvr = etree.SubElement(set_elem, _p("cBhvr"))
            set_cTn = etree.SubElement(set_cBhvr, _p("cTn"))
            set_cTn.set("id", str(cTn_id + 4))
            set_cTn.set("dur", "1")
            set_cTn.set("fill", "hold")
            set_stCondLst = etree.SubElement(set_cTn, _p("stCondLst"))
            set_cond = etree.SubElement(set_stCondLst, _p("cond"))
            set_cond.set("delay", str(duration_ms))
            _build_paragraph_tgt(set_cBhvr, shape_id, para_idx)
            set_attrNameLst = etree.SubElement(set_cBhvr, _p("attrNameLst"))
            set_attrName = etree.SubElement(set_attrNameLst, _p("attrName"))
            set_attrName.text = "style.visibility"
            set_to = etree.SubElement(set_elem, _p("to"))
            set_strVal = etree.SubElement(set_to, _p("strVal"))
            set_strVal.set("val", "hidden")

        nodes.append(outer_par)
        cTn_id += 5

    return nodes


def _build_anim_effect_node(spec: dict[str, Any], shape_id: int,
                            duration_ms: int, delay_ms: int,
                            trigger: str, cTn_id_start: int) -> etree._Element:
    """
    构建单个动画的 <p:par> 节点（含 cTn + childTnLst + 行为元素）

    ECMA-376 约束：p:childTnLst 必须是 p:cTn 的子元素

    :param spec: ANIMATION_CATALOG 中的规范
    :param shape_id: 目标 shape 的 ID
    :param duration_ms: 时长（毫秒）
    :param delay_ms: 延迟（毫秒）
    :param trigger: 触发类型
    :param cTn_id_start: cTn ID 起始值（每个动画需要 3 个递增 ID）
    :return: <p:par> Element
    """
    node_type, delay_value = TRIGGER_MAP.get(trigger, TRIGGER_MAP["on_load"])
    preset_class = spec["preset_class"]
    preset_id = spec["preset_id"]
    preset_subtype = spec["preset_subtype"]

    # 构建三层嵌套 par > cTn > childTnLst > par > cTn > childTnLst > par > cTn
    # 外层 par（包装层）
    outer_par = etree.Element(_p("par"))
    outer_cTn = etree.SubElement(outer_par, _p("cTn"))
    outer_cTn.set("id", str(cTn_id_start))
    outer_cTn.set("fill", "hold")
    outer_stCondLst = etree.SubElement(outer_cTn, _p("stCondLst"))
    outer_cond = etree.SubElement(outer_stCondLst, _p("cond"))
    outer_cond.set("delay", delay_value)

    # ECMA-376 约束：p:childTnLst 必须是 p:cTn 的子元素
    outer_childTnLst = etree.SubElement(outer_cTn, _p("childTnLst"))

    # 中层 par
    mid_par = etree.SubElement(outer_childTnLst, _p("par"))
    mid_cTn = etree.SubElement(mid_par, _p("cTn"))
    mid_cTn.set("id", str(cTn_id_start + 1))
    mid_cTn.set("fill", "hold")
    mid_stCondLst = etree.SubElement(mid_cTn, _p("stCondLst"))
    mid_cond = etree.SubElement(mid_stCondLst, _p("cond"))
    mid_cond.set("delay", "0")
    mid_childTnLst = etree.SubElement(mid_cTn, _p("childTnLst"))

    # 内层 par（实际动画节点）
    inner_par = etree.SubElement(mid_childTnLst, _p("par"))
    inner_cTn = etree.SubElement(inner_par, _p("cTn"))
    inner_cTn.set("id", str(cTn_id_start + 2))
    inner_cTn.set("presetID", str(preset_id))
    inner_cTn.set("presetClass", preset_class)
    inner_cTn.set("presetSubtype", str(preset_subtype))
    inner_cTn.set("fill", "hold")
    inner_cTn.set("grpId", "0")
    inner_cTn.set("nodeType", node_type)

    inner_stCondLst = etree.SubElement(inner_cTn, _p("stCondLst"))
    inner_cond = etree.SubElement(inner_stCondLst, _p("cond"))
    inner_cond.set("delay", str(delay_ms))

    # childTnLst 包含实际行为元素
    inner_childTnLst = etree.SubElement(inner_cTn, _p("childTnLst"))

    anim_elem_type = spec["anim_elem"]
    filter_str = spec.get("filter")
    transition = spec.get("transition")

    # 入场动画：先设置 visible，再 animEffect
    if preset_class == "entr":
        # <p:set> 设置 visibility visible
        set_elem = etree.SubElement(inner_childTnLst, _p("set"))
        set_cBhvr = etree.SubElement(set_elem, _p("cBhvr"))
        set_cTn = etree.SubElement(set_cBhvr, _p("cTn"))
        set_cTn.set("id", str(cTn_id_start + 3))
        set_cTn.set("dur", "1")
        set_cTn.set("fill", "hold")
        set_stCondLst = etree.SubElement(set_cTn, _p("stCondLst"))
        set_cond = etree.SubElement(set_stCondLst, _p("cond"))
        set_cond.set("delay", "0")
        set_tgtEl = etree.SubElement(set_cBhvr, _p("tgtEl"))
        set_spTgt = etree.SubElement(set_tgtEl, _p("spTgt"))
        set_spTgt.set("spid", str(shape_id))
        set_attrNameLst = etree.SubElement(set_cBhvr, _p("attrNameLst"))
        set_attrName = etree.SubElement(set_attrNameLst, _p("attrName"))
        set_attrName.text = "style.visibility"
        set_to = etree.SubElement(set_elem, _p("to"))
        set_strVal = etree.SubElement(set_to, _p("strVal"))
        set_strVal.set("val", "visible")

        # 如果是 appear，仅 set 即可
        if anim_elem_type == "animEffect" and filter_str:
            _append_anim_effect(inner_childTnLst, spec, shape_id,
                                cTn_id_start + 4, duration_ms)

    # 退场动画：先 animEffect，再 set hidden
    elif preset_class == "exit":
        if anim_elem_type == "animEffect" and filter_str:
            _append_anim_effect(inner_childTnLst, spec, shape_id,
                                cTn_id_start + 3, duration_ms)
        # <p:set> 设置 visibility hidden
        set_elem = etree.SubElement(inner_childTnLst, _p("set"))
        set_cBhvr = etree.SubElement(set_elem, _p("cBhvr"))
        set_cTn = etree.SubElement(set_cBhvr, _p("cTn"))
        set_cTn.set("id", str(cTn_id_start + 4))
        set_cTn.set("dur", "1")
        set_cTn.set("fill", "hold")
        set_stCondLst = etree.SubElement(set_cTn, _p("stCondLst"))
        set_cond = etree.SubElement(set_stCondLst, _p("cond"))
        set_cond.set("delay", str(duration_ms))
        set_tgtEl = etree.SubElement(set_cBhvr, _p("tgtEl"))
        set_spTgt = etree.SubElement(set_tgtEl, _p("spTgt"))
        set_spTgt.set("spid", str(shape_id))
        set_attrNameLst = etree.SubElement(set_cBhvr, _p("attrNameLst"))
        set_attrName = etree.SubElement(set_attrNameLst, _p("attrName"))
        set_attrName.text = "style.visibility"
        set_to = etree.SubElement(set_elem, _p("to"))
        set_strVal = etree.SubElement(set_to, _p("strVal"))
        set_strVal.set("val", "hidden")

    # 强调动画：使用 anim（属性动画）
    elif preset_class == "emph":
        _append_emph_anim(inner_childTnLst, spec, shape_id,
                          cTn_id_start + 3, duration_ms)

    return outer_par


def _append_anim_effect(parent: etree._Element, spec: dict[str, Any],
                       shape_id: int, cTn_id: int, duration_ms: int) -> None:
    """
    追加 <p:animEffect> 元素到 parent
    """
    animEffect = etree.SubElement(parent, _p("animEffect"))
    animEffect.set("transition", spec.get("transition", "in"))
    animEffect.set("filter", spec["filter"])
    cBhvr = etree.SubElement(animEffect, _p("cBhvr"))
    cTn = etree.SubElement(cBhvr, _p("cTn"))
    cTn.set("id", str(cTn_id))
    cTn.set("dur", str(duration_ms))
    tgtEl = etree.SubElement(cBhvr, _p("tgtEl"))
    spTgt = etree.SubElement(tgtEl, _p("spTgt"))
    spTgt.set("spid", str(shape_id))


def _append_emph_anim(parent: etree._Element, spec: dict[str, Any],
                     shape_id: int, cTn_id: int, duration_ms: int) -> None:
    """
    追加强调动画的 <p:anim> 元素到 parent
    根据 preset_id 选择不同的属性动画
    """
    preset_id = spec["preset_id"]
    anim = etree.SubElement(parent, _p("anim"))
    cBhvr = etree.SubElement(anim, _p("cBhvr"))
    cTn = etree.SubElement(cBhvr, _p("cTn"))
    cTn.set("id", str(cTn_id))
    cTn.set("dur", str(duration_ms))
    cTn.set("autoRev", "1")
    tgtEl = etree.SubElement(cBhvr, _p("tgtEl"))
    spTgt = etree.SubElement(tgtEl, _p("spTgt"))
    spTgt.set("spid", str(shape_id))
    attrNameLst = etree.SubElement(cBhvr, _p("attrNameLst"))

    if preset_id == 1:  # pulse：缩放
        attrName = etree.SubElement(attrNameLst, _p("attrName"))
        attrName.text = "ScaleX"
        from_el = etree.SubElement(anim, _p("from"))
        fltVal = etree.SubElement(from_el, _p("fltVal"))
        fltVal.set("val", "1.0")
        to_el = etree.SubElement(anim, _p("to"))
        fltVal2 = etree.SubElement(to_el, _p("fltVal"))
        fltVal2.set("val", "1.1")
    elif preset_id == 5:  # spin：旋转
        attrName = etree.SubElement(attrNameLst, _p("attrName"))
        attrName.text = "Rotation"
        from_el = etree.SubElement(anim, _p("from"))
        fltVal = etree.SubElement(from_el, _p("fltVal"))
        fltVal.set("val", "0.0")
        to_el = etree.SubElement(anim, _p("to"))
        fltVal2 = etree.SubElement(to_el, _p("fltVal"))
        fltVal2.set("val", "360.0")
    elif preset_id == 3:  # grow_shrink：缩放
        attrName = etree.SubElement(attrNameLst, _p("attrName"))
        attrName.text = "ScaleX"
        from_el = etree.SubElement(anim, _p("from"))
        fltVal = etree.SubElement(from_el, _p("fltVal"))
        fltVal.set("val", "1.0")
        to_el = etree.SubElement(anim, _p("to"))
        fltVal2 = etree.SubElement(to_el, _p("fltVal"))
        fltVal2.set("val", "1.5")
    elif preset_id == 6:  # bold_flash：加粗
        attrName = etree.SubElement(attrNameLst, _p("attrName"))
        attrName.text = "style.fontWeight"
        from_el = etree.SubElement(anim, _p("from"))
        strVal = etree.SubElement(from_el, _p("strVal"))
        strVal.set("val", "normal")
        to_el = etree.SubElement(anim, _p("to"))
        strVal2 = etree.SubElement(to_el, _p("strVal"))
        strVal2.set("val", "bold")
    else:
        # 默认：缩放
        attrName = etree.SubElement(attrNameLst, _p("attrName"))
        attrName.text = "ScaleX"
        from_el = etree.SubElement(anim, _p("from"))
        fltVal = etree.SubElement(from_el, _p("fltVal"))
        fltVal.set("val", "1.0")
        to_el = etree.SubElement(anim, _p("to"))
        fltVal2 = etree.SubElement(to_el, _p("fltVal"))
        fltVal2.set("val", "1.1")


def _build_timing_tree(animation_nodes: list[etree._Element]) -> etree._Element:
    """
    构建完整的 <p:timing> 树结构

    结构：
    <p:timing>
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    [animation_nodes...]
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst>...</p:prevCondLst>
                <p:nextCondLst>...</p:nextCondLst>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
    </p:timing>

    :param animation_nodes: 已构建好的 <p:par> 动画节点列表
    :return: <p:timing> 根 Element
    """
    timing = etree.Element(_p("timing"))
    tnLst = etree.SubElement(timing, _p("tnLst"))
    par_root = etree.SubElement(tnLst, _p("par"))
    cTn_root = etree.SubElement(par_root, _p("cTn"))
    cTn_root.set("id", "1")
    cTn_root.set("dur", "indefinite")
    cTn_root.set("restart", "never")
    cTn_root.set("nodeType", "tmRoot")
    # childTnLst 必须是 cTn 的子元素
    childTnLst_root = etree.SubElement(cTn_root, _p("childTnLst"))

    # mainSeq
    seq = etree.SubElement(childTnLst_root, _p("seq"))
    seq.set("concurrent", "1")
    seq.set("nextAc", "seek")
    cTn_seq = etree.SubElement(seq, _p("cTn"))
    cTn_seq.set("id", "2")
    cTn_seq.set("dur", "indefinite")
    cTn_seq.set("nodeType", "mainSeq")
    # childTnLst
    childTnLst_seq = etree.SubElement(cTn_seq, _p("childTnLst"))

    # 插入所有动画节点
    for node in animation_nodes:
        childTnLst_seq.append(node)

    # prevCondLst
    prevCondLst = etree.SubElement(seq, _p("prevCondLst"))
    prev_cond = etree.SubElement(prevCondLst, _p("cond"))
    prev_cond.set("evt", "onPrev")
    prev_cond.set("delay", "0")
    prev_tgtEl = etree.SubElement(prev_cond, _p("tgtEl"))
    etree.SubElement(prev_tgtEl, _p("sldTgt"))

    # nextCondLst
    nextCondLst = etree.SubElement(seq, _p("nextCondLst"))
    next_cond = etree.SubElement(nextCondLst, _p("cond"))
    next_cond.set("evt", "onNext")
    next_cond.set("delay", "0")
    next_tgtEl = etree.SubElement(next_cond, _p("tgtEl"))
    etree.SubElement(next_tgtEl, _p("sldTgt"))

    return timing


def inject_animations(slide: Any, animations_spec: list[dict[str, Any]],
                     slide_type: str = "CONTENT") -> bool:
    """
    为 slide 注入动画效果（直接修改 slide XML，添加 <p:timing> 子元素）

    :param slide: python-pptx 的 Slide 对象
    :param animations_spec: 动画配置列表
        [{"shape": "title", "effect": "fade", "trigger": "on_load", "duration_ms": 800,
          "delay_ms": 0, "dir": "from_left"}, ...]
        - shape: 角色名（title/subtitle/desc/number/year）或具体 shape_id
        - effect: 效果名（见 ANIMATION_CATALOG）
        - trigger: 触发类型（on_load/on_click/after_prev/with_prev）
        - duration_ms: 时长（毫秒），默认 800
        - delay_ms: 延迟（毫秒），默认 0
        - dir: 方向（可选，部分效果支持）
    :param slide_type: 页面类型（COVER/CHAPTER/CONTENT/KPI/TIMELINE/END），用于 fallback
    :return: True 表示注入成功，False 表示失败

    示例：
        # 整体动画
        inject_animations(slide, [{"shape": "title", "effect": "fade", "trigger": "on_load"}], "CONTENT")
        # 按段落动画（一次点击显示一条 bullet）
        inject_animations(slide, [{"shape": "body", "effect": "wipe", "trigger": "on_click",
                                    "text_build": "by_bullet", "dir": "from_left"}], "CONTENT")
    """
    if not animations_spec or not isinstance(animations_spec, list):
        return False

    animation_nodes = []
    cTn_id_counter = 3  # ID 1 和 2 已被 root 和 mainSeq 占用

    for spec_idx, anim_spec in enumerate(animations_spec):
        if not isinstance(anim_spec, dict):
            continue
        effect_name = anim_spec.get("effect")
        if not effect_name or effect_name not in ANIMATION_CATALOG:
            print(f"⚠️  动画 {spec_idx}: 不支持的效果 {effect_name}")
            continue

        catalog_spec = dict(ANIMATION_CATALOG[effect_name])

        # 处理方向（dir 覆盖 preset_subtype）
        user_dir = anim_spec.get("dir")
        if user_dir and catalog_spec.get("dir_map") and user_dir in catalog_spec["dir_map"]:
            catalog_spec["preset_subtype"] = catalog_spec["dir_map"][user_dir]
            # 同步更新 filter（基础方向映射）
            filter_str = catalog_spec.get("filter", "")
            if filter_str and user_dir.startswith("from_"):
                # 简单的方向字符串替换
                dir_str = user_dir.replace("from_", "")
                # 首字母大写
                dir_cap = dir_str.title().replace("_", "")
                if "wipe(from" in filter_str:
                    filter_str = f"wipe(from{dir_cap})"
                    catalog_spec["filter"] = filter_str

        # 定位 shape
        shape_role = anim_spec.get("shape")
        shape_id = None
        if isinstance(shape_role, int):
            shape_id = shape_role
        elif shape_role:
            shape_id = _find_shape_by_role(slide, shape_role)
        if shape_id is None:
            print(f"⚠️  动画 {spec_idx}: 未找到 shape '{shape_role}'，跳过")
            continue

        # 触发类型
        trigger = anim_spec.get("trigger", "on_load")
        if trigger not in TRIGGER_MAP:
            print(f"⚠️  动画 {spec_idx}: 不支持的触发类型 {trigger}，使用 on_load")
            trigger = "on_load"

        # 时长与延迟
        duration_ms = int(anim_spec.get("duration_ms", 800))
        delay_ms = int(anim_spec.get("delay_ms", 0))

        # 按段落（by_bullet）构建分支
        if anim_spec.get("text_build") == "by_bullet":
            # 找到目标 shape 对象，统计段落
            target_shape = None
            for s in slide.shapes:
                try:
                    if s.shape_id == shape_id:
                        target_shape = s
                        break
                except Exception:
                    pass
            if target_shape is None or not target_shape.has_text_frame:
                # 兜底：无段落信息时退回整体动画
                node = _build_anim_effect_node(
                    catalog_spec, shape_id, duration_ms, delay_ms, trigger, cTn_id_counter
                )
                animation_nodes.append(node)
                cTn_id_counter += 5
                continue

            para_count = len(target_shape.text_frame.paragraphs)
            # 仅对有内容的段落生成动画；空段落跳过
            non_empty = [p for p in target_shape.text_frame.paragraphs
                         if p.text.strip()]
            para_count = max(1, len(non_empty))

            # 超强方案 P0/P1：从 spec 读取 bullet_delay_ms 与 sequence 透传给节点树
            spec_bullet_delay = anim_spec.get("bullet_delay_ms")
            if spec_bullet_delay is not None:
                spec_bullet_delay = int(spec_bullet_delay)
            else:
                spec_bullet_delay = 500  # 默认段间延迟
            spec_sequence = anim_spec.get("sequence", "sequential")
            staggered_mode = (spec_sequence == "staggered")

            # 每个 by_bullet 段落消耗 5 个 ID
            nodes = _build_by_bullet_nodes(
                catalog_spec, shape_id, para_count, duration_ms, delay_ms,
                trigger, cTn_id_counter,
                bullet_delay_ms=spec_bullet_delay,
                staggered=staggered_mode,
            )
            animation_nodes.extend(nodes)
            cTn_id_counter += 5 * para_count
            continue

        # 构建动画节点（每个动画消耗 5 个 ID）
        node = _build_anim_effect_node(
            catalog_spec, shape_id, duration_ms, delay_ms, trigger, cTn_id_counter
        )
        animation_nodes.append(node)
        cTn_id_counter += 5

    if not animation_nodes:
        return False

    # 构建 <p:timing> 树
    timing_tree = _build_timing_tree(animation_nodes)

    # 获取 slide 根元素，移除已有 timing，插入新的
    sld_elem = slide._element
    # 移除已有的 timing
    for child in list(sld_elem):
        if etree.QName(child).localname == "timing":
            sld_elem.remove(child)

    # 按 schema 顺序插入：cSld, clrMapOvr, transition, timing
    schema_order = ["cSld", "clrMapOvr", "transition", "AlternateContent", "timing"]
    new_idx = schema_order.index("timing")
    insert_pos = len(sld_elem)
    for i, child in enumerate(sld_elem):
        child_local = etree.QName(child).localname
        if child_local in schema_order:
            child_idx = schema_order.index(child_local)
            if child_idx > new_idx:
                insert_pos = i
                break
    sld_elem.insert(insert_pos, timing_tree)

    return True


def validate_animations(animations_spec: list[dict[str, Any]],
                       slide_num: Optional[int] = None) -> list[str]:
    """
    校验动画配置是否合法

    :param animations_spec: 动画配置列表
    :param slide_num: 页码（用于警告信息）
    :return: 警告列表（空列表表示无警告）
    """
    warnings = []
    if not animations_spec:
        return warnings

    if not isinstance(animations_spec, list):
        warnings.append(f"页{slide_num or '?'}: 动画配置应为列表")
        return warnings

    for i, spec in enumerate(animations_spec):
        if not isinstance(spec, dict):
            warnings.append(f"页{slide_num or '?'} 动画[{i}]: 配置应为 dict")
            continue

        effect = spec.get("effect")
        if not effect:
            warnings.append(f"页{slide_num or '?'} 动画[{i}]: 缺少 effect 字段")
            continue
        if effect not in ANIMATION_CATALOG:
            warnings.append(f"页{slide_num or '?'} 动画[{i}]: 不支持的效果 {effect}")
            continue

        trigger = spec.get("trigger", "on_load")
        if trigger not in TRIGGER_MAP:
            warnings.append(f"页{slide_num or '?'} 动画[{i}]: 不支持的触发类型 {trigger}")

        duration = spec.get("duration_ms", 800)
        if not isinstance(duration, (int, float)) or duration <= 0:
            warnings.append(f"页{slide_num or '?'} 动画[{i}]: duration_ms 应为正数")

        delay = spec.get("delay_ms", 0)
        if not isinstance(delay, (int, float)) or delay < 0:
            warnings.append(f"页{slide_num or '?'} 动画[{i}]: delay_ms 应为非负数")

        shape = spec.get("shape")
        if not shape:
            warnings.append(f"页{slide_num or '?'} 动画[{i}]: 缺少 shape 字段")

        catalog_spec = ANIMATION_CATALOG[effect]
        user_dir = spec.get("dir")
        if user_dir and catalog_spec.get("dir_map") and user_dir not in catalog_spec["dir_map"]:
            warnings.append(
                f"页{slide_num or '?'} 动画[{i}]: 效果 {effect} 不支持方向 {user_dir}"
                f"（支持: {list(catalog_spec['dir_map'].keys())}）"
            )

        # text_build 校验：目前仅支持 by_bullet
        text_build = spec.get("text_build")
        if text_build and text_build != "by_bullet":
            warnings.append(
                f"页{slide_num or '?'} 动画[{i}]: 不支持的 text_build {text_build}（仅支持 by_bullet）"
            )
        # 强调类动画不支持 by_bullet（无段落可见性语义）
        if text_build == "by_bullet" and catalog_spec["preset_class"] == "emph":
            warnings.append(
                f"页{slide_num or '?'} 动画[{i}]: 强调类效果 {effect} 不支持 text_build=by_bullet"
            )

    return warnings


def list_animations() -> dict[str, Any]:
    """列出所有可用动画效果（按类别分组）

    :return: 包含 entrance/exit/emphasis/total 四个键的字典
    """
    entrance = [(k, v) for k, v in ANIMATION_CATALOG.items() if v["preset_class"] == "entr"]
    exit_a = [(k, v) for k, v in ANIMATION_CATALOG.items() if v["preset_class"] == "exit"]
    emph = [(k, v) for k, v in ANIMATION_CATALOG.items() if v["preset_class"] == "emph"]
    return {
        "entrance": [{"effect": k, "description": v["description"],
                      "best_for": v["best_for"]} for k, v in entrance],
        "exit": [{"effect": k, "description": v["description"],
                  "best_for": v["best_for"]} for k, v in exit_a],
        "emphasis": [{"effect": k, "description": v["description"],
                      "best_for": v["best_for"]} for k, v in emph],
        "total": len(ANIMATION_CATALOG),
    }


if __name__ == "__main__":
    catalog = list_animations()
    print(f"动画目录总数: {catalog['total']}")
    print(f"入场动画: {len(catalog['entrance'])} 种")
    print(f"退场动画: {len(catalog['exit'])} 种")
    print(f"强调动画: {len(catalog['emphasis'])} 种")
    print("\n入场动画：")
    for a in catalog["entrance"]:
        print(f"  {a['effect']:15s} - {a['description']} ({a['best_for']})")
    print("\n退场动画：")
    for a in catalog["exit"]:
        print(f"  {a['effect']:15s} - {a['description']} ({a['best_for']})")
    print("\n强调动画：")
    for a in catalog["emphasis"]:
        print(f"  {a['effect']:15s} - {a['description']} ({a['best_for']})")
    print("\n推荐动画配置（页面类型）：")
    for ptype, anims in RECOMMENDED_ANIMATIONS.items():
        print(f"  {ptype}: {len(anims)} 个动画")
