"""
PPT 转场效果注入模块
功能：在 slide 的 XML 上写入 <p:transition> 子元素，支持 38 种转场效果
依赖：lxml（python-pptx 自带依赖）

参考：ECMA-376 第 4 版 + PowerPoint 2010+ 扩展（p14 命名空间）
"""
from lxml import etree


# ==================== 命名空间常量 ====================
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

NSMAP_COMMON = {
    "p": NS_P,
    "p14": NS_P14,
    "mc": NS_MC,
    "r": NS_R,
}

# 速度映射（毫秒）
SPEED_MAP = {
    "slow": 1500,
    "med": 800,
    "fast": 400,
}

# 速度枚举值（写入 spd 属性）
SPEED_ENUM = {
    "slow": "slow",
    "med": "med",
    "fast": "fast",
}


# ==================== 转场效果目录 ====================
# 每种转场包含：
#   ns: 命名空间（"p" 或 "p14"）
#   xml_tag: XML 标签名（不含命名空间前缀）
#   attrs: 默认属性字典
#   dir_values: 支持的方向值列表（用于 dir 映射）
#   description: 中文描述
TRANSITION_CATALOG = {
    # ---------- ECMA-376 核心 19 种 ----------
    "fade": {
        "ns": "p",
        "xml_tag": "fade",
        "attrs": {},
        "dir_values": [],
        "description": "淡入淡出（最常用，柔和过渡）",
    },
    "cut": {
        "ns": "p",
        "xml_tag": "cut",
        "attrs": {},
        "dir_values": [],
        "description": "切换（无过渡直接切换）",
    },
    "push": {
        "ns": "p",
        "xml_tag": "push",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "推入（新页推走旧页）",
    },
    "cover": {
        "ns": "p",
        "xml_tag": "cover",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "覆盖（新页覆盖旧页）",
    },
    "pull": {
        "ns": "p",
        "xml_tag": "pull",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "拉出（旧页被拉走）",
    },
    "wipe": {
        "ns": "p",
        "xml_tag": "wipe",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "擦除（按方向擦出）",
    },
    "dissolve": {
        "ns": "p",
        "xml_tag": "dissolve",
        "attrs": {},
        "dir_values": [],
        "description": "溶解（像素化淡出）",
    },
    "split": {
        "ns": "p",
        "xml_tag": "split",
        "attrs": {"orient": "horz", "dir": "out"},
        "dir_values": ["in", "out"],
        "description": "分裂（按方向分裂展开）",
    },
    "zoom": {
        "ns": "p",
        "xml_tag": "zoom",
        "attrs": {"dir": "out"},
        "dir_values": ["in", "out"],
        "description": "缩放（放大或缩小切换）",
    },
    "wheel": {
        "ns": "p",
        "xml_tag": "wheel",
        "attrs": {"spokes": "1"},
        "dir_values": [],
        "description": "轮辐（辐射状展开，spokes 1-8）",
    },
    "blinds": {
        "ns": "p",
        "xml_tag": "blinds",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "百叶窗（按方向展开）",
    },
    "checker": {
        "ns": "p",
        "xml_tag": "checker",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "棋盘格（按方向展开）",
    },
    "circle": {
        "ns": "p",
        "xml_tag": "circle",
        "attrs": {},
        "dir_values": [],
        "description": "圆形展开（圆形扩散）",
    },
    "diamond": {
        "ns": "p",
        "xml_tag": "diamond",
        "attrs": {},
        "dir_values": [],
        "description": "菱形展开（菱形扩散）",
    },
    "plus": {
        "ns": "p",
        "xml_tag": "plus",
        "attrs": {},
        "dir_values": [],
        "description": "十字展开（十字扩散）",
    },
    "wedge": {
        "ns": "p",
        "xml_tag": "wedge",
        "attrs": {},
        "dir_values": [],
        "description": "楔形（双向楔形展开）",
    },
    "comb": {
        "ns": "p",
        "xml_tag": "comb",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "梳齿（按方向梳齿展开）",
    },
    "randomBar": {
        "ns": "p",
        "xml_tag": "randomBar",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "随机条（随机条纹展开）",
    },
    "strips": {
        "ns": "p",
        "xml_tag": "strips",
        "attrs": {"dir": "ld"},
        "dir_values": ["l", "r", "u", "d", "ld", "lu", "rd", "ru"],
        "description": "条带（按对角方向展开）",
    },

    # ---------- PowerPoint 2010+ 扩展 19 种 ----------
    "conveyor": {
        "ns": "p14",
        "xml_tag": "conveyor",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r"],
        "description": "传送带（左右传送切换）",
    },
    "doors": {
        "ns": "p14",
        "xml_tag": "doors",
        "attrs": {"dir": "horz"},
        "dir_values": ["horz", "vert"],
        "description": "开门（门式开合切换）",
    },
    "ferris": {
        "ns": "p14",
        "xml_tag": "ferris",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r"],
        "description": "摩天轮（摩天轮式旋转）",
    },
    "flip": {
        "ns": "p14",
        "xml_tag": "flip",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r"],
        "description": "翻转（页面翻转切换）",
    },
    "flythrough": {
        "ns": "p14",
        "xml_tag": "flythrough",
        "attrs": {"dir": "in"},
        "dir_values": ["in", "out"],
        "description": "飞行穿越（立体穿越效果）",
    },
    "gallery": {
        "ns": "p14",
        "xml_tag": "gallery",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r"],
        "description": "画廊（画廊滑动切换）",
    },
    "glitter": {
        "ns": "p14",
        "xml_tag": "glitter",
        "attrs": {"dir": "l", "pattern": "hexagon"},
        "dir_values": ["l", "r"],
        "description": "闪烁（闪光颗粒切换）",
    },
    "honeycomb": {
        "ns": "p14",
        "xml_tag": "honeycomb",
        "attrs": {},
        "dir_values": [],
        "description": "蜂窝（六边形蜂窝展开）",
    },
    "pan": {
        "ns": "p14",
        "xml_tag": "pan",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "平移（页面平移切换）",
    },
    "prism": {
        "ns": "p14",
        "xml_tag": "prism",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "棱镜（立体棱镜切换）",
    },
    "reveal": {
        "ns": "p14",
        "xml_tag": "reveal",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r"],
        "description": "显露（从一侧显露切换）",
    },
    "ripple": {
        "ns": "p14",
        "xml_tag": "ripple",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "波纹（水波纹扩散）",
    },
    "shred": {
        "ns": "p14",
        "xml_tag": "shred",
        "attrs": {"pattern": "strip"},
        "dir_values": [],
        "description": "碎纸（碎片化切换）",
    },
    "switch": {
        "ns": "p14",
        "xml_tag": "switch",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "切换（立体翻转切换）",
    },
    "vortex": {
        "ns": "p14",
        "xml_tag": "vortex",
        "attrs": {"dir": "l"},
        "dir_values": ["l", "r", "u", "d"],
        "description": "漩涡（漩涡式旋转切换）",
    },
    "warp": {
        "ns": "p14",
        "xml_tag": "warp",
        "attrs": {"dir": "in"},
        "dir_values": ["in", "out"],
        "description": "扭曲（空间扭曲切换）",
    },
    "window": {
        "ns": "p14",
        "xml_tag": "window",
        "attrs": {"dir": "horz"},
        "dir_values": ["horz", "vert"],
        "description": "窗户（百叶窗式开合）",
    },
    "flash": {
        "ns": "p14",
        "xml_tag": "flash",
        "attrs": {},
        "dir_values": [],
        "description": "闪光（白闪切换）",
    },
    "wheelReverse": {
        "ns": "p14",
        "xml_tag": "wheelReverse",
        "attrs": {"spokes": "1"},
        "dir_values": [],
        "description": "反向轮辐（反向辐射展开）",
    },
}


# ==================== 方向别名映射 ====================
# 用户友好方向名 → ECMA-376 方向值
DIR_ALIASES = {
    "from_left": "l",
    "from_right": "r",
    "from_top": "u",
    "from_bottom": "d",
    "left": "l",
    "right": "r",
    "top": "u",
    "up": "u",
    "bottom": "d",
    "down": "d",
    # 对角方向（strips 专用）
    "from_top_left": "ld",
    "from_top_right": "rd",
    "from_bottom_left": "lu",
    "from_bottom_right": "ru",
    # 通用 in/out（zoom/split/warp/flythrough 用）
    "in": "in",
    "out": "out",
    # orient 别名
    "horizontal": "horz",
    "vertical": "vert",
    "horz": "horz",
    "vert": "vert",
}


def _resolve_dir(dir_alias):
    """将方向别名解析为 ECMA-376 方向值"""
    if not dir_alias:
        return None
    return DIR_ALIASES.get(dir_alias, dir_alias)


def _build_transition_element(spec, speed_enum, duration_ms):
    """
    根据转场规范构建单个 <p:transition> 元素（不含 mc:AlternateContent 包裹）

    :param spec: 转场目录中的规范 dict
    :param speed_enum: 速度枚举值（slow/med/fast）
    :param duration_ms: 时长（毫秒），用于 p14:dur 属性
    :return: lxml Element
    """
    ns_prefix = spec["ns"]
    tag = spec["xml_tag"]
    default_attrs = dict(spec.get("attrs", {}))

    # 用户传入的 dir 覆盖默认值
    user_dir = default_attrs.pop("dir", None) if "dir" in default_attrs else None

    # 创建 p:transition 根元素
    p_transition = etree.Element(f"{{{NS_P}}}transition", nsmap={"p": NS_P})
    p_transition.set("spd", speed_enum)

    # 构建子元素
    if ns_prefix == "p":
        # ECMA-376 转场：直接作为 p:transition 子元素
        child_tag = f"{{{NS_P}}}{tag}"
        child = etree.SubElement(p_transition, child_tag)
        # 写入属性（合并默认属性 + dir）
        for k, v in default_attrs.items():
            child.set(k, v)
        if user_dir:
            child.set("dir", user_dir)
    else:
        # p14 转场：作为 p:transition 子元素，但带 p14 命名空间
        # 注意：p14 元素必须在 p:transition 内部
        child_tag = f"{{{NS_P14}}}{tag}"
        child = etree.SubElement(p_transition, child_tag)
        child.set(f"{{{NS_P14}}}dur", str(duration_ms))
        for k, v in default_attrs.items():
            # p14 元素的部分属性需带 p14 前缀，部分属性无前缀
            # 实践中：dir/orient 等使用无前缀（PowerPoint 兼容），dur 带 p14 前缀
            child.set(k, v)
        if user_dir:
            child.set("dir", user_dir)

    return p_transition


def _build_alternate_content(transition_elem, spec, speed_enum, duration_ms):
    """
    用 mc:AlternateContent 包裹 PowerPoint 2010+ 转场
    Choice：使用 p14 转场；Fallback：使用 fade 作为降级

    :param transition_elem: 已构建好的 p:transition 元素（含 p14 子元素）
    :return: mc:AlternateContent 元素
    """
    nsmap = {"mc": NS_MC, "p": NS_P, "p14": NS_P14}

    mc_ac = etree.Element(f"{{{NS_MC}}}AlternateContent", nsmap=nsmap)
    mc_choice = etree.SubElement(mc_ac, f"{{{NS_MC}}}Choice")
    mc_choice.set("Requires", "p14")

    # Choice 中放完整的 p:transition（含 p14 子元素）
    mc_choice.append(transition_elem)

    # Fallback 中放 fade 降级版本
    mc_fallback = etree.SubElement(mc_ac, f"{{{NS_MC}}}Fallback")
    fallback_transition = etree.Element(f"{{{NS_P}}}transition", nsmap={"p": NS_P})
    fallback_transition.set("spd", speed_enum)
    etree.SubElement(fallback_transition, f"{{{NS_P}}}fade")
    mc_fallback.append(fallback_transition)

    return mc_ac


def inject_transition(slide, transition_spec):
    """
    为 slide 注入转场效果（直接修改 slide XML）

    :param slide: python-pptx 的 Slide 对象
    :param transition_spec: 转场配置 dict
        - type: 转场类型（如 "fade"），见 TRANSITION_CATALOG
        - speed: 速度（slow/med/fast），默认 "med"
        - dir: 方向（可选，部分转场支持）
    :return: True 表示注入成功，False 表示失败

    示例：
        inject_transition(slide, {"type": "fade", "speed": "med"})
        inject_transition(slide, {"type": "push", "dir": "from_left", "speed": "slow"})
    """
    if not transition_spec or not isinstance(transition_spec, dict):
        return False

    trans_type = transition_spec.get("type", "fade")
    speed_key = transition_spec.get("speed", "med")
    user_dir_alias = transition_spec.get("dir")

    spec = TRANSITION_CATALOG.get(trans_type)
    if not spec:
        print(f"⚠️  不支持的转场类型: {trans_type}")
        return False

    # 解析 speed
    if speed_key not in SPEED_MAP:
        print(f"⚠️  不支持的速度值: {speed_key}，使用默认 med")
        speed_key = "med"
    speed_enum = SPEED_ENUM[speed_key]
    duration_ms = SPEED_MAP[speed_key]

    # 解析 dir（覆盖默认值）
    spec_copy = dict(spec)
    attrs_copy = dict(spec_copy.get("attrs", {}))
    if user_dir_alias and spec_copy.get("dir_values"):
        resolved = _resolve_dir(user_dir_alias)
        if resolved in spec_copy["dir_values"]:
            attrs_copy["dir"] = resolved
        else:
            print(f"⚠️  转场 {trans_type} 不支持方向 {user_dir_alias}（支持: {spec_copy['dir_values']}）")
    spec_copy["attrs"] = attrs_copy

    # 获取 slide 的根 XML 元素（<p:sld>）
    sld_elem = slide._element

    # 移除已有的 transition 或 AlternateContent（避免重复）
    for child in list(sld_elem):
        tag_local = etree.QName(child).localname
        if tag_local in ("transition", "AlternateContent"):
            # 仅移除含 transition 的 AlternateContent
            if tag_local == "AlternateContent":
                has_transition = any(
                    etree.QName(c).localname == "transition"
                    for sub in child.iter()
                    for c in [sub]
                )
                # 简化：只要 AlternateContent 内嵌套有 transition 即移除
                if any(etree.QName(c).localname == "transition" for c in child.iter()):
                    sld_elem.remove(child)
            else:
                sld_elem.remove(child)

    # 构建新的 transition 元素
    transition_elem = _build_transition_element(spec_copy, speed_enum, duration_ms)

    if spec_copy["ns"] == "p14":
        # PowerPoint 2010+ 转场：用 mc:AlternateContent 包裹
        mc_ac = _build_alternate_content(transition_elem, spec_copy, speed_enum, duration_ms)
        # 插入位置：p:transition 应在 p:cSld 之后、p:timing 之前
        # ECMA-376 顺序：cSld, clrMapOvr, transition, timing
        _insert_at_schema_position(sld_elem, mc_ac, "AlternateContent")
    else:
        # ECMA-376 转场：直接插入
        _insert_at_schema_position(sld_elem, transition_elem, "transition")

    return True


def _insert_at_schema_position(sld_elem, new_elem, elem_local_name):
    """
    按 ECMA-376 schema 顺序插入新元素到 p:sld 中
    p:sld 子元素顺序：cSld, clrMapOvr, transition, timing
    """
    # schema 顺序索引
    schema_order = ["cSld", "clrMapOvr", "transition", "timing", "AlternateContent"]
    new_idx = schema_order.index(elem_local_name) if elem_local_name in schema_order else len(schema_order)

    # 寻找插入位置
    insert_pos = len(sld_elem)  # 默认末尾
    for i, child in enumerate(sld_elem):
        child_local = etree.QName(child).localname
        if child_local in schema_order:
            child_idx = schema_order.index(child_local)
            if child_idx > new_idx:
                insert_pos = i
                break

    sld_elem.insert(insert_pos, new_elem)


def validate_transition(transition_spec, slide_num=None):
    """
    校验转场配置是否合法

    :param transition_spec: 转场配置 dict
    :param slide_num: 页码（用于警告信息）
    :return: 警告列表（空列表表示无警告）
    """
    warnings = []
    if not transition_spec or not isinstance(transition_spec, dict):
        warnings.append(f"页{slide_num or '?'}: 转场配置为空或非 dict")
        return warnings

    trans_type = transition_spec.get("type")
    if not trans_type:
        warnings.append(f"页{slide_num or '?'}: 转场配置缺少 type 字段")
        return warnings

    if trans_type not in TRANSITION_CATALOG:
        warnings.append(f"页{slide_num or '?'}: 不支持的转场类型 {trans_type}")
        return warnings

    spec = TRANSITION_CATALOG[trans_type]
    speed = transition_spec.get("speed", "med")
    if speed not in SPEED_MAP:
        warnings.append(f"页{slide_num or '?'}: 速度值 {speed} 不合法（应为 slow/med/fast）")

    user_dir = transition_spec.get("dir")
    if user_dir and spec.get("dir_values"):
        resolved = _resolve_dir(user_dir)
        if resolved not in spec["dir_values"]:
            warnings.append(
                f"页{slide_num or '?'}: 转场 {trans_type} 不支持方向 {user_dir}"
                f"（支持: {spec['dir_values']}）"
            )
    elif user_dir and not spec.get("dir_values"):
        warnings.append(
            f"页{slide_num or '?'}: 转场 {trans_type} 不支持方向参数（将忽略）"
        )

    return warnings


def list_transitions():
    """列出所有可用转场效果（按命名空间分组）"""
    ecma = [(k, v) for k, v in TRANSITION_CATALOG.items() if v["ns"] == "p"]
    p14 = [(k, v) for k, v in TRANSITION_CATALOG.items() if v["ns"] == "p14"]
    return {
        "ecma376": [{"type": k, "description": v["description"],
                     "supports_dir": bool(v["dir_values"])} for k, v in ecma],
        "p14_extension": [{"type": k, "description": v["description"],
                           "supports_dir": bool(v["dir_values"])} for k, v in p14],
        "total": len(TRANSITION_CATALOG),
    }


if __name__ == "__main__":
    # 简单自测：列出转场目录
    import json
    catalog = list_transitions()
    print(f"转场目录总数: {catalog['total']}")
    print(f"ECMA-376 核心: {len(catalog['ecma376'])} 种")
    print(f"PowerPoint 2010+ 扩展: {len(catalog['p14_extension'])} 种")
    print("\nECMA-376 核心：")
    for t in catalog["ecma376"]:
        dir_mark = " [支持方向]" if t["supports_dir"] else ""
        print(f"  {t['type']:15s} - {t['description']}{dir_mark}")
    print("\nPowerPoint 2010+ 扩展：")
    for t in catalog["p14_extension"]:
        dir_mark = " [支持方向]" if t["supports_dir"] else ""
        print(f"  {t['type']:15s} - {t['description']}{dir_mark}")
