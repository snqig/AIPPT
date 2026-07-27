"""
AIPPT 六层防御体系 · 校验引擎
功能：
  1. JSON Schema 机器化校验（需求参数 / 大纲）
  2. 模板槽位精准匹配校验（业务级格式对齐）
  3. 运行时兜底自动修复（容错机制）
  4. 标准化错误反馈体系（错误码 + 修正建议）

校验层级：
  Layer 1: 模型端自检（SKILL 指令引导，不在此代码内）
  Layer 2: JSON Schema 机器化校验（validate_schema）
  Layer 3: 分步校验流程（validate_requirement / validate_outline / validate_template_match）
  Layer 4: 模板槽位匹配校验（validate_template_match）
  Layer 5: 运行时兜底修复（auto_fix_outline）
  Layer 6: 标准化错误反馈（ValidationResult / ValidationError）
"""
import json
import re
from pathlib import Path
from typing import Any, Optional

from jsonschema import Draft7Validator, ValidationError as JsonSchemaValidationError

from aippt.logger import logger


# ==================== 错误码定义（第六层）====================
# F0xx: 基础格式错误（JSON 解析失败、类型不匹配）
# F1xx: 字段规则错误（字段缺失、枚举非法、长度超限）
# S0xx: 结构逻辑错误（页数不匹配、ID 不连续）
# T0xx: 模板匹配错误（槽位不匹配、页面类型不兼容）
# A0xx: 动画转场错误（名称非法、字段层级错、by_bullet 类型错）

SCHEMAS_DIR = Path(__file__).parent.parent / "schemas"

# 场景枚举（与 SCENE_SCHEMAS 保持同步）
SCENE_ENUM = [
    "工作总结", "年终总结", "工作汇报", "工作计划", "述职报告",
    "个人简历", "自我介绍", "开题报告", "公司简介", "职业规划", "安全教育",
]

# 页面类型枚举（12 类）
PAGE_TYPE_ENUM = [
    "cover", "catalog", "divider", "numbered_list", "kpi", "timeline",
    "two_column", "skill_percent", "preset_titles", "chart", "table", "ending",
]

# 转场效果枚举（与 ppt_transitions.TRANSITION_CATALOG 保持同步，38 种）
TRANSITION_ENUM = [
    # ECMA-376 核心 19 种
    "fade", "cut", "push", "cover", "pull", "wipe", "dissolve", "split",
    "zoom", "wheel", "blinds", "checker", "circle", "diamond", "plus",
    "wedge", "comb", "randomBar", "strips",
    # PowerPoint 2010+ 扩展 19 种
    "conveyor", "doors", "ferris", "flip", "flythrough", "gallery", "glitter",
    "honeycomb", "pan", "prism", "reveal", "ripple", "shred", "switch",
    "vortex", "warp", "window", "flash", "wheelReverse",
]

# 入场动画枚举（与 ppt_animations.ANIMATION_CATALOG entry 保持同步）
ANIM_ENTRY_ENUM = ["fade", "fly_in", "zoom", "wipe", "slide_in", "bounce", "spin"]
# 退场动画枚举
ANIM_EXIT_ENUM = ["fade_out", "fly_out", "zoom_out", "slide_out"]
# 强调动画枚举
ANIM_EMPHASIS_ENUM = ["pulse", "spin", "shake", "grow_shrink", "color_blast"]

# 不支持 by_bullet 的页面类型（封面/分隔/KPI/结尾等单元素页）
BY_BULLET_FORBIDDEN_PAGES = {"cover", "divider", "kpi", "ending", "chart", "table"}
# 商务场景不推荐使用的高动态转场特效
HIGH_DYNAMIC_TRANSITIONS = {"vortex", "fling", "switch", "ferris", "flythrough", "glitter", "shred"}

# 各页面类型的数组长度限制
PAGE_ARRAY_LIMITS = {
    "catalog": {"items": (2, 8)},
    "numbered_list": {"items": (3, 5)},
    "preset_titles": {"items": (2, 8)},
    "kpi": {"kpi_items": (2, 4)},
    "timeline": {"timeline_items": (3, 6)},
    "two_column": {"left_items": (2, 6), "right_items": (2, 6)},
    "skill_percent": {"skills": (2, 6)},
    "table": {"headers": (2, 6), "rows": (1, 10)},
}


class ValidationError:
    """标准化校验错误项"""

    def __init__(self, code: str, level: str, path: str, message: str, suggestion: str = ""):
        self.code = code        # 错误码，如 F102
        self.level = level      # error / warning
        self.path = path        # 错误位置，如 pages[2].page_type
        self.message = message  # 错误描述
        self.suggestion = suggestion  # 修正建议

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "level": self.level,
            "path": self.path,
            "message": self.message,
            "suggestion": self.suggestion,
        }


class ValidationResult:
    """标准化校验结果"""

    def __init__(self):
        self.errors: list[ValidationError] = []
        self.warnings: list[ValidationError] = []
        self.fixed: list[str] = []  # 自动修复记录

    @property
    def is_valid(self) -> bool:
        """是否通过（无 error 级别错误）"""
        return len(self.errors) == 0

    def add_error(self, code: str, path: str, message: str, suggestion: str = ""):
        self.errors.append(ValidationError(code, "error", path, message, suggestion))

    def add_warning(self, code: str, path: str, message: str, suggestion: str = ""):
        self.warnings.append(ValidationError(code, "warning", path, message, suggestion))

    def add_fix(self, description: str):
        self.fixed.append(description)

    def to_dict(self) -> dict[str, Any]:
        return {
            "validate_pass": self.is_valid,
            "error_count": len(self.errors),
            "warning_count": len(self.warnings),
            "fixed_count": len(self.fixed),
            "errors": [e.to_dict() for e in self.errors],
            "warnings": [w.to_dict() for w in self.warnings],
            "fixes": self.fixed,
        }

    def merge(self, other: "ValidationResult"):
        self.errors.extend(other.errors)
        self.warnings.extend(other.warnings)
        self.fixed.extend(other.fixed)


# ==================== Layer 2: JSON Schema 机器化校验 ====================

def _load_schema(schema_name: str) -> dict[str, Any]:
    """加载内置 Schema 文件"""
    schema_path = SCHEMAS_DIR / f"{schema_name}.schema.json"
    if not schema_path.exists():
        raise FileNotFoundError(f"Schema 文件不存在: {schema_path}")
    with open(schema_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _jsonschema_path_to_str(error: JsonSchemaValidationError) -> str:
    """将 jsonschema 的 deque path 转为可读字符串，如 pages[2].page_type"""
    parts = []
    for part in error.absolute_path:
        if isinstance(part, int):
            parts.append(f"[{part}]")
        else:
            if parts:
                parts.append(f".{part}")
            else:
                parts.append(str(part))
    return "".join(parts) or "(root)"


def _map_jsonschema_error(error: JsonSchemaValidationError) -> ValidationError:
    """将 jsonschema 原始错误映射为标准化 ValidationError"""
    path = _jsonschema_path_to_str(error)
    msg = error.message

    # 根据错误信息分类错误码
    if "is not of type" in msg or "is not valid" in msg:
        code = "F001"
    elif "is a required property" in msg or "is required" in msg:
        code = "F100"
    elif "is not one of" in msg or "is not a valid choice" in msg:
        code = "F102"
        suggestion = f"请从枚举列表中选择合法值（详见 Schema 定义）"
        return ValidationError(code, "error", path, f"枚举值非法: {msg}", suggestion)
    elif "is longer than" in msg or "too long" in msg.lower():
        code = "F103"
    elif "is less than" in msg or "is greater than" in msg or "too short" in msg.lower():
        code = "F104"
    elif "is not valid under any of the given schemas" in msg:
        code = "F105"
    else:
        code = "F000"

    return ValidationError(code, "error", path, f"Schema 校验失败: {msg}", "")


def validate_schema(data: Any, schema_name: str) -> ValidationResult:
    """
    执行 JSON Schema 校验

    :param data: 待校验数据
    :param schema_name: Schema 名称（不含 .schema.json 后缀）
    :return: ValidationResult
    """
    result = ValidationResult()
    try:
        schema = _load_schema(schema_name)
        validator = Draft7Validator(schema)
        for error in validator.iter_errors(data):
            result.errors.append(_map_jsonschema_error(error))
    except FileNotFoundError as e:
        result.add_error("F002", "(root)", str(e), "请检查 schemas/ 目录是否存在对应 Schema 文件")
    except json.JSONDecodeError as e:
        result.add_error("F003", "(root)", f"Schema 文件 JSON 解析失败: {e}", "请检查 Schema 文件语法")
    except Exception as e:
        result.add_error("F004", "(root)", f"校验引擎异常: {e}", "请联系开发者排查")
    return result


# ==================== Layer 3: 分步校验流程 ====================

def validate_requirement(params: dict[str, Any]) -> ValidationResult:
    """Step 1 后校验：需求参数合法性"""
    result = validate_schema(params, "requirement_params")

    # 业务规则补充校验
    scene = params.get("scene")
    if scene and scene not in SCENE_ENUM:
        result.add_error(
            "F102", "scene",
            f"场景枚举值非法: {scene}",
            f"请从以下列表选择: {', '.join(SCENE_ENUM)}",
        )

    page_count = params.get("page_count")
    if page_count is not None:
        if not isinstance(page_count, int) or page_count < 4 or page_count > 60:
            result.add_error(
                "F104", "page_count",
                f"页数范围非法: {page_count}",
                "page_count 必须为 4-60 的正整数",
            )

    return result


def validate_outline(outline: dict[str, Any]) -> ValidationResult:
    """Step 2 后校验：大纲结构完整性"""
    result = validate_schema(outline, "outline")

    pages = outline.get("pages", [])
    total_pages = outline.get("total_pages")

    # 仅当使用 pages 数组格式时才校验 pages 相关规则（cover/sections/end 格式跳过）
    if not pages:
        return result

    # S001: total_pages 与 pages 数组长度一致性
    if total_pages is not None and len(pages) != total_pages:
        result.add_error(
            "S001", "total_pages",
            f"total_pages({total_pages}) 与 pages 数组长度({len(pages)}) 不一致",
            f"请将 total_pages 修正为 {len(pages)}，或调整 pages 数组",
        )

    # S002: page_id 连续性
    page_ids = [p.get("page_id") for p in pages if "page_id" in p]
    expected_ids = list(range(1, len(pages) + 1))
    if page_ids != expected_ids:
        result.add_warning(
            "S002", "pages",
            f"page_id 不连续或非从1开始: {page_ids}",
            "建议 page_id 从1开始连续递增，可调用 auto_fix 自动修复",
        )

    # S003: 首页必须为 cover
    if pages and pages[0].get("page_type") != "cover":
        result.add_error(
            "S003", "pages[0].page_type",
            f"首页类型应为 cover，实际为 {pages[0].get('page_type')}",
            "请将首页 page_type 改为 cover",
        )

    # S004: 末页必须为 ending
    if pages and pages[-1].get("page_type") != "ending":
        result.add_warning(
            "S004", f"pages[{len(pages)-1}].page_type",
            f"末页类型建议为 ending，实际为 {pages[-1].get('page_type')}",
            "建议末页 page_type 改为 ending",
        )

    # F102: 页面类型枚举校验（已在 Schema 中，此处补充业务建议）
    for i, page in enumerate(pages):
        ptype = page.get("page_type", "")
        if ptype and ptype not in PAGE_TYPE_ENUM:
            result.add_error(
                "F102", f"pages[{i}].page_type",
                f"页面类型枚举值非法: {ptype}",
                f"请从以下列表选择: {', '.join(PAGE_TYPE_ENUM)}",
            )

    # A0xx: 动画转场字段校验（merge validate_animations 结果）
    result.merge(validate_animations(outline))

    return result


# ==================== Layer 3.5: 动画转场字段校验（A0xx 错误码）====================

def validate_animations(outline: dict[str, Any]) -> ValidationResult:
    """
    校验 outline 中单页动画转场字段格式合规性

    覆盖错误码：
      A001: 转场效果名称非法
      A002: 动画效果名称非法
      A003: animations 字段类型错误（应为对象，实为字符串等）
      A004: by_bullet 类型错误（应为布尔值）
      A005: 页面类型不支持逐段动画（warning，可自动关闭）
      A006: 不推荐使用高动态特效（warning）

    :param outline: 大纲数据
    :return: ValidationResult
    """
    result = ValidationResult()
    pages = outline.get("pages", [])
    if not pages:
        return result

    for i, page in enumerate(pages):
        page_path = f"pages[{i}]"
        ptype = page.get("page_type", "")

        # A001: 单页 transition 字段校验
        transition = page.get("transition")
        if transition is not None:
            if not isinstance(transition, str):
                result.add_error(
                    "A001", f"{page_path}.transition",
                    f"转场效果类型错误: 应为 string，实际为 {type(transition).__name__}",
                    "transition 必须为字符串枚举值或 null",
                )
            elif transition != "none" and transition not in TRANSITION_ENUM:
                result.add_error(
                    "A001", f"{page_path}.transition",
                    f"转场效果名称非法: {transition}",
                    f"请从枚举列表选择: {', '.join(TRANSITION_ENUM[:12])}... 共 {len(TRANSITION_ENUM)} 种",
                )
            elif transition in HIGH_DYNAMIC_TRANSITIONS:
                result.add_warning(
                    "A006", f"{page_path}.transition",
                    f"不推荐使用高动态特效: {transition}",
                    "商务场景建议使用 fade/push/wipe 等简约转场",
                )

        # A002/A003/A004: 单页 animations 字段校验
        animations = page.get("animations")
        if animations is None:
            continue

        # A003: animations 必须为对象结构
        if not isinstance(animations, dict):
            result.add_error(
                "A003", f"{page_path}.animations",
                f"animations 字段类型错误: 应为对象，实际为 {type(animations).__name__}",
                'animations 必须为对象结构，如 {"entry": "fly_in", "by_bullet": true}',
            )
            continue

        # A002: entry 动画枚举校验
        entry = animations.get("entry")
        if entry is not None:
            if not isinstance(entry, str):
                result.add_error(
                    "A002", f"{page_path}.animations.entry",
                    f"入场动画类型错误: 应为 string，实际为 {type(entry).__name__}",
                    "entry 必须为字符串枚举值或 null",
                )
            elif entry not in ANIM_ENTRY_ENUM:
                result.add_error(
                    "A002", f"{page_path}.animations.entry",
                    f"入场动画名称非法: {entry}",
                    f"入场动画合法值: {', '.join(ANIM_ENTRY_ENUM)}",
                )

        # A002: exit 动画枚举校验
        exit_anim = animations.get("exit")
        if exit_anim is not None:
            if not isinstance(exit_anim, str):
                result.add_error(
                    "A002", f"{page_path}.animations.exit",
                    f"退场动画类型错误: 应为 string，实际为 {type(exit_anim).__name__}",
                    "exit 必须为字符串枚举值或 null",
                )
            elif exit_anim not in ANIM_EXIT_ENUM:
                result.add_error(
                    "A002", f"{page_path}.animations.exit",
                    f"退场动画名称非法: {exit_anim}",
                    f"退场动画合法值: {', '.join(ANIM_EXIT_ENUM)}",
                )

        # A002: emphasis 动画枚举校验
        emphasis = animations.get("emphasis")
        if emphasis is not None:
            if not isinstance(emphasis, str):
                result.add_error(
                    "A002", f"{page_path}.animations.emphasis",
                    f"强调动画类型错误: 应为 string，实际为 {type(emphasis).__name__}",
                    "emphasis 必须为字符串枚举值或 null",
                )
            elif emphasis not in ANIM_EMPHASIS_ENUM:
                result.add_error(
                    "A002", f"{page_path}.animations.emphasis",
                    f"强调动画名称非法: {emphasis}",
                    f"强调动画合法值: {', '.join(ANIM_EMPHASIS_ENUM)}",
                )

        # A004: by_bullet 类型校验
        by_bullet = animations.get("by_bullet")
        if by_bullet is not None:
            if not isinstance(by_bullet, bool):
                result.add_error(
                    "A004", f"{page_path}.animations.by_bullet",
                    f"by_bullet 类型错误: 应为 boolean，实际为 {type(by_bullet).__name__} ({by_bullet!r})",
                    "by_bullet 必须是布尔值 true/false，不能写字符串",
                )
            elif by_bullet is True and ptype in BY_BULLET_FORBIDDEN_PAGES:
                # A005: 页面类型不支持逐段动画（warning，可自动关闭）
                result.add_warning(
                    "A005", f"{page_path}.animations.by_bullet",
                    f"页面类型 {ptype} 不支持逐段动画，已标记需自动关闭",
                    f"by_bullet=true 仅适用于 numbered_list/catalog/timeline/preset_titles 等多段落页面",
                )

    return result


# ==================== Layer 4: 模板槽位匹配校验 ====================

def validate_template_match(outline: dict[str, Any], meta: dict[str, Any]) -> ValidationResult:
    """
    Step 3/4 后校验：大纲与模板元数据的业务级匹配

    :param outline: 大纲数据
    :param meta: 模板 meta.json
    :return: ValidationResult
    """
    result = ValidationResult()

    pages = outline.get("pages", [])
    template_pages = meta.get("total_pages", 0)

    # T001: 总页数匹配度（仅对 pages 数组格式校验；cover/sections/end 格式由渲染器自动适配模板页数）
    if pages:
        total_pages = outline.get("total_pages", len(pages))
        removable = set(meta.get("removable_pages", []))
        effective_template_pages = template_pages - len(removable)

        if effective_template_pages > 0:
            diff = abs(total_pages - effective_template_pages)
            if diff > 5:
                result.add_error(
                    "T001", "total_pages",
                    f"大纲页数({total_pages}) 与模板有效页数({effective_template_pages}) 偏差过大({diff}页)",
                    f"建议大纲页数控制在 {effective_template_pages}±5 页",
                )
            elif diff > 2:
                result.add_warning(
                    "T001", "total_pages",
                    f"大纲页数({total_pages}) 与模板有效页数({effective_template_pages}) 轻微偏差({diff}页)",
                    "可自动适配，但建议页数更接近",
                )

    # T002: 场景匹配
    outline_scene = outline.get("scene", "")
    template_category = meta.get("category", "")
    if outline_scene and template_category and outline_scene != template_category:
        result.add_warning(
            "T002", "scene",
            f"大纲场景({outline_scene}) 与模板分类({template_category}) 不一致",
            "场景不匹配可能导致内容风格不协调，建议使用同分类模板",
        )

    # T003: 逐页槽位数量匹配（仅对 pages 数组格式）
    for page in pages:
        ptype = page.get("page_type", "")
        if ptype in ("cover", "catalog", "divider", "ending"):
            continue

        if ptype == "kpi":
            items_count = len(page.get("kpi_items", []))
            page_limits = PAGE_ARRAY_LIMITS.get("kpi", {}).get("kpi_items", (2, 4))
        elif ptype == "timeline":
            items_count = len(page.get("timeline_items", []))
            page_limits = PAGE_ARRAY_LIMITS.get("timeline", {}).get("timeline_items", (3, 6))
        elif ptype in ("numbered_list", "catalog", "preset_titles"):
            items_count = len(page.get("items", []))
            page_limits = PAGE_ARRAY_LIMITS.get(ptype, {}).get("items", (2, 8))
        elif ptype == "skill_percent":
            items_count = len(page.get("skills", []))
            page_limits = PAGE_ARRAY_LIMITS.get("skill_percent", {}).get("skills", (2, 6))
        else:
            continue

        min_limit, max_limit = page_limits
        if items_count > max_limit:
            result.add_warning(
                "T003", f"pages[{page.get('page_id', '?')}].{ptype}",
                f"{ptype} 页条目数({items_count}) 超出建议上限({max_limit})",
                f"建议精简为 {max_limit} 条以内，可调用 auto_fix 自动截断",
            )
        elif items_count < min_limit:
            result.add_warning(
                "T003", f"pages[{page.get('page_id', '?')}].{ptype}",
                f"{ptype} 页条目数({items_count}) 低于建议下限({min_limit})",
                f"建议补充至 {min_limit} 条以上",
            )

    return result


# ==================== Layer 5: 运行时兜底自动修复 ====================

def auto_fix_outline(outline: dict[str, Any]) -> tuple[dict[str, Any], ValidationResult]:
    """
    对非原则性格式问题自动修复

    可修复：
      - page_id 不连续/重复/从0开始 → 重排为 1,2,3...
      - 要点数量超出限制 → 截断保留前 N 条
      - 枚举值大小写不敏感 → 标准化
      - 多余非关键字段 → 忽略

    不可修复（返回 error 由模型修正）：
      - 核心必填字段缺失
      - 页面类型不在枚举列表内
      - JSON 语法错误

    :param outline: 原始大纲
    :return: (修复后大纲, 校验结果)
    """
    result = ValidationResult()
    fixed = dict(outline)
    pages = fixed.get("pages", [])

    # Fix 1: page_id 重排
    if pages:
        old_ids = [p.get("page_id") for p in pages]
        expected = list(range(1, len(pages) + 1))
        if old_ids != expected:
            for i, page in enumerate(pages):
                page["page_id"] = i + 1
            result.add_fix(f"page_id 已重排: {old_ids} → {expected}")

    # Fix 2: total_pages 同步
    if "total_pages" in fixed and fixed["total_pages"] != len(pages):
        old_tp = fixed["total_pages"]
        fixed["total_pages"] = len(pages)
        result.add_fix(f"total_pages 已修正: {old_tp} → {len(pages)}")

    # Fix 3: 页面类型枚举大小写标准化
    for i, page in enumerate(pages):
        ptype = page.get("page_type", "")
        if isinstance(ptype, str):
            ptype_lower = ptype.lower()
            # 映射常见的大小写/下划线变体
            type_aliases = {
                "cover_page": "cover", "封面": "cover",
                "catalog_page": "catalog", "目录": "catalog",
                "kpi_card": "kpi", "kpi_page": "kpi",
                "timeline_page": "timeline",
                "ending_page": "ending", "结尾": "ending", "end": "ending",
            }
            if ptype_lower in type_aliases:
                page["page_type"] = type_aliases[ptype_lower]
                result.add_fix(f"pages[{i}].page_type 标准化: {ptype} → {page['page_type']}")
            elif ptype != ptype_lower and ptype_lower in PAGE_TYPE_ENUM:
                page["page_type"] = ptype_lower
                result.add_fix(f"pages[{i}].page_type 大小写修正: {ptype} → {ptype_lower}")

    # Fix 4: 数组超限截断
    for i, page in enumerate(pages):
        ptype = page.get("page_type", "")
        limits = PAGE_ARRAY_LIMITS.get(ptype, {})
        for field, (_, max_limit) in limits.items():
            if field in page and isinstance(page[field], list) and len(page[field]) > max_limit:
                old_len = len(page[field])
                page[field] = page[field][:max_limit]
                result.add_fix(
                    f"pages[{i}].{field} 截断: {old_len} → {max_limit} 条"
                )

    # Fix 5: 移除空值字段
    for i, page in enumerate(pages):
        empty_keys = [k for k, v in page.items() if v is None or (isinstance(v, str) and v.strip() == "")]
        for k in empty_keys:
            if k not in ("title", "page_type", "page_id"):  # 必填字段不移除
                del page[k]
                result.add_fix(f"pages[{i}].{k} 移除空值字段")

    # Fix 6: 动画转场字段兜底修复（A003/A004/A005）
    for i, page in enumerate(pages):
        ptype = page.get("page_type", "")
        page_path = f"pages[{i}]"

        # Fix 6.1: animations 为字符串 → 降级为对象结构
        animations = page.get("animations")
        if isinstance(animations, str):
            # 字符串形式的 animations 视为 entry 动画名称
            anim_name = animations
            page["animations"] = {"entry": anim_name} if anim_name in ANIM_ENTRY_ENUM else {}
            result.add_fix(
                f"{page_path}.animations 字段类型修复: 字符串 → 对象结构"
            )
            animations = page["animations"]

        # Fix 6.2: by_bullet 字符串 "true"/"false" → 布尔值
        if isinstance(animations, dict):
            by_bullet = animations.get("by_bullet")
            if isinstance(by_bullet, str):
                normalized = by_bullet.lower() in ("true", "1", "yes")
                animations["by_bullet"] = normalized
                result.add_fix(
                    f"{page_path}.animations.by_bullet 类型修复: '{by_bullet}' → {normalized}"
                )
                by_bullet = normalized

            # Fix 6.3: A005 - 不支持 by_bullet 的页面类型自动关闭
            if by_bullet is True and ptype in BY_BULLET_FORBIDDEN_PAGES:
                animations["by_bullet"] = False
                result.add_fix(
                    f"{page_path}.animations.by_bullet 自动关闭: 页面类型 {ptype} 不支持逐段动画"
                )

            # Fix 6.4: 非法动画名称降级为 auto（移除字段，由渲染器继承全局）
            for sub_key, valid_enum in (
                ("entry", ANIM_ENTRY_ENUM),
                ("exit", ANIM_EXIT_ENUM),
                ("emphasis", ANIM_EMPHASIS_ENUM),
            ):
                val = animations.get(sub_key)
                if isinstance(val, str) and val not in valid_enum:
                    del animations[sub_key]
                    result.add_fix(
                        f"{page_path}.animations.{sub_key} 非法值 '{val}' 已移除，将继承全局 auto"
                    )

        # Fix 6.5: 非法 transition 名称降级（移除字段）
        transition = page.get("transition")
        if isinstance(transition, str) and transition != "none" and transition not in TRANSITION_ENUM:
            del page["transition"]
            result.add_fix(
                f"{page_path}.transition 非法值 '{transition}' 已移除，将继承全局 auto"
            )

    # 仅当原大纲已有 pages 字段或 pages 非空时才写回，避免给 cover/sections/end 格式强加空 pages
    if "pages" in outline or pages:
        fixed["pages"] = pages
    return fixed, result


# ==================== 全链路综合校验入口 ====================

def validate_all(
    outline: dict[str, Any],
    params: Optional[dict[str, Any]] = None,
    meta: Optional[dict[str, Any]] = None,
    auto_fix: bool = True,
) -> tuple[dict[str, Any], ValidationResult]:
    """
    全链路综合校验（Layer 2-5 合并）

    :param outline: 大纲数据
    :param params: 需求参数（可选，Step 1 产出）
    :param meta: 模板元数据（可选，Step 3 产出）
    :param auto_fix: 是否执行自动修复
    :return: (可能修复后的大纲, 综合校验结果)
    """
    result = ValidationResult()

    # Layer 3.1: 需求参数校验（可选）
    if params:
        result.merge(validate_requirement(params))

    # Layer 5: 先尝试自动修复
    working_outline = outline
    if auto_fix:
        working_outline, fix_result = auto_fix_outline(outline)
        result.merge(fix_result)

    # Layer 3.2: 大纲结构校验
    result.merge(validate_outline(working_outline))

    # Layer 4: 模板槽位匹配校验（可选）
    if meta:
        result.merge(validate_template_match(working_outline, meta))

    return working_outline, result
