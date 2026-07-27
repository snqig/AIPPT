"""
场景适配层 - 阶段三核心模块
功能：将业务字段按场景 Schema 自动映射到模板槽位，用户无需关心页码与槽位名
依赖：ppt_renderer.py, meta 文件
"""
import json
from pathlib import Path
from typing import Any, Optional

from ppt_renderer import PptRenderer
from aippt.config import EN_DECORATIVE_KEYWORDS, CN_DECORATIVE_PHRASES
from aippt.constants import is_placeholder, is_chart_decoration_text, PLACEHOLDER_KEYWORDS
from aippt.logger import logger


# ==================== 场景 Schema 定义 ====================
# 每个场景定义业务字段到槽位的映射规则
# 章节内容（items）按章节页范围顺序填充 title/desc 槽位

SCENE_SCHEMAS = {
    "工作总结": {
        "name": "工作总结场景",
        "cover_fields": {
            "title": "工作总结标题",
            "subtitle": "副标题",
            "reporter": "汇报人",
            "period": "汇报周期",
            "department": "所属部门"
        },
        "chapter_sections": [
            {"key": "work_content", "name": "工作内容汇报", "desc": "本期主要工作内容"},
            {"key": "project_progress", "name": "项目完成进度", "desc": "项目推进情况"},
            {"key": "issues", "name": "问题不足讨论", "desc": "存在的问题与不足"},
            {"key": "plan", "name": "下步工作计划", "desc": "下步工作规划"}
        ],
        "end_fields": {
            "thanks": "结束致谢语"
        }
    },
    "个人简历": {
        "name": "个人简历场景",
        "cover_fields": {
            "title": "竞聘岗位/职位名称",
            "reporter": "姓名",
            "period": "竞聘时间"
        },
        "chapter_sections": [
            {"key": "basic_info", "name": "个人信息", "desc": "基本情况/学习经历/工作经历/荣誉证书"},
            {"key": "competency", "name": "胜任能力", "desc": "核心能力/技能熟练度"},
            {"key": "job_awareness", "name": "岗位认知", "desc": "岗位职责/关键能力"},
            {"key": "career_plan", "name": "职业规划", "desc": "职业目标/年度计划"}
        ],
        "end_fields": {
            "thanks": "致谢语"
        }
    },
    "自我介绍": {
        "name": "自我介绍场景",
        "cover_fields": {
            "title": "自我介绍主题",
            "reporter": "姓名",
            "period": "时间"
        },
        "chapter_sections": [
            {"key": "basic_info", "name": "个人基本信息", "desc": "基本情况/学习经历"},
            {"key": "education", "name": "教育背景", "desc": "教育经历/课程学习"},
            {"key": "internship", "name": "实习经历", "desc": "实习工作经历"},
            {"key": "research", "name": "科研经历", "desc": "科研项目/学术研究"},
            {"key": "awards", "name": "获奖情况", "desc": "获奖荣誉/证书"},
            {"key": "self_assessment", "name": "自我评价", "desc": "自我评价/自学能力"},
            {"key": "future_plan", "name": "读研展望", "desc": "读研规划/未来展望"}
        ],
        "end_fields": {
            "thanks": "致谢语"
        }
    },
    "年终总结": {
        "name": "年终总结场景",
        "cover_fields": {
            "title": "年终总结标题",
            "reporter": "汇报人",
            "period": "年度",
            "department": "所属部门"
        },
        "chapter_sections": [
            {"key": "annual_review", "name": "年度工作回顾", "desc": "全年工作概述"},
            {"key": "achievements", "name": "主要业绩成果", "desc": "关键业绩与亮点"},
            {"key": "experience", "name": "经验与不足", "desc": "总结经验与不足"},
            {"key": "next_year", "name": "新年工作规划", "desc": "下年度工作计划"}
        ],
        "end_fields": {
            "thanks": "结束致谢语"
        }
    },
    "工作汇报": {
        "name": "工作汇报场景",
        "cover_fields": {
            "title": "工作汇报标题",
            "reporter": "汇报人",
            "period": "汇报时间",
            "department": "所属部门"
        },
        "chapter_sections": [
            {"key": "progress", "name": "工作进展汇报", "desc": "当前工作进展"},
            {"key": "results", "name": "阶段成果展示", "desc": "阶段性成果"},
            {"key": "challenges", "name": "困难与挑战", "desc": "面临的困难"},
            {"key": "next_steps", "name": "后续工作安排", "desc": "下一步工作"}
        ],
        "end_fields": {
            "thanks": "结束致谢语"
        }
    },
    "工作计划": {
        "name": "工作计划场景",
        "cover_fields": {
            "title": "工作计划标题",
            "reporter": "编制人",
            "period": "计划周期",
            "department": "所属部门"
        },
        "chapter_sections": [
            {"key": "objectives", "name": "工作目标", "desc": "总体目标与分解"},
            {"key": "tasks", "name": "重点任务", "desc": "关键任务清单"},
            {"key": "schedule", "name": "进度安排", "desc": "时间节点与里程碑"},
            {"key": "resources", "name": "资源与保障", "desc": "资源需求与保障措施"}
        ],
        "end_fields": {
            "thanks": "结束致谢语"
        }
    },
    "述职报告": {
        "name": "述职报告场景",
        "cover_fields": {
            "title": "述职报告标题",
            "reporter": "述职人",
            "period": "述职时间",
            "position": "现任职务"
        },
        "chapter_sections": [
            {"key": "duty_performance", "name": "履职情况", "desc": "岗位职责履行情况"},
            {"key": "achievements", "name": "工作业绩", "desc": "主要工作业绩"},
            {"key": "problems", "name": "问题与不足", "desc": "存在的问题"},
            {"key": "improvement", "name": "改进方向", "desc": "改进措施与方向"}
        ],
        "end_fields": {
            "thanks": "结束致谢语"
        }
    },
    "开题报告": {
        "name": "开题报告场景",
        "cover_fields": {
            "title": "开题报告标题",
            "reporter": "答辩人",
            "period": "答辩时间",
            "advisor": "指导老师"
        },
        "chapter_sections": [
            {"key": "background", "name": "选题背景与意义", "desc": "研究背景与意义"},
            {"key": "literature", "name": "国内外研究现状", "desc": "文献综述"},
            {"key": "methodology", "name": "研究方法与思路", "desc": "研究方法与思路"},
            {"key": "conclusion", "name": "论文结论与总结", "desc": "预期结论与总结"}
        ],
        "end_fields": {
            "thanks": "结束致谢语"
        }
    },
    "公司简介": {
        "name": "公司简介场景",
        "cover_fields": {
            "title": "公司名称",
            "period": "成立时间",
            "slogan": "公司口号"
        },
        "chapter_sections": [
            {"key": "overview", "name": "公司概况", "desc": "公司基本情况"},
            {"key": "honors", "name": "企业荣誉", "desc": "获得的荣誉与资质"},
            {"key": "leadership", "name": "领导团队", "desc": "核心管理团队介绍"},
            {"key": "business", "name": "产品与服务", "desc": "主营业务介绍"},
            {"key": "main_products", "name": "主打产品", "desc": "核心产品亮点"},
            {"key": "marketing", "name": "市场与营销", "desc": "市场布局与营销策略"},
            {"key": "market_analysis", "name": "市场分析", "desc": "行业市场分析"},
            {"key": "development", "name": "发展与计划", "desc": "发展规划与计划"},
            {"key": "future", "name": "前景与未来", "desc": "未来前景展望"}
        ],
        "end_fields": {
            "thanks": "结束致谢语"
        }
    },
    "职业规划": {
        "name": "职业规划场景",
        "cover_fields": {
            "title": "职业规划标题",
            "reporter": "姓名",
            "period": "规划周期"
        },
        "chapter_sections": [
            {"key": "self_analysis", "name": "自我分析", "desc": "个人优势与劣势"},
            {"key": "career_goal", "name": "职业目标", "desc": "短期与长期目标"},
            {"key": "action_plan", "name": "行动计划", "desc": "具体行动步骤"},
            {"key": "risk_assessment", "name": "风险评估", "desc": "风险与应对"}
        ],
        "end_fields": {
            "thanks": "结束致谢语"
        }
    },
    "安全教育": {
        "name": "安全教育场景",
        "cover_fields": {
            "title": "安全教育标题",
            "subtitle": "副标题",
            "reporter": "主讲人",
            "period": "培训日期"
        },
        "chapter_sections": [
            {"key": "overview", "name": "安全概述", "desc": "安全意义与目标"},
            {"key": "hazards", "name": "危险源识别", "desc": "风险辨识与分类"},
            {"key": "measures", "name": "安全措施", "desc": "防护与管控措施"},
            {"key": "training", "name": "安全培训", "desc": "培训内容与要求"},
            {"key": "emergency", "name": "应急预案", "desc": "应急响应与处置"}
        ],
        "end_fields": {
            "thanks": "结束致谢语"
        }
    }
}


class SceneAdapter:
    """场景适配器：业务字段 → 模板槽位"""

    def __init__(self, templates_root: str = "models") -> None:
        self.templates_root = Path(templates_root)
        self.index: dict[str, Any] = self._load_index()
        self.current_meta: dict[str, Any] = {}

    def _load_index(self) -> dict[str, Any]:
        index_path = self.templates_root / "templates_index.json"
        if not index_path.exists():
            return {"total": 0, "categories": {}, "templates": []}
        with open(index_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def list_scenes(self) -> list[dict[str, str]]:
        return [{"category": k, "name": v["name"]} for k, v in SCENE_SCHEMAS.items()]

    def get_scene_schema(self, scene: str) -> dict[str, Any]:
        """
        获取场景的Schema定义（字段结构说明）
        :param scene: 场景名
        :return: 场景Schema字典
        """
        schema = SCENE_SCHEMAS.get(scene)
        if not schema:
            raise ValueError(f"不支持的场景: {scene}")
        return {
            "scene": scene,
            "name": schema["name"],
            "cover_fields": schema["cover_fields"],
            "chapter_sections": [
                {"key": s["key"], "name": s["name"], "desc": s["desc"]}
                for s in schema["chapter_sections"]
            ],
            "end_fields": schema["end_fields"],
            "input_format": {
                "cover": "对象，字段对应 cover_fields",
                "sections": "对象，key 为 chapter_sections[].key，值为 [{title, desc}] 数组",
                "end": "对象，字段对应 end_fields"
            }
        }

    def get_template_detail(self, template_id):
        """
        获取模板详情（meta 摘要 + 章节结构 + 槽位统计）
        :param template_id: 模板ID
        :return: 模板详情字典
        """
        meta, meta_path = self.get_template_meta(template_id=template_id)
        total_slots = sum(len(v) for v in meta.get("page_slots", {}).values())
        chapters = meta.get("chapters", [])
        return {
            "template_id": template_id,
            "category": meta.get("category"),
            "total_pages": meta.get("total_pages"),
            "removable_pages": meta.get("removable_pages", []),
            "chapter_count": len(chapters),
            "total_slots": total_slots,
            "chapters": [
                {"key": c.get("key"), "name": c.get("name", ""),
                 "page": c.get("page"), "start_page": c.get("start_page"),
                 "end_page": c.get("end_page")}
                for c in chapters
            ],
            "page_slot_summary": {
                page: len(slots) for page, slots in meta.get("page_slots", {}).items()
            }
        }

    def validate_business_data(self, scene, business_data):
        """
        校验业务数据是否符合场景Schema
        :param scene: 场景名
        :param business_data: 业务数据
        :return: (is_valid, issues列表)
        """
        schema = SCENE_SCHEMAS.get(scene)
        if not schema:
            return False, [f"不支持的场景: {scene}"]

        issues = []
        # 校验cover字段
        cover = business_data.get("cover", {})
        if not isinstance(cover, dict):
            issues.append("cover 应为对象")
        else:
            for field in schema["cover_fields"]:
                if field not in cover:
                    issues.append(f"cover 缺少推荐字段: {field}（{schema['cover_fields'][field]}）")

        # 校验sections
        sections = business_data.get("sections", {})
        if not isinstance(sections, dict):
            issues.append("sections 应为对象")
        else:
            for sec_def in schema["chapter_sections"]:
                key = sec_def["key"]
                if key not in sections:
                    issues.append(f"sections 缺少章节: {key}（{sec_def['name']}）")
                elif not isinstance(sections[key], list):
                    issues.append(f"sections.{key} 应为数组")
                else:
                    for i, item in enumerate(sections[key]):
                        if not isinstance(item, dict):
                            issues.append(f"sections.{key}[{i}] 应为对象")
                        elif "title" not in item:
                            issues.append(f"sections.{key}[{i}] 缺少 title 字段")

        # 校验end
        end = business_data.get("end", {})
        if not isinstance(end, dict):
            issues.append("end 应为对象")

        is_valid = len(issues) == 0
        return is_valid, issues

    def list_templates(self, category=None, style_tag=None, min_pages=None, max_pages=None,
                       color_scheme=None, industry=None):
        """
        按条件筛选模板
        :param category: 分类名（如"工作总结"）
        :param style_tag: 风格标签（如"商务"）
        :param min_pages: 最小页数
        :param max_pages: 最大页数
        :param color_scheme: 色系筛选（如"蓝色系"）。缺字段的模板视为不匹配，仅过滤时跳过
        :param industry: 行业筛选（如"金融"）。匹配 industry 数组中任一元素即可
        :return: 模板列表
        """
        results = []
        for t in self.index.get("templates", []):
            if category and t["category"] != category:
                continue
            if style_tag and style_tag not in t.get("style_tags", []):
                continue
            if min_pages and t["total_pages"] < min_pages:
                continue
            if max_pages and t["total_pages"] > max_pages:
                continue
            # 色系筛选：缺字段时该模板跳过（仅过滤时跳过，不影响其他条件下的列出）
            if color_scheme is not None:
                if "color_scheme" not in t:
                    continue
                if t["color_scheme"] != color_scheme:
                    continue
            # 行业筛选：industry 为数组，命中任一元素即可
            if industry is not None:
                t_industries = t.get("industry", [])
                if not isinstance(t_industries, list) or industry not in t_industries:
                    continue
            results.append(t)
        return results

    def get_template_meta(self, template_id=None, path=None):
        """加载模板 meta"""
        if path:
            meta_path = self.templates_root / path
        else:
            for t in self.index.get("templates", []):
                if t["template_id"] == template_id:
                    meta_path = self.templates_root / t["path"]
                    break
            else:
                raise ValueError(f"模板不存在: {template_id}")
        with open(meta_path, 'r', encoding='utf-8') as f:
            return json.load(f), str(meta_path)

    def get_template_pptx(self, meta):
        """根据 meta 获取对应的 pptx 路径"""
        meta_path = Path(meta.get("_meta_file_path", ""))
        # meta 文件名格式为 xxx.meta.json，对应模板为 xxx.pptx
        name = meta_path.name
        if name.endswith(".meta.json"):
            base = name[:-10]  # strip ".meta.json"
            pptx_name = base if base.endswith(".pptx") else base + ".pptx"
        else:
            pptx_name = meta_path.stem + ".pptx"
        return str(meta_path.parent / pptx_name)

    def adapt(self, scene, business_data, meta):
        """
        将业务字段适配为渲染所需的 slot_data
        :param scene: 场景名（如"工作总结"）
        :param business_data: 业务字段数据
        :param meta: 模板 meta
        :return: {"页码": {"槽位名": "值"}}
        """
        schema = SCENE_SCHEMAS.get(scene)
        if not schema:
            raise ValueError(f"不支持的场景: {scene}")

        slot_data = {}
        page_slots = meta.get("page_slots", {})
        chapters = meta.get("chapters", [])
        # 暴露 meta 给 _fill_chapters / _fill_orphan_pages，用于读取 page_meta（chart/table 标志）
        self.current_meta = meta

        # 1. 填充封面
        self._fill_cover(slot_data, chapters, page_slots, business_data, schema)

        # 2. 填充目录页
        self._fill_catalog(slot_data, chapters, page_slots, schema)

        # 3. 填充各章节
        # 给 chapter_sections 加 idx 字段
        for i, sec in enumerate(schema["chapter_sections"]):
            sec["idx"] = i + 1
        self._fill_chapters(slot_data, chapters, page_slots, business_data, schema)

        # 3.5 兜底：填充未归属章节的页面（meta 章节识别偏差导致）
        self._fill_orphan_pages(slot_data, chapters, page_slots, business_data, schema)

        # 4. 填充结束页
        self._fill_end(slot_data, chapters, page_slots, business_data, schema)

        return slot_data

    def _is_decorative_text(self, text: str) -> bool:
        if not text:
            return False
        cleaned = text.replace(" ", "").replace(".", "").replace(",", "")
        if len(cleaned) > 3 and cleaned.isalpha() and cleaned.isupper() and cleaned.isascii():
            return True
        stripped = text.strip()
        if len(stripped) <= 2 and stripped and not any(c.isalnum() for c in stripped):
            return True
        upper_text = text.upper()
        for kw in EN_DECORATIVE_KEYWORDS:
            if kw in upper_text:
                return True
        for kw in CN_DECORATIVE_PHRASES:
            if text.strip() == kw or kw in text:
                return True
        return False

    def _is_placeholder_text(self, text: str) -> bool:
        return is_placeholder(text)

    def _is_chart_decoration(self, text: str) -> bool:
        return is_chart_decoration_text(text)

    def _is_chapter_title(self, text, section_def):
        """判断是否为章节标题文本（与 section name 匹配）"""
        if not text or not section_def:
            return False
        section_name = section_def.get("name", "")
        return text.strip() == section_name.strip()

    def _fill_cover(self, slot_data, chapters, page_slots, business_data, schema):
        """填充封面页字段"""
        cover_chapter = next((c for c in chapters if c.get("key") == "cover"), None)
        if not cover_chapter:
            return
        cover_page = str(cover_chapter["page"])
        cover_slots = page_slots.get(cover_page, [])
        cover_input = business_data.get("cover", {})

        # 按封面槽位顺序填充业务字段
        cover_field_keys = list(schema["cover_fields"].keys())
        used_fields = set()
        field_idx = 0
        for slot_info in cover_slots:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            # 优先精确匹配槽位名（reporter/period 等）
            if slot_name in cover_input:
                slot_data.setdefault(cover_page, {})[slot_name] = cover_input[slot_name]
                used_fields.add(slot_name)
            # 跳过纯装饰性英文标题（如 WORK REPORT / RESUME / PPT TEMPLATE）
            elif slot_name.startswith("title") and self._is_decorative_text(match_text):
                continue
            # 按顺序填充 title 系列（跳过已精确匹配的字段）
            elif slot_name.startswith("title"):
                while field_idx < len(cover_field_keys) and cover_field_keys[field_idx] in used_fields:
                    field_idx += 1
                if field_idx < len(cover_field_keys):
                    field_key = cover_field_keys[field_idx]
                    if field_key in cover_input:
                        slot_data.setdefault(cover_page, {})[slot_name] = cover_input[field_key]
                    field_idx += 1

    def _fill_catalog(self, slot_data, chapters, page_slots, schema):
        """填充目录页：章节名 + 序号"""
        catalog_chapter = next((c for c in chapters if c.get("key") == "catalog"), None)
        if not catalog_chapter:
            return
        catalog_page = str(catalog_chapter["page"])
        catalog_slots = page_slots.get(catalog_page, [])
        if not catalog_slots:
            return

        chapter_sections = schema["chapter_sections"]
        # 目录页：number 槽位填序号，title 槽位填章节名
        num_idx = 0
        title_idx = 0
        for slot_info in catalog_slots:
            slot_name = slot_info.get("slot", "")
            if slot_name.startswith("number"):
                num_idx += 1
                slot_data.setdefault(catalog_page, {})[slot_name] = f"{num_idx:02d}"
            elif slot_name.startswith("title") and title_idx < len(chapter_sections):
                slot_data.setdefault(catalog_page, {})[slot_name] = chapter_sections[title_idx]["name"]
                title_idx += 1

    def _detect_page_pattern(self, page_slot_list, page_meta=None, page_idx=0):
        """
        识别页面槽位模式
        返回: 'cover' | 'divider' | 'numbered_list' | 'timeline' | 'preset_titles'
              | 'skill_percent' | 'kpi' | 'two_column' | 'chart' | 'table' | 'content'

        :param page_slot_list: 该页槽位列表
        :param page_meta: 可选，meta 中的 page_meta[page_num]，含 has_chart/has_table 等标志
        :param page_idx: 该页 0 基页码，默认 0；首页强制识别为 cover（向后兼容）
        """
        # 首页强制识别为封面
        if page_idx == 0:
            return "cover"
        # 优先识别复合页面：含 chart/table 形状的页面
        if page_meta:
            if page_meta.get("has_chart"):
                return "chart"
            if page_meta.get("has_table"):
                return "table"

        slot_names = [s.get("slot", "") for s in page_slot_list]
        slot_types = set()
        for name in slot_names:
            if name.startswith("number"):
                slot_types.add("number")
            elif name.startswith("year"):
                slot_types.add("year")
            elif name.startswith("percent"):
                slot_types.add("percent")
            elif name.startswith("title"):
                slot_types.add("title")
            elif name.startswith("desc"):
                slot_types.add("desc")
            elif name.startswith("item"):
                slot_types.add("item")

        match_texts = [s.get("match_text", "") for s in page_slot_list]

        # 章节分隔页：含 PART.0N 或"第N章"关键词
        has_part_kw = any("PART" in t.upper() or "第" in t or "章" in t for t in match_texts)
        if "title" in slot_types and has_part_kw and len(page_slot_list) <= 6:
            return "divider"

        # 百分比页：含 percent 槽位
        if "percent" in slot_types:
            return "skill_percent"

        # 年份/时间轴页：含 year 槽位或多个年份格式文本
        if "year" in slot_types:
            return "timeline"
        # year_like 仅统计非 title 槽位的年份格式文本
        # （避免将 title 中的"2013某某省"、"1280,000"等误判为时间轴）
        non_title_texts = [s.get("match_text", "") for s in page_slot_list
                           if not s.get("slot", "").startswith("title")]
        year_like = sum(1 for t in non_title_texts
                        if len(t) >= 4 and (t[:4].isdigit() or "20xx" in t.lower() or "20XX" in t)
                        or "20xx" in t.lower() or ".5—" in t or ".12—" in t)
        if year_like >= 2:
            return "timeline"

        # 数字列表页：number 槽位 + title 槽位，无 desc
        # 但当 title 数量 >= 4 时，视为预设标题页（单个 number 是装饰序号）
        title_count = sum(1 for n in slot_names if n.startswith("title"))
        number_count = sum(1 for n in slot_names if n.startswith("number"))
        desc_count = sum(1 for n in slot_names if n.startswith("desc"))

        # KPI页面：多个短title（>=3）配对数值，无desc，无number
        # 特征：title数量>=3且每个title的match_text较短（<15字符），无desc配对
        if title_count >= 3 and desc_count == 0 and "number" not in slot_types and "percent" not in slot_types:
            # 检查是否是短标题（KPI卡片特征）
            short_titles = sum(1 for s in page_slot_list
                              if s.get("slot", "").startswith("title") and len(s.get("match_text", "")) < 15)
            if short_titles >= 3:
                return "kpi"

        # 双栏页面：title数量>=2且desc数量>=2，且title/desc交替排列
        # 特征：左右两列对称的title+desc结构
        if title_count >= 2 and desc_count >= 2 and title_count == desc_count:
            return "two_column"

        if "number" in slot_types and "title" in slot_types and "desc" not in slot_types:
            if title_count >= 4 and number_count <= 1:
                return "preset_titles"
            return "numbered_list"

        # 预设标题列表页：多个 title_N 槽位（>=4），无 desc
        if title_count >= 4 and "desc" not in slot_types and "number" not in slot_types:
            return "preset_titles"

        return "content"

    def _fill_orphan_pages(self, slot_data, chapters, page_slots, business_data, schema):
        """
        兜底填充未归属章节的页面
        策略：cover 之后、第一个 chapter 之前的页面，归入第一个 section
        """
        cover_page = next((c.get("page") for c in chapters if c.get("key") == "cover"), None)
        first_chapter = next((c for c in chapters if c.get("key", "").startswith("chapter_")), None)
        end_page = next((c.get("page") for c in chapters if c.get("key") == "end"), None)

        if not first_chapter or not cover_page:
            return

        first_chapter_start = first_chapter.get("start_page", 1)
        chapter_sections = schema["chapter_sections"]
        if not chapter_sections:
            return

        # 第一个 section 的数据
        first_section_def = chapter_sections[0]
        first_section_key = first_section_def["key"]
        sections_data = business_data.get("sections", {})
        section_data = sections_data.get(first_section_key, [])
        if not section_data:
            return

        # 遍历 cover 之后到第一个 chapter 之前的页面
        item_idx = 0
        catalog_page = next((c.get("page") for c in chapters if c.get("key") == "catalog"), None)
        page_meta = self.current_meta.get("page_meta", {}) if hasattr(self, "current_meta") else {}
        for page_num in range(cover_page + 1, first_chapter_start):
            page_str = str(page_num)
            # 跳过目录页
            if catalog_page and page_num == catalog_page:
                continue

            page_slot_list = page_slots.get(page_str, [])
            if not page_slot_list:
                continue

            # 已填充的页面跳过
            if page_str in slot_data:
                continue

            pattern = self._detect_page_pattern(page_slot_list, page_meta.get(page_str), page_idx=page_num - 1)
            page_input = {}

            if pattern == "divider":
                page_input = self._fill_divider_page(page_slot_list, first_section_def, section_data)
            elif pattern == "timeline":
                page_input, item_idx = self._fill_timeline_page(page_slot_list, section_data, item_idx)
            elif pattern == "preset_titles":
                page_input, item_idx = self._fill_preset_titles_page(page_slot_list, section_data, item_idx)
            elif pattern == "chart":
                page_input = self._fill_chart_page(page_slot_list, first_section_def, section_data)
            elif pattern == "table":
                page_input = self._fill_table_page(page_slot_list, first_section_def, section_data)
            else:
                page_input, item_idx = self._fill_content_page(page_slot_list, section_data, item_idx)

            if page_input:
                slot_data[page_str] = page_input

    def _fill_chapters(self, slot_data, chapters, page_slots, business_data, schema):
        """填充各章节内容"""
        chapter_chapters = [c for c in chapters if c.get("key", "").startswith("chapter_")]
        chapter_sections = schema["chapter_sections"]
        sections_data = business_data.get("sections", {})
        page_meta = self.current_meta.get("page_meta", {}) if hasattr(self, "current_meta") else {}

        for idx, chapter in enumerate(chapter_chapters):
            if idx >= len(chapter_sections):
                break
            section_def = chapter_sections[idx]
            section_key = section_def["key"]
            section_data = sections_data.get(section_key, [])
            if not section_data:
                continue

            start_page = chapter.get("start_page")
            end_page = chapter.get("end_page")
            if not start_page or not end_page:
                continue

            # 遍历章节内每一页，按页面模式分派填充
            item_idx = 0
            for page_num in range(start_page, end_page + 1):
                page_str = str(page_num)
                page_slot_list = page_slots.get(page_str, [])
                if not page_slot_list:
                    continue

                pattern = self._detect_page_pattern(page_slot_list, page_meta.get(page_str), page_idx=page_num - 1)
                page_input = {}

                if pattern == "divider":
                    page_input = self._fill_divider_page(page_slot_list, section_def, section_data)
                elif pattern == "numbered_list":
                    page_input, item_idx = self._fill_numbered_list_page(page_slot_list, section_data, item_idx)
                elif pattern == "timeline":
                    page_input, item_idx = self._fill_timeline_page(page_slot_list, section_data, item_idx)
                elif pattern == "preset_titles":
                    page_input, item_idx = self._fill_preset_titles_page(page_slot_list, section_data, item_idx)
                elif pattern == "skill_percent":
                    page_input = self._fill_skill_percent_page(page_slot_list, section_data)
                elif pattern == "kpi":
                    page_input, item_idx = self._fill_kpi_page(page_slot_list, section_data, item_idx)
                elif pattern == "two_column":
                    page_input, item_idx = self._fill_two_column_page(page_slot_list, section_data, item_idx, section_def)
                elif pattern == "chart":
                    page_input = self._fill_chart_page(page_slot_list, section_def, section_data)
                elif pattern == "table":
                    page_input = self._fill_table_page(page_slot_list, section_def, section_data)
                else:
                    page_input, item_idx = self._fill_content_page(page_slot_list, section_data, item_idx, section_def)

                if page_input:
                    slot_data[page_str] = page_input

    def _fill_divider_page(self, page_slot_list, section_def, section_data):
        """填充章节分隔页：PART.0N + 章节名 + 序号"""
        page_input = {}
        section_name = section_data[0].get("section_title", section_def["name"]) if section_data else section_def["name"]
        section_idx = section_def.get("idx", 1)

        title_count = 0
        for slot_info in page_slot_list:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            # 装饰性文本跳过：清空避免残留
            if self._is_decorative_text(match_text):
                page_input[slot_name] = ""
                continue
            if slot_name.startswith("number"):
                page_input[slot_name] = f"{section_idx:02d}"
            elif slot_name.startswith("title"):
                if "PART" in match_text.upper():
                    # PART 标识
                    page_input[slot_name] = f"PART.{section_idx:02d}"
                else:
                    # 章节名（title 和 title_2 都填章节名）
                    page_input[slot_name] = section_name
                    title_count += 1
        return page_input

    def _fill_numbered_list_page(self, page_slot_list, section_data, item_idx):
        """填充数字列表页：number=01/02/03 + title=名称 + desc=描述"""
        page_input = {}
        num_counter = 1
        current_item = None
        last_slot_was_title = False
        for slot_info in page_slot_list:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            if self._is_decorative_text(match_text):
                continue
            if self._is_chart_decoration(match_text):
                if match_text.strip() == "添加文字":
                    page_input[slot_name] = ""
                continue
            if slot_name.startswith("number"):
                page_input[slot_name] = f"{num_counter:02d}"
                num_counter += 1
                # number 后通常对应新 item
                if item_idx < len(section_data):
                    current_item = section_data[item_idx]
                    item_idx += 1
                    last_slot_was_title = False
            elif slot_name.startswith("title"):
                if current_item:
                    page_input[slot_name] = current_item.get("title", "")
                    last_slot_was_title = True
                elif item_idx < len(section_data):
                    current_item = section_data[item_idx]
                    page_input[slot_name] = current_item.get("title", "")
                    item_idx += 1
                    last_slot_was_title = True
                else:
                    page_input[slot_name] = ""
                    last_slot_was_title = False
            elif slot_name.startswith("desc"):
                if last_slot_was_title and current_item:
                    page_input[slot_name] = current_item.get("desc", "")
                    last_slot_was_title = False
                elif current_item:
                    page_input[slot_name] = current_item.get("desc", "")
                else:
                    page_input[slot_name] = ""
        return page_input, item_idx

    def _fill_timeline_page(self, page_slot_list, section_data, item_idx):
        """填充时间轴页：year/title 交替，按时间线展开 items"""
        page_input = {}
        year_slots = [s for s in page_slot_list if s.get("slot", "").startswith("year")]
        content_slots = [s for s in page_slot_list if s.get("slot", "").startswith(("title", "desc"))]

        # 为每个时间点配对一个 item
        items_for_page = []
        for _ in year_slots:
            if item_idx < len(section_data):
                items_for_page.append(section_data[item_idx])
                item_idx += 1
            else:
                items_for_page.append(None)

        # 填充 year 槽位：优先 time 字段，其次 title
        for i, slot_info in enumerate(year_slots):
            slot_name = slot_info.get("slot", "")
            if i < len(items_for_page) and items_for_page[i]:
                item = items_for_page[i]
                time_val = item.get("time", item.get("title", ""))
                page_input[slot_name] = time_val

        # 填充 content 槽位：title/desc 配对同一个 item
        item_ptr = 0
        last_slot_was_title = False
        for slot_info in content_slots:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            # 跳过装饰性/占位文本标题：清空避免残留
            if self._is_decorative_text(match_text):
                page_input[slot_name] = ""
                continue
            if slot_name.startswith("title") and item_ptr < len(items_for_page):
                item = items_for_page[item_ptr]
                if item:
                    page_input[slot_name] = item.get("title", "")
                else:
                    page_input[slot_name] = ""
                last_slot_was_title = True
            elif slot_name.startswith("title"):
                page_input[slot_name] = ""
                last_slot_was_title = False
            elif slot_name.startswith("desc"):
                if last_slot_was_title and item_ptr < len(items_for_page):
                    # desc 配对当前 title 的 item
                    item = items_for_page[item_ptr]
                    if item:
                        page_input[slot_name] = item.get("desc", "")
                    else:
                        page_input[slot_name] = ""
                    item_ptr += 1
                    last_slot_was_title = False
                elif item_ptr < len(items_for_page):
                    # 连续 desc，取下一个 item
                    item = items_for_page[item_ptr]
                    if item:
                        page_input[slot_name] = item.get("desc", "")
                    else:
                        page_input[slot_name] = ""
                    item_ptr += 1
                else:
                    page_input[slot_name] = ""

        return page_input, item_idx

    def _fill_preset_titles_page(self, page_slot_list, section_data, item_idx):
        """
        填充预设标题列表页：多个并列 title_N，依次填 items 的 title
        规则：title 后紧跟的 desc 配对同一 item；连续 desc（无新 title）取下一个 item
        """
        page_input = {}
        last_title_item_idx = -1
        last_slot_was_title = False
        num_counter = 1
        for slot_info in page_slot_list:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            if self._is_decorative_text(match_text):
                page_input[slot_name] = ""
                continue
            if self._is_chart_decoration(match_text):
                if match_text.strip() == "添加文字":
                    page_input[slot_name] = ""
                continue
            if slot_name == "number" or slot_name.startswith("number_"):
                page_input[slot_name] = f"{num_counter:02d}"
                num_counter += 1
            elif slot_name.startswith("title"):
                if item_idx < len(section_data):
                    page_input[slot_name] = section_data[item_idx].get("title", "")
                    last_title_item_idx = item_idx
                    item_idx += 1
                    last_slot_was_title = True
                else:
                    page_input[slot_name] = ""
                    last_slot_was_title = False
            elif slot_name.startswith("desc"):
                if last_slot_was_title and last_title_item_idx >= 0:
                    # 配对上一个 title 的 item
                    page_input[slot_name] = section_data[last_title_item_idx].get("desc", "")
                    last_slot_was_title = False
                elif item_idx < len(section_data):
                    # 连续 desc，取下一个 item
                    page_input[slot_name] = section_data[item_idx].get("desc", "")
                    item_idx += 1
                    last_slot_was_title = False
                else:
                    page_input[slot_name] = ""
        return page_input, item_idx

    def _fill_kpi_page(self, page_slot_list, section_data, item_idx):
        """
        填充KPI卡片页：多个短title展示关键指标
        规则：每个title槽位取一个item的title，无desc配对
        - 数据耗尽时清空占位文本
        """
        page_input = {}
        for slot_info in page_slot_list:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            # 装饰文本清空
            if self._is_decorative_text(match_text):
                page_input[slot_name] = ""
                continue
            if self._is_chart_decoration(match_text):
                page_input[slot_name] = ""
                continue
            if slot_name.startswith("title"):
                if item_idx < len(section_data):
                    page_input[slot_name] = section_data[item_idx].get("title", "")
                    item_idx += 1
                else:
                    page_input[slot_name] = ""
            elif slot_name.startswith("number"):
                # KPI页面中的number保留原值或清空
                page_input[slot_name] = ""
        return page_input, item_idx

    def _fill_chart_page(self, page_slot_list, section_def, section_data):
        """
        填充图表页：仅填充标题类槽位，图表数据由模板自带或由 renderer 单独处理
        策略：
          - title 槽位填 section 名称或首个 item 标题
          - desc/number 等其他槽位清空（避免占位文本残留）
          - 不修改 chart 形状本身（保留模板原有图表样式与数据）
        """
        page_input = {}
        title_filled = False
        for slot_info in page_slot_list:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            if self._is_decorative_text(match_text):
                page_input[slot_name] = ""
                continue
            if self._is_chart_decoration(match_text):
                page_input[slot_name] = ""
                continue
            if slot_name.startswith("title"):
                if not title_filled:
                    # 优先使用 section_data 首项标题，回退到 section 名称
                    if section_data:
                        page_input[slot_name] = section_data[0].get("title", "") or section_def.get("name", "")
                    else:
                        page_input[slot_name] = section_def.get("name", "")
                    title_filled = True
                else:
                    # 其余 title 槽位填 item 标题或清空
                    page_input[slot_name] = section_data[0].get("title", "") if section_data else ""
            else:
                # 非标题槽位（desc/number/year 等）清空，避免占位文本残留
                page_input[slot_name] = ""
        return page_input

    def _fill_table_page(self, page_slot_list, section_def, section_data):
        """
        填充表格页：仅填充标题类槽位，表格数据由模板自带或后续扩展
        策略：与 chart 页类似，title 填 section/item 标题，其他槽位清空
        """
        page_input = {}
        title_filled = False
        for slot_info in page_slot_list:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            if self._is_decorative_text(match_text):
                page_input[slot_name] = ""
                continue
            if self._is_chart_decoration(match_text):
                page_input[slot_name] = ""
                continue
            if slot_name.startswith("title"):
                if not title_filled:
                    if section_data:
                        page_input[slot_name] = section_data[0].get("title", "") or section_def.get("name", "")
                    else:
                        page_input[slot_name] = section_def.get("name", "")
                    title_filled = True
                else:
                    page_input[slot_name] = section_data[0].get("title", "") if section_data else ""
            else:
                page_input[slot_name] = ""
        return page_input

    def _fill_two_column_page(self, page_slot_list, section_data, item_idx, section_def=None):
        """
        填充双栏对比页：左右两列对称的title+desc结构
        规则：title/desc交替填充，左列先填，右列后填
        - 数据耗尽时清空占位文本
        """
        page_input = {}
        current_item = None
        last_slot_was_title = False
        for slot_info in page_slot_list:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            # 装饰/章节标题清空
            if self._is_decorative_text(match_text):
                page_input[slot_name] = ""
                continue
            if section_def and self._is_chapter_title(match_text, section_def):
                page_input[slot_name] = ""
                continue
            if self._is_chart_decoration(match_text):
                page_input[slot_name] = ""
                continue
            if slot_name.startswith("title"):
                if item_idx < len(section_data):
                    current_item = section_data[item_idx]
                    page_input[slot_name] = current_item.get("title", "")
                    item_idx += 1
                    last_slot_was_title = True
                else:
                    page_input[slot_name] = ""
                    last_slot_was_title = False
            elif slot_name.startswith("desc"):
                if last_slot_was_title and current_item:
                    page_input[slot_name] = current_item.get("desc", "")
                    last_slot_was_title = False
                elif item_idx < len(section_data):
                    current_item = section_data[item_idx]
                    page_input[slot_name] = current_item.get("desc", "")
                    item_idx += 1
                    last_slot_was_title = False
                else:
                    page_input[slot_name] = ""
            elif slot_name.startswith("item"):
                if item_idx < len(section_data):
                    current_item = section_data[item_idx]
                    page_input[slot_name] = current_item.get("desc", "")
                    item_idx += 1
                else:
                    page_input[slot_name] = ""
        return page_input, item_idx

    def _fill_skill_percent_page(self, page_slot_list, section_data):
        """填充技能百分比页：title=技能名 + percent=数值 + desc=描述"""
        page_input = {}
        # 按原始顺序遍历，支持 title/desc/percent 配对
        skills = []
        for item in section_data:
            if "percent" in item or "level" in item:
                skills.append(item)
            else:
                skills.append({"title": item.get("title", ""), "percent": 80,
                               "desc": item.get("desc", "")})

        skill_idx = 0
        last_was_title = False
        for slot_info in page_slot_list:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            if self._is_decorative_text(match_text):
                page_input[slot_name] = ""
                continue
            if self._is_chart_decoration(match_text):
                if match_text.strip() == "添加文字":
                    page_input[slot_name] = ""
                continue
            if slot_name.startswith("percent"):
                if skill_idx < len(skills) and skill_idx > 0:
                    pct = skills[skill_idx - 1].get("percent", 80)
                    page_input[slot_name] = f"{pct}%"
                elif skills:
                    page_input[slot_name] = f"{skills[0].get('percent', 80)}%"
            elif slot_name.startswith("title"):
                if skill_idx < len(skills):
                    page_input[slot_name] = skills[skill_idx].get("title", skills[skill_idx].get("name", ""))
                    skill_idx += 1
                    last_was_title = True
                else:
                    page_input[slot_name] = ""
            elif slot_name.startswith("desc"):
                if last_was_title and skill_idx > 0 and skill_idx <= len(skills):
                    page_input[slot_name] = skills[skill_idx - 1].get("desc", "")
                    last_was_title = False
                elif skill_idx < len(skills):
                    page_input[slot_name] = skills[skill_idx].get("desc", "")
                    skill_idx += 1
                else:
                    page_input[slot_name] = ""
            elif slot_name.startswith("number"):
                # number 槽位保留原值（序号或统计数据）
                page_input[slot_name] = match_text.strip()
        return page_input

    def _fill_content_page(self, page_slot_list, section_data, item_idx, section_def=None):
        """
        填充标准内容页：title/desc 配对，跳过装饰性/章节标题
        规则：
        - title 后紧跟的 desc 配对同一 item
        - 连续 desc（无新 title）取下一个 item
        - title 被装饰文本跳过时，后续 desc 独立取下一个 item
        - number 槽位自动填充序号
        - 跳过图表装饰文本（坐标轴数字、单位标注）
        - 数据耗尽时清空占位提示文本，避免残留
        """
        page_input = {}
        current_item = None
        last_slot_was_title = False
        num_counter = 1
        for slot_info in page_slot_list:
            slot_name = slot_info.get("slot", "")
            match_text = slot_info.get("match_text", "")
            # 跳过装饰性文本（BUSINESS 等）：清空避免残留
            if self._is_decorative_text(match_text):
                page_input[slot_name] = ""
                continue
            # 跳过章节标题（与 section name 匹配）：清空避免残留
            if section_def and self._is_chapter_title(match_text, section_def):
                page_input[slot_name] = ""
                continue
            # 跳过图表装饰文本
            if self._is_chart_decoration(match_text):
                if match_text.strip() == "添加文字":
                    page_input[slot_name] = ""
                continue
            if slot_name == "number" or slot_name.startswith("number_"):
                page_input[slot_name] = f"{num_counter:02d}"
                num_counter += 1
            elif slot_name == "reporter":
                # reporter 槽位（如"请输入姓名 CEO"）：填入当前 item 的 title（人名）
                if item_idx < len(section_data):
                    current_item = section_data[item_idx]
                    page_input[slot_name] = current_item.get("title", "")
                    item_idx += 1
                    last_slot_was_title = True
                elif current_item:
                    page_input[slot_name] = current_item.get("title", "")
                    last_slot_was_title = True
                else:
                    # 数据耗尽：清空占位文本
                    page_input[slot_name] = ""
            elif slot_name == "title" or slot_name.startswith("title_"):
                # 占位提示文本（请在此添加文字说明等）同样需要被实际内容替换
                if item_idx < len(section_data):
                    current_item = section_data[item_idx]
                    page_input[slot_name] = current_item.get("title", "")
                    item_idx += 1
                    last_slot_was_title = True
                else:
                    # 数据耗尽：清空占位文本（避免残留"添加小标题"等）
                    page_input[slot_name] = ""
                    last_slot_was_title = False
            elif slot_name.startswith("desc"):
                # 占位 desc 需要被替换为实际内容
                if last_slot_was_title and current_item:
                    # title 后紧跟的 desc，配对同一 item
                    page_input[slot_name] = current_item.get("desc", "")
                    last_slot_was_title = False
                elif item_idx < len(section_data):
                    # 连续 desc 或 title 被跳过后的 desc，取下一个 item
                    current_item = section_data[item_idx]
                    page_input[slot_name] = current_item.get("desc", "")
                    item_idx += 1
                    last_slot_was_title = False
                else:
                    # 数据耗尽：清空占位文本
                    page_input[slot_name] = ""
            elif slot_name.startswith("item"):
                # item 槽位：作为 desc 处理
                if last_slot_was_title and current_item:
                    page_input[slot_name] = current_item.get("desc", "")
                    last_slot_was_title = False
                elif item_idx < len(section_data):
                    current_item = section_data[item_idx]
                    page_input[slot_name] = current_item.get("desc", "")
                    item_idx += 1
                    last_slot_was_title = False
                else:
                    page_input[slot_name] = ""
        return page_input, item_idx

    def _fill_end(self, slot_data, chapters, page_slots, business_data, schema):
        """填充结束页"""
        end_chapter = next((c for c in chapters if c.get("key") == "end"), None)
        if not end_chapter:
            return
        end_page = str(end_chapter["page"])
        end_slots = page_slots.get(end_page, [])
        end_input = business_data.get("end", {})
        if not end_input:
            return

        thanks = end_input.get("thanks", "")
        if not thanks:
            return

        # 结束页第一个 title 槽位填致谢语
        for slot_info in end_slots:
            if slot_info.get("slot", "").startswith("title"):
                slot_data.setdefault(end_page, {})[slot_info["slot"]] = thanks
                break

    def render(self, scene, business_data, template_id=None, output_path=None, auto_fit=True,
               transitions=None, animations=None):
        """
        一键渲染：业务数据 → 适配 → 渲染
        :param scene: 场景名
        :param business_data: 业务字段
        :param template_id: 模板ID（不指定则取该分类第一个）
        :param output_path: 输出路径
        :param auto_fit: 是否启用长文本字号自适应
        :param transitions: 可选，转场配置 {"页码": {"type": "fade", "speed": "med"}, ...} 或 "auto"
                            "auto" 表示为每页添加 fade 转场（speed=med）
        :param animations: 可选，动画配置 {"页码": [...], ...} 或 "auto"
                           "auto" 表示为每页添加推荐动画（基于页面位置推荐）
        :return: 输出文件路径
        """
        # 选取模板
        if template_id:
            templates = self.list_templates()
            template = next((t for t in templates if t["template_id"] == template_id), None)
            if not template:
                raise ValueError(f"模板不存在: {template_id}")
        else:
            templates = self.list_templates(category=scene)
            if not templates:
                raise ValueError(f"场景 {scene} 无可用模板")
            template = templates[0]
            template_id = template["template_id"]

        meta, meta_path = self.get_template_meta(path=template["path"])
        meta["_meta_file_path"] = meta_path
        pptx_path = self.get_template_pptx(meta)

        # 适配字段
        slot_data = self.adapt(scene, business_data, meta)

        # 渲染
        if not output_path:
            output_path = f"output_{scene}_{template_id}.pptx"

        renderer = PptRenderer(pptx_path, meta_path)
        renderer.render(slot_data, output_path, remove_copyright=True, auto_fit=auto_fit,
                        transitions=transitions, animations=animations)
        return output_path

    def render_batch(self, scene, business_data, output_dir="output_batch", auto_fit=True):
        """
        批量生成：同一份内容生成同分类所有模板 PPT
        :param scene: 场景名
        :param business_data: 业务字段
        :param output_dir: 输出目录
        :return: 生成文件列表
        """
        templates = self.list_templates(category=scene)
        if not templates:
            raise ValueError(f"场景 {scene} 无可用模板")

        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        results = []
        for template in templates:
            template_id = template["template_id"]
            meta, meta_path = self.get_template_meta(path=template["path"])
            meta["_meta_file_path"] = meta_path
            pptx_path = self.get_template_pptx(meta)

            slot_data = self.adapt(scene, business_data, meta)
            output_path = str(output_dir / f"{scene}_{template_id}.pptx")

            try:
                renderer = PptRenderer(pptx_path, meta_path)
                renderer.render(slot_data, output_path, remove_copyright=True, auto_fit=auto_fit)
                results.append({"template_id": template_id, "output": output_path, "status": "success"})
            except Exception as e:
                results.append({"template_id": template_id, "output": None, "status": f"failed: {e}"})

        return results


# ==================== CLI 入口 ====================
def main():
    import argparse
    parser = argparse.ArgumentParser(description='场景适配层 - 业务字段一键生成 PPT')
    sub = parser.add_subparsers(dest='command', required=True)

    # 列出场景
    sub.add_parser('scenes', help='列出所有支持的场景')

    # 场景Schema
    ls = sub.add_parser('schema', help='查看场景Schema定义')
    ls.add_argument('--scene', required=True, help='场景名')

    # 列出模板
    lt = sub.add_parser('templates', help='列出模板（可按分类筛选）')
    lt.add_argument('--category', help='分类名')

    # 模板详情
    ld = sub.add_parser('detail', help='查看模板详情')
    ld.add_argument('--template', required=True, help='模板ID')

    # 数据校验
    lv = sub.add_parser('validate', help='校验业务数据是否符合场景Schema')
    lv.add_argument('--scene', required=True, help='场景名')
    lv.add_argument('--data', required=True, help='业务数据 JSON 路径')

    # 单模板渲染
    r = sub.add_parser('render', help='单模板渲染')
    r.add_argument('--scene', required=True, help='场景名')
    r.add_argument('--data', required=True, help='业务数据 JSON 路径')
    r.add_argument('--template', help='模板ID（不指定则取该分类第一个）')
    r.add_argument('--output', help='输出路径')

    # 批量渲染
    rb = sub.add_parser('batch', help='批量渲染（同分类所有模板）')
    rb.add_argument('--scene', required=True, help='场景名')
    rb.add_argument('--data', required=True, help='业务数据 JSON 路径')
    rb.add_argument('--output-dir', default='output_batch', help='输出目录')

    args = parser.parse_args()
    adapter = SceneAdapter()

    if args.command == 'scenes':
        print("支持的场景：")
        for s in adapter.list_scenes():
            print(f"  {s['category']}: {s['name']}")
        print(f"\n模板索引：共 {adapter.index['total']} 套")
        for cat, cnt in adapter.index.get('categories', {}).items():
            print(f"  {cat}: {cnt} 套")

    elif args.command == 'schema':
        schema = adapter.get_scene_schema(args.scene)
        print(f"场景: {schema['scene']} ({schema['name']})\n")
        print("封面字段 (cover):")
        for k, v in schema['cover_fields'].items():
            print(f"  {k}: {v}")
        print("\n章节 (sections):")
        for s in schema['chapter_sections']:
            print(f"  {s['key']}: {s['name']} - {s['desc']}")
        print("\n结束字段 (end):")
        for k, v in schema['end_fields'].items():
            print(f"  {k}: {v}")
        print(f"\n输入格式: {schema['input_format']}")

    elif args.command == 'templates':
        templates = adapter.list_templates(category=args.category)
        print(f"找到 {len(templates)} 个模板：")
        for t in templates:
            print(f"  [{t['template_id']}] {t['name']} | {t['category']} | {t['total_pages']}页")

    elif args.command == 'detail':
        detail = adapter.get_template_detail(args.template)
        print(f"模板: {detail['template_id']}")
        print(f"分类: {detail['category']}")
        print(f"总页数: {detail['total_pages']} (版权页: {detail['removable_pages']})")
        print(f"章节数: {detail['chapter_count']}")
        print(f"槽位总数: {detail['total_slots']}")
        print("\n章节结构:")
        for c in detail['chapters']:
            if c.get('page'):
                print(f"  {c['key']}: 页{c['page']} ({c['name']})")
            else:
                print(f"  {c['key']}: 页{c['start_page']}-{c['end_page']} ({c['name']})")
        print(f"\n各页槽位数: {detail['page_slot_summary']}")

    elif args.command == 'validate':
        with open(args.data, 'r', encoding='utf-8') as f:
            business_data = json.load(f)
        is_valid, issues = adapter.validate_business_data(args.scene, business_data)
        if is_valid:
            print(f"✅ 数据校验通过: {args.scene} 场景数据格式正确")
        else:
            print(f"❌ 数据校验失败，发现 {len(issues)} 个问题:")
            for issue in issues:
                print(f"  ⚠️  {issue}")

    elif args.command == 'render':
        with open(args.data, 'r', encoding='utf-8') as f:
            business_data = json.load(f)
        out = adapter.render(args.scene, business_data, template_id=args.template, output_path=args.output)
        print(f"\n📁 输出文件: {out}")

    elif args.command == 'batch':
        with open(args.data, 'r', encoding='utf-8') as f:
            business_data = json.load(f)
        results = adapter.render_batch(args.scene, business_data, output_dir=args.output_dir)
        print(f"\n📊 批量生成完成：{len([r for r in results if r['status']=='success'])}/{len(results)} 成功")
        for r in results:
            status = "✅" if r['status'] == 'success' else "❌"
            print(f"  {status} {r['template_id']}: {r['output'] or r['status']}")


if __name__ == '__main__':
    main()
