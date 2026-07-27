import os
import re
import json
import argparse
from pptx import Presentation
from pathlib import Path
from collections import defaultdict
from typing import Any, Optional

from aippt.config import PAGE_TYPE_KEYWORDS, MODELS_ROOT
from aippt.constants import SLOT_MATCH_KEYWORDS
from aippt.logger import logger
from aippt.ppt_element_classifier import classify_page

# meta必填字段校验
META_REQUIRED_FIELDS = ['template_id', 'category', 'total_pages', 'chapters', 'page_slots']

# ==================== 核心解析函数 ====================
def extract_page_texts(slide) -> list[dict[str, Any]]:
    texts: list[dict[str, Any]] = []

    def _extract_from_shape(shape) -> None:
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if text:
                width_emu: Optional[int] = None
                height_emu: Optional[int] = None
                try:
                    width_emu = int(shape.width)
                    height_emu = int(shape.height)
                except Exception:
                    pass
                font_size_pt: Optional[float] = None
                try:
                    for para in shape.text_frame.paragraphs:
                        if para.runs:
                            fs = para.runs[0].font.size
                            if fs is not None:
                                font_size_pt = fs.pt
                                break
                except Exception:
                    pass

                texts.append({
                    'text': text,
                    'shape_name': shape.name,
                    'is_placeholder': shape.is_placeholder,
                    'width_emu': width_emu,
                    'height_emu': height_emu,
                    'font_size_pt': font_size_pt,
                })
        if shape.shape_type == 6:
            for sub in shape.shapes:
                _extract_from_shape(sub)

    for shape in slide.shapes:
        _extract_from_shape(shape)
    return texts


def detect_page_shapes(slide) -> dict[str, Any]:
    flags: dict[str, Any] = {"has_chart": False, "has_table": False, "has_picture": False,
                              "chart_type": None, "picture_count": 0,
                              "chart_shape_name": None, "table_shape_name": None}

    def _scan(shape) -> None:
        # GRAPH_CHART = 3
        if shape.shape_type == 3 or (hasattr(shape, "has_chart") and shape.has_chart):
            flags["has_chart"] = True
            # 记录第一个图表形状名（用于 chart_slot 标识）
            if not flags["chart_shape_name"]:
                flags["chart_shape_name"] = shape.name
            try:
                if hasattr(shape, "chart") and shape.chart is not None:
                    flags["chart_type"] = str(shape.chart.chart_type)
            except Exception:
                pass
        # TABLE = 19 / has_table
        elif (hasattr(shape, "has_table") and shape.has_table) or shape.shape_type == 19:
            flags["has_table"] = True
            # 记录第一个表格形状名（用于 table_slot 标识）
            if not flags["table_shape_name"]:
                flags["table_shape_name"] = shape.name
        # PICTURE = 13
        elif shape.shape_type == 13:
            flags["has_picture"] = True
            flags["picture_count"] += 1
        # GROUP 递归
        if shape.shape_type == 6:
            for sub in shape.shapes:
                _scan(sub)

    for shape in slide.shapes:
        try:
            _scan(shape)
        except Exception:
            continue
    return flags

def detect_page_type(page_texts: list[dict[str, Any]], page_idx: int, total_pages: int) -> str:
    all_text = ' '.join([t['text'] for t in page_texts])

    # 首页强制识别为封面
    if page_idx == 0:
        return 'cover'

    # 目录页优先判断（CONTENTS/目录是强信号）：放在 end 之前，避免目录页含 THANK YOU 装饰文本被误判为 end
    # 放宽文本框数量限制（许多模板目录页有15+文本框）
    catalog_strong = ['CONTENTS', '目 录', '目录页', '目录', '目\n录']
    for kw in catalog_strong:
        if kw in all_text and len(page_texts) <= 25:
            # 但若同时含 PART ONE/第N章 等强章节信号，且文本框少，则可能是章节页带目录样式
            if len(page_texts) <= 6 and any(k in all_text for k in ['PART ONE', 'PART TWO', 'PART.0', 'Part.0']):
                pass  # 让章节判断处理
            else:
                return 'catalog'

    # 末尾3页判断版权页（含版权关键词）
    if page_idx >= total_pages - 3:
        for kw in PAGE_TYPE_KEYWORDS['copyright']:
            if kw.lower() in all_text.lower():
                return 'copyright'
        # 末页若与首页样式重复（封面/装饰页重复），识别为 copyright
        if page_idx == total_pages - 1:
            cover_kw = ['WORK REPORT', 'WORK SUMMARY', 'BUSINESS', 'YOUR LOGO', 'COMPANY']
            # 末页若非结束页且无明显内容，归为版权/装饰页
            if not any(k in all_text for k in ['感谢聆听', '谢谢聆听', 'THANK YOU', '感谢各位', '谢谢观看', '感谢观看']):
                if any(k in all_text for k in cover_kw) or 'CONTENTS' in all_text:
                    return 'copyright'

    # 结束页：要求在末5页内，避免目录页 THANK YOU 装饰文本误命中
    if page_idx >= total_pages - 5:
        end_strong = ['感谢聆听', '谢谢聆听', '谢谢观看', '感谢观看', 'THANK YOU', '感谢各位']
        for kw in end_strong:
            if kw in all_text:
                return 'end'
    if page_idx >= total_pages - 3:
        for kw in ['感谢', '致谢', '谢谢']:
            if kw in all_text:
                return 'end'

    # 章节页优先判断（含 PART ONE/第N章/Part.01 等强信号）
    chapter_strong = ['PART ONE', 'PART TWO', 'PART THREE', 'PART FOUR', 'PART FIVE', 'PART SIX', 'PART SEVEN',
                      'Part.0', 'Part.1', 'Part.01', 'Part.02', 'Part.03', 'Part.04', 'Part.05', 'Part.06', 'Part.07',
                      'PART.0', 'PART.1', 'PART.01', 'PART.02', 'PART.03', 'PART.04', 'PART.05', 'PART.06', 'PART.07',
                      '第1部分', '第2部分', '第3部分', '第4部分', '第1章', '第2章', '第3章', '第4章',
                      'COMPETENCY', 'JOB AWARENESS', 'PART 01', 'PART 02', 'PART 03', 'PART 04']
    for kw in chapter_strong:
        if kw in all_text:
            return 'chapter'
    # 通用章节关键词（需文本框少，避免误判内容页）
    for kw in ['PART ', '第', '章']:
        if kw in all_text and len(page_texts) <= 6:
            return 'chapter'

    # 章节分隔页（数字+标题模式）：文本框少且含纯数字 01-09 + 短中文标题，无目录关键词
    if len(page_texts) <= 6:
        has_number = any(re.match(r'^0[1-9]$', t['text'].strip()) for t in page_texts)
        has_short_cn_title = any(2 <= len(t['text'].strip()) <= 10 and re.search(r'[\u4e00-\u9fa5]', t['text'])
                                 and t['text'].strip() not in ['目录', '目 录'] for t in page_texts)
        no_catalog_kw = not any(k in all_text for k in catalog_strong)
        no_end_kw = not any(k in all_text for k in ['感谢', 'THANK', '致谢'])
        if has_number and has_short_cn_title and no_catalog_kw and no_end_kw:
            return 'chapter'

    return 'content'

def extract_slots(page_texts: list[dict[str, Any]], page_idx: int) -> list[dict[str, Any]]:
    slots = []
    title_counter = 0
    desc_counter = 0
    item_counter = 0
    year_counter = 0
    percent_counter = 0
    number_counter = 0

    for item in page_texts:
        text = item['text']
        # 判断是否为可替换槽位
        is_slot = any(kw in text for kw in SLOT_MATCH_KEYWORDS)
        is_short_title = len(text) < 15 and '\n' not in text

        if not (is_slot or is_short_title):
            continue

        stripped = text.strip()
        # 语义化命名（优先匹配强语义模式，避免数字/年份/百分比被误归为 title）
        if '汇报人' in text or '姓名' in text:
            slot_name = 'reporter'
        elif '年度' in text or '2O2X' in text or '202X' in text:
            slot_name = 'period'
        elif re.match(r'^\d{4}$', stripped):  # 纯年份，如 2019、2027
            year_counter += 1
            slot_name = f'year_{year_counter}' if year_counter > 1 else 'year'
        elif re.match(r'^\d+%$', stripped):  # 百分比，如 67%
            percent_counter += 1
            slot_name = f'percent_{percent_counter}' if percent_counter > 1 else 'percent'
        elif re.match(r'^\d+$', stripped):  # 纯数字，如 01、02
            number_counter += 1
            slot_name = f'number_{number_counter}' if number_counter > 1 else 'number'
        elif '标题' in text or 'TITLE' in text.upper() or '某某' in text or is_short_title:
            title_counter += 1
            slot_name = f'title_{title_counter}' if title_counter > 1 else 'title'
        elif '内容' in text or '录入' in text or '输入' in text or 'text content' in text.lower() or 'your text' in text.lower() or len(text) > 20:
            desc_counter += 1
            slot_name = f'desc_{desc_counter}' if desc_counter > 1 else 'desc'
        elif '项目' in text and ('名称' in text or 'PROJECT' in text.upper()):
            item_counter += 1
            slot_name = f'project_name_{item_counter}'
        elif '%' in text:
            item_counter += 1
            slot_name = f'progress_{item_counter}'
        else:
            item_counter += 1
            slot_name = f'item_{item_counter}'

        # 取前20字符作为匹配特征
        match_text = text[:20].replace('\n', ' ').strip()
        if match_text:
            slot_entry = {
                'slot': slot_name,
                'match_text': match_text,
                'shape_name': item.get('shape_name', '')
            }
            # 计算 capacity（事前预警用）：基于 shape 几何和字号
            capacity = _compute_capacity(item)
            if capacity:
                slot_entry['capacity'] = capacity
            slots.append(slot_entry)

    return slots


def _compute_capacity(item: dict[str, Any]) -> Optional[dict[str, int]]:
    width_emu = item.get('width_emu')
    height_emu = item.get('height_emu')
    font_size_pt = item.get('font_size_pt')
    if not width_emu or not height_emu or not font_size_pt:
        return None

    EMU_PER_PT = 12700
    char_w = font_size_pt * 0.55 * EMU_PER_PT
    line_h = font_size_pt * 1.2 * EMU_PER_PT
    if char_w <= 0 or line_h <= 0:
        return None
    max_chars_per_line = max(1, int(width_emu / char_w))
    max_lines = max(1, int(height_emu / line_h))
    total_chars = max_chars_per_line * max_lines
    # 容量过小（<5）通常意味着 shape 是装饰性小框/角标，capacity 不可靠
    # 此时不写入 capacity，让渲染器回退到运行时几何估算，避免误报预警
    if total_chars < 5:
        return None
    return {
        "max_chars_per_line": max_chars_per_line,
        "max_lines": max_lines,
        "total_chars": total_chars,
    }

def detect_chapters(slides, page_types: list[str]) -> list[dict[str, Any]]:
    chapters = []
    chapter_idx = 0
    chapter_start = None
    chapter_name = ''
    
    for i, ptype in enumerate(page_types):
        if ptype == 'chapter':
            if chapter_start is not None:
                chapters.append({
                    'key': f'chapter_{chapter_idx}',
                    'start_page': chapter_start + 1,
                    'end_page': i,
                    'name': chapter_name
                })
            chapter_idx += 1
            chapter_start = i
            # 提取中文章节名
            page_text = ' '.join([t['text'] for t in extract_page_texts(slides[i])])
            cn_titles = re.findall(r'[\u4e00-\u9fa5]{2,10}', page_text)
            chapter_name = cn_titles[-1] if cn_titles else f'第{chapter_idx}章'
    
    # 补最后一个章节
    if chapter_start is not None:
        end_page = len(page_types)
        if 'end' in page_types:
            end_page = page_types.index('end')
        chapters.append({
            'key': f'chapter_{chapter_idx}',
            'start_page': chapter_start + 1,
            'end_page': end_page,
            'name': chapter_name
        })

    return chapters

def generate_single_meta(pptx_path: Path, category: str) -> tuple[Optional[dict[str, Any]], Optional[str]]:
    try:
        prs = Presentation(str(pptx_path))
    except Exception as e:
        return None, f"解析失败: {str(e)}"
    
    total_pages = len(prs.slides)
    page_types = []
    all_page_slots = {}
    page_meta = {}

    for idx, slide in enumerate(prs.slides):
        page_texts = extract_page_texts(slide)
        ptype = detect_page_type(page_texts, idx, total_pages)
        page_types.append(ptype)
        slots = extract_slots(page_texts, idx)

        # 收集页面非文本形状信息（chart/table/picture）
        shape_flags = detect_page_shapes(slide)

        # 追加图表/表格专用槽位（标识 chart/table 位置，供渲染器识别）
        if shape_flags["has_chart"]:
            slots.append({
                "slot": "chart_data",
                "shape_name": shape_flags.get("chart_shape_name", ""),
                "chart_type": shape_flags.get("chart_type"),
                "match_text": "",  # 图表形状无文本，不参与文本匹配
            })
        if shape_flags["has_table"]:
            slots.append({
                "slot": "table_data",
                "shape_name": shape_flags.get("table_shape_name", ""),
                "match_text": "",  # 表格形状无文本，不参与文本匹配
            })

        if slots:
            all_page_slots[str(idx + 1)] = slots

        page_key = str(idx + 1)
        page_entry: Optional[dict[str, Any]] = None
        if shape_flags["has_chart"] or shape_flags["has_table"] or shape_flags["has_picture"]:
            page_entry = {
                "has_chart": shape_flags["has_chart"],
                "has_table": shape_flags["has_table"],
                "has_picture": shape_flags["has_picture"],
                "chart_type": shape_flags["chart_type"],
                "picture_count": shape_flags["picture_count"],
                "chart_shape_name": shape_flags.get("chart_shape_name"),
                "table_shape_name": shape_flags.get("table_shape_name"),
            }

        # 元素分类告警（低置信度标注，来自 classify_page）
        try:
            classification = classify_page(slide)
            cls_warnings = classification.get("low_confidence_warnings", [])
            if cls_warnings:
                if page_entry is None:
                    page_entry = {}
                page_entry["classification_warnings"] = cls_warnings
        except Exception as e:
            logger.debug("页面分类失败 page %d: %s", idx + 1, e)

        if page_entry is not None:
            page_meta[page_key] = page_entry
    
    # 可删除版权页
    removable_pages = [i+1 for i, t in enumerate(page_types) if t == 'copyright']
    
    # 章节结构
    chapters = detect_chapters(prs.slides, page_types)

    # 补充固定页面（按页码升序合并，避免 cover 缺失时插入位置错乱）
    fixed_pages = []
    if 'cover' in page_types:
        fixed_pages.append({
            'key': 'cover',
            'page': page_types.index('cover') + 1,
            'name': '封面'
        })
    if 'catalog' in page_types:
        fixed_pages.append({
            'key': 'catalog',
            'page': page_types.index('catalog') + 1,
            'name': '目录'
        })
    if 'end' in page_types:
        fixed_pages.append({
            'key': 'end',
            'page': page_types.index('end') + 1,
            'name': '结束页'
        })
    chapters = fixed_pages + chapters
    chapters.sort(key=lambda c: c.get('start_page', c.get('page', 0)))
    
    # 生成模板ID（使用完整文件名确保唯一）
    template_id = f"{category}_{pptx_path.stem}".replace(' ', '_').replace('-', '_').lower()
    
    meta = {
        "template_id": template_id,
        "category": category,
        "style_tags": ["商务", "16:9"],
        "total_pages": total_pages,
        "removable_pages": removable_pages,
        "chapters": chapters,
        "page_slots": all_page_slots
    }
    # 仅在有复合页面时才写入 page_meta，避免空字段污染 meta
    if page_meta:
        meta["page_meta"] = page_meta

    return meta, None

# ==================== 功能1：批量生成meta ====================
def cmd_generate(args):
    root = Path(args.dir)
    pptx_files = [p for p in root.rglob("*.pptx") if not p.name.startswith('~$')]
    logger.info("找到PPT文件: %d 个", len(pptx_files))
    
    success = 0
    failed = []
    
    for pptx_path in pptx_files:
        meta_path = pptx_path.with_suffix('.meta.json')
        if meta_path.exists() and not args.force:
            logger.info("跳过(已存在): %s", pptx_path.name)
            success += 1
            continue
            
        category = pptx_path.parent.name if pptx_path.parent.name != root.name else '通用'
        
        meta, err = generate_single_meta(pptx_path, category)
        if err:
            failed.append(f"{pptx_path.relative_to(root)}: {err}")
            logger.error("失败: %s - %s", pptx_path.name, err)
            continue
        
        with open(meta_path, 'w', encoding='utf-8') as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        
        logger.info("生成成功: %s", pptx_path.name)
        success += 1
    
    logger.info("生成完成：成功 %d 个，失败 %d 个", success, len(failed))
    if failed:
        logger.warning("失败列表：")
        for f in failed:
            logger.warning("  - %s", f)

# ==================== 功能2：批量校验meta质量 ====================
def cmd_check(args):
    root = Path(args.dir)
    meta_files = list(root.rglob("*.meta.json"))
    logger.info("找到meta文件: %d 个", len(meta_files))
    
    issues = []
    category_count = defaultdict(int)
    
    for meta_path in meta_files:
        try:
            with open(meta_path, 'r', encoding='utf-8') as f:
                meta = json.load(f)
        except Exception:
            issues.append(f"{meta_path.relative_to(root)}: 文件损坏无法读取")
            continue
        
        rel_path = str(meta_path.relative_to(root))
        category = meta.get('category', '未知')
        category_count[category] += 1
        
        for field in META_REQUIRED_FIELDS:
            if field not in meta:
                issues.append(f"{rel_path}: 缺少必填字段 {field}")
        
        chapters = meta.get('chapters', [])
        if len(chapters) < 2:
            issues.append(f"{rel_path}: 章节数量过少({len(chapters)}个)，识别可能不准确")
        
        slots_count = sum(len(v) for v in meta.get('page_slots', {}).values())
        if slots_count < 5:
            issues.append(f"{rel_path}: 可替换槽位过少({slots_count}个)，识别可能不准确")
        
        if not meta.get('removable_pages'):
            end_page = next((c.get('page') for c in chapters if c.get('key') == 'end'), None)
            has_end = end_page is not None
            if has_end and end_page != meta.get('total_pages'):
                issues.append(f"{rel_path}: 未识别到版权页，需手动确认")
    
    logger.info("分类统计：")
    for cat, cnt in category_count.items():
        logger.info("  %s: %d 套", cat, cnt)
    
    logger.info("校验完成：共检查 %d 个meta，发现问题 %d 个", len(meta_files), len(issues))
    if issues:
        logger.warning("问题清单：")
        for issue in issues:
            logger.warning("  %s", issue)
    else:
        logger.info("全部meta校验通过！")

# ==================== 功能3：生成总索引文件 ====================
def cmd_index(args):
    root = Path(args.dir)
    meta_files = list(root.rglob("*.meta.json"))
    logger.info("找到meta文件: %d 个", len(meta_files))
    
    index = {
        "total": len(meta_files),
        "categories": {},
        "templates": []
    }
    
    category_map = defaultdict(list)
    
    for meta_path in meta_files:
        try:
            with open(meta_path, 'r', encoding='utf-8') as f:
                meta = json.load(f)
        except:
            continue
        
        # 精简索引信息
        template_info = {
            "template_id": meta['template_id'],
            "category": meta['category'],
            "name": meta_path.stem.replace('.meta', ''),
            "path": str(meta_path.relative_to(root)),
            "style_tags": meta.get('style_tags', []),
            "total_pages": meta['total_pages'],
            "chapter_count": len(meta.get('chapters', []))
        }
        
        index['templates'].append(template_info)
        category_map[meta['category']].append(template_info)
    
    index['categories'] = {k: len(v) for k, v in category_map.items()}
    
    output_path = root / 'templates_index.json'
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
    
    logger.info("总索引生成完成，共 %d 个模板", len(index['templates']))
    logger.info("索引文件路径: %s", output_path)
    logger.info("分类数量统计：")
    for cat, cnt in index['categories'].items():
        logger.info("  %s: %d 套", cat, cnt)

# ==================== 主入口 ====================
def _find_meta_by_template_id(template_id: str, models_root: Path = None) -> Optional[Path]:
    """根据 template_id 在 models 目录下查找对应的 meta.json 文件"""
    root = models_root or MODELS_ROOT
    for meta_path in root.rglob("*.meta.json"):
        try:
            with open(meta_path, 'r', encoding='utf-8') as f:
                meta = json.load(f)
            if meta.get('template_id') == template_id:
                return meta_path
        except Exception:
            continue
    return None


def cmd_info(args):
    """查看单个模板的元数据详情"""
    meta_path = _find_meta_by_template_id(args.template_id)
    if not meta_path:
        logger.error("未找到 template_id: %s", args.template_id)
        return

    with open(meta_path, 'r', encoding='utf-8') as f:
        meta = json.load(f)

    info = {
        "template_id": meta.get("template_id"),
        "category": meta.get("category"),
        "total_pages": meta.get("total_pages"),
        "removable_pages": meta.get("removable_pages", []),
        "chapters": meta.get("chapters", []),
        "page_slot_count": len(meta.get("page_slots", {})),
        "total_slots": sum(len(v) for v in meta.get("page_slots", {}).values()),
        "page_meta": meta.get("page_meta", {}),
        "meta_path": str(meta_path),
    }
    print(json.dumps(info, ensure_ascii=False, indent=2))


def cmd_check_one(args):
    """单个模板质量校验"""
    meta_path = _find_meta_by_template_id(args.template_id)
    if not meta_path:
        logger.error("未找到 template_id: %s", args.template_id)
        return

    with open(meta_path, 'r', encoding='utf-8') as f:
        meta = json.load(f)

    issues = []
    for field in META_REQUIRED_FIELDS:
        if field not in meta:
            issues.append(f"缺少必填字段 {field}")

    chapters = meta.get('chapters', [])
    if len(chapters) < 2:
        issues.append(f"章节数量过少({len(chapters)}个)")

    slots_count = sum(len(v) for v in meta.get('page_slots', {}).values())
    if slots_count < 5:
        issues.append(f"可替换槽位过少({slots_count}个)")

    result = {
        "template_id": meta.get("template_id"),
        "total_pages": meta.get("total_pages"),
        "chapter_count": len(chapters),
        "slots_count": slots_count,
        "issues": issues,
        "is_valid": len(issues) == 0,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


def main():
    parser = argparse.ArgumentParser(description='PPT模板元数据批量处理工具')
    subparsers = parser.add_subparsers(dest='command', required=True, help='支持的命令')

    # generate 子命令
    gen_parser = subparsers.add_parser('generate', help='批量生成meta.json元数据')
    gen_parser.add_argument('--dir', required=True, help='模板根目录路径')
    gen_parser.add_argument('--force', action='store_true', help='强制覆盖已存在的meta文件')

    # check 子命令（批量）
    check_parser = subparsers.add_parser('check', help='批量校验meta质量')
    check_parser.add_argument('--dir', help='模板根目录路径（批量校验）')
    check_parser.add_argument('--template-id', help='单个模板 ID（单模板校验）')

    # info 子命令（单模板详情）
    info_parser = subparsers.add_parser('info', help='查看单个模板元数据详情')
    info_parser.add_argument('--template-id', required=True, help='模板 ID')

    # index 子命令
    index_parser = subparsers.add_parser('index', help='生成模板总索引文件')
    index_parser.add_argument('--dir', required=True, help='模板根目录路径')

    args = parser.parse_args()

    if args.command == 'generate':
        cmd_generate(args)
    elif args.command == 'check':
        if args.template_id:
            cmd_check_one(args)
        elif args.dir:
            cmd_check(args)
        else:
            parser.error("check 命令需要 --dir 或 --template-id 参数")
    elif args.command == 'info':
        cmd_info(args)
    elif args.command == 'index':
        cmd_index(args)

if __name__ == '__main__':
    main()