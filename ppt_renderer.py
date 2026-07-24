"""
PPT 渲染引擎 - 阶段二核心模块
功能：基于 meta 元数据，将结构化内容精准替换到模板，100% 保留原样式
依赖：python-pptx
"""
import json
import argparse
from pathlib import Path
from copy import deepcopy
from typing import Any, Optional
from pptx import Presentation
from pptx.slide import Slide
from pptx.shapes.base import BaseShape
from pptx.util import Pt

from aippt.config import DEFAULT_REMOVE_COPYRIGHT, DEFAULT_AUTO_FIT
from aippt.logger import logger

try:
    from ppt_transitions import inject_transition
    from ppt_animations import inject_animations, RECOMMENDED_ANIMATIONS
    _ANIM_MODULES_READY = True
except Exception:
    _ANIM_MODULES_READY = False
    _ANIM_IMPORT_ERROR = "animation/transition modules not available"


class PptRenderer:
    """PPT 渲染引擎：模板 + meta + 结构化内容 → 成品 PPT"""

    def __init__(self, template_path: str, meta_path: str) -> None:
        template = Path(template_path)
        if not template.exists():
            raise FileNotFoundError(f"模板文件不存在: {template_path}")
        self.template_path = str(template)
        meta_file = Path(meta_path)
        if not meta_file.exists():
            raise FileNotFoundError(f"meta 文件不存在: {meta_path}")
        with open(meta_file, 'r', encoding='utf-8') as f:
            self.meta: dict[str, Any] = json.load(f)

    def render(
        self,
        slot_data: dict[str, dict[str, str]],
        output_path: str,
        remove_copyright: bool = DEFAULT_REMOVE_COPYRIGHT,
        auto_fit: bool = DEFAULT_AUTO_FIT,
        transitions: Optional[Any] = None,
        animations: Optional[Any] = None,
    ) -> str:
        prs = Presentation(self.template_path)
        page_slots = self.meta.get('page_slots', {})
        stats: dict[str, int] = {'replaced': 0, 'missed': 0, 'skipped': 0}

        for page_str, slots in page_slots.items():
            page_num = int(page_str)
            if page_num < 1 or page_num > len(prs.slides):
                continue
            slide = prs.slides[page_num - 1]
            page_input = slot_data.get(page_str, {})
            if not page_input:
                continue

            available_shapes = [
                s for s in self._iter_all_shapes(slide.shapes)
                if s.has_text_frame and s.text_frame.text.strip()
            ]
            used: set[int] = set()

            for slot_info in slots:
                slot_name = slot_info.get('slot', '')
                if slot_name not in page_input:
                    stats['skipped'] += 1
                    continue
                new_value = page_input[slot_name]
                if new_value is None:
                    stats['skipped'] += 1
                    continue
                shape = self._find_shape(available_shapes, slot_info, used)
                if shape is None:
                    stats['missed'] += 1
                    logger.warning("页%s 槽位 %s 未匹配到 shape", page_str, slot_name)
                    continue
                original_text = shape.text_frame.text
                self._replace_text(shape, str(new_value))
                if auto_fit:
                    self._auto_fit(shape, str(new_value), original_text, slot_info, page_str, slot_name)
                used.add(id(shape._element))
                stats['replaced'] += 1

        removed_pages: list[int] = []
        if remove_copyright and self.meta.get('removable_pages'):
            removed_pages = self._remove_slides(prs, self.meta['removable_pages'])

        if transitions or animations:
            if not _ANIM_MODULES_READY:
                logger.warning("动画/转场模块加载失败，跳过注入: %s", _ANIM_IMPORT_ERROR)
            else:
                self._inject_effects(prs, transitions, animations)

        prs.save(str(output_path))
        logger.info("渲染完成: %s", output_path)
        logger.info("替换 %d / 未匹配 %d / 跳过 %d", stats['replaced'], stats['missed'], stats['skipped'])
        if removed_pages:
            logger.info("已删除版权页: %s", removed_pages)
        return str(output_path)

    def _inject_effects(
        self,
        prs: Any,
        transitions: Any,
        animations: Any,
    ) -> None:
        chapters = self.meta.get('chapters', [])
        total = len(prs.slides)

        for page_num, slide in enumerate(prs.slides, 1):
            if transitions:
                if transitions == "auto":
                    inject_transition(slide, {"type": "fade", "speed": "med"})
                elif isinstance(transitions, dict) and str(page_num) in transitions:
                    inject_transition(slide, transitions[str(page_num)])

            if animations:
                slide_type = self._infer_slide_type(page_num, chapters, total)

                if animations == "auto":
                    anim_spec = RECOMMENDED_ANIMATIONS.get(
                        slide_type, RECOMMENDED_ANIMATIONS.get("CONTENT", [])
                    )
                    if anim_spec:
                        inject_animations(slide, anim_spec, slide_type)
                elif isinstance(animations, dict) and str(page_num) in animations:
                    anim_spec = animations[str(page_num)]
                    if anim_spec:
                        inject_animations(slide, anim_spec, slide_type)

    def _infer_slide_type(self, page_num: int, chapters: list[dict[str, Any]], total_pages: int) -> str:
        # 优先识别复合页面：含 chart/table 形状的页面
        page_meta_entry = self.meta.get('page_meta', {}).get(str(page_num))
        if page_meta_entry:
            if page_meta_entry.get("has_chart"):
                return "CHART"
            if page_meta_entry.get("has_table"):
                return "TABLE"

        # 章节页匹配
        for ch in chapters:
            key = ch.get("key", "")
            page = ch.get("page")
            start_page = ch.get("start_page")
            end_page = ch.get("end_page")

            if key == "cover" and page == page_num:
                return "COVER"
            if key == "end" and page == page_num:
                return "END"
            # 章节分隔页（单页 chapter）
            if key and key.startswith("chapter_") and start_page == end_page and start_page == page_num:
                return "CHAPTER"
            # 章节首页也可能是分隔页
            if key and key.startswith("chapter_") and start_page == page_num and start_page != end_page:
                # 仅当 page_slots 中 title 含 PART 字样时算 CHAPTER
                page_slots = self.meta.get('page_slots', {}).get(str(page_num), [])
                for s in page_slots:
                    mt = s.get('match_text', '')
                    if 'PART' in mt.upper() or '第' in mt and '章' in mt:
                        return "CHAPTER"
                return "CONTENT"

        # 末页通常是结束页
        if page_num == total_pages:
            return "END"

        return "CONTENT"

    def _iter_all_shapes(self, shapes):
        for s in shapes:
            yield s
            if s.shape_type == 6:  # GROUP
                yield from self._iter_all_shapes(s.shapes)

    def _find_shape(self, shapes, slot_info: dict[str, Any], used: set[int]) -> Optional[BaseShape]:
        match_text = slot_info.get('match_text', '').strip()
        shape_name = slot_info.get('shape_name', '')
        all_shapes = list(self._iter_all_shapes(shapes))

        # 策略1：shape_name + match_text 双匹配（最精确）
        if shape_name and match_text:
            for s in all_shapes:
                if id(s._element) in used:
                    continue
                if not s.has_text_frame:
                    continue
                if s.name == shape_name and match_text in s.text_frame.text:
                    return s

        # 策略2：match_text 文本匹配
        if match_text:
            for s in all_shapes:
                if id(s._element) in used:
                    continue
                if not s.has_text_frame:
                    continue
                if match_text in s.text_frame.text:
                    return s

        # 策略3：shape_name 匹配（兜底）
        if shape_name:
            for s in all_shapes:
                if id(s._element) in used:
                    continue
                if not s.has_text_frame:
                    continue
                if s.name == shape_name:
                    return s

        return None

    def _replace_text(self, shape: BaseShape, new_text: str) -> None:
        tf = shape.text_frame
        paragraphs = list(tf.paragraphs)
        if not paragraphs:
            return

        first_para = paragraphs[0]

        # 保留第一个 run 的格式，仅改其 text，删除同段其余 run
        if first_para.runs:
            first_para.runs[0].text = new_text
            for run in first_para.runs[1:]:
                run._r.getparent().remove(run._r)
        else:
            run = first_para.add_run()
            run.text = new_text

        # 删除其余段落（保留首段段落属性）
        for para in paragraphs[1:]:
            para._p.getparent().remove(para._p)

    def _auto_fit(
        self,
        shape: BaseShape,
        new_text: str,
        original_text: str,
        slot_info: Optional[dict[str, Any]] = None,
        page_str: Optional[str] = None,
        slot_name: Optional[str] = None,
    ) -> None:
        tf = shape.text_frame
        if not tf.paragraphs or not tf.paragraphs[0].runs:
            return
        run = tf.paragraphs[0].runs[0]
        if run.font.size is None:
            return
        original_size_pt = run.font.size.pt

        # 策略1：优先使用 meta 中的 capacity（事前预警）
        capacity = None
        if slot_info and isinstance(slot_info.get('capacity'), dict):
            capacity = slot_info['capacity'].get('total_chars')

        # 策略2：回退到运行时几何估算
        if capacity is None:
            try:
                box_w = shape.width        # EMU
                box_h = shape.height       # EMU
            except Exception:
                return
            if not box_w or not box_h:
                return
            EMU_PER_PT = 12700
            char_w = original_size_pt * 0.55 * EMU_PER_PT
            line_h = original_size_pt * 1.2 * EMU_PER_PT
            chars_per_line = max(1, int(box_w / char_w))
            lines_available = max(1, int(box_h / line_h))
            capacity = chars_per_line * lines_available

        text_len = len(new_text)
        if text_len <= capacity:
            return

        if page_str and slot_name:
            logger.warning("页%s 槽位 %s: 文本长度 %d 超过容量 %d，自动缩字号", page_str, slot_name, text_len, capacity)

        # 按比例缩小字号，下限 8pt
        ratio = (capacity / text_len) ** 0.5
        new_size = max(8, int(original_size_pt * ratio))
        if new_size < original_size_pt:
            run.font.size = Pt(new_size)

    def _remove_slides(self, prs: Any, page_numbers: list[int]) -> list[int]:
        removed: list[int] = []
        for page_num in sorted(page_numbers, reverse=True):
            idx = page_num - 1
            if 0 <= idx < len(prs.slides):
                xml_slides = prs.slides._sldIdLst
                slides = list(xml_slides)
                xml_slides.remove(slides[idx])
                removed.append(page_num)
        return removed

    def duplicate_slide(self, prs: Any, slide_index: int):
        source = prs.slides[slide_index]
        new_slide = prs.slides.add_slide(source.slide_layout)
        # 清空新 slide 自带占位符
        for shape in list(new_slide.shapes):
            sp = shape._element
            sp.getparent().remove(sp)
        # 深拷贝源 slide 所有 shape
        for shape in source.shapes:
            new_el = deepcopy(shape._element)
            new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
        return new_slide


# ==================== CLI 入口 ====================
def cmd_render(args):
    with open(args.data, 'r', encoding='utf-8') as f:
        slot_data = json.load(f)
    renderer = PptRenderer(args.template, args.meta)
    renderer.render(
        slot_data,
        args.output,
        remove_copyright=not args.keep_copyright,
        auto_fit=not args.no_fit
    )


def main():
    parser = argparse.ArgumentParser(description='PPT 渲染引擎')
    sub = parser.add_subparsers(dest='command', required=True)

    r = sub.add_parser('render', help='渲染生成 PPT')
    r.add_argument('--template', required=True, help='模板 pptx 路径')
    r.add_argument('--meta', required=True, help='meta.json 路径')
    r.add_argument('--data', required=True, help='内容 JSON 路径（{"页码":{"槽位名":"值"}}）')
    r.add_argument('--output', required=True, help='输出 pptx 路径')
    r.add_argument('--keep-copyright', action='store_true', help='保留版权页')
    r.add_argument('--no-fit', action='store_true', help='关闭字号自适应')
    r.set_defaults(func=cmd_render)

    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
