"""
PPT 渲染引擎 - 阶段二核心模块
功能：基于 meta 元数据，将结构化内容精准替换到模板，100% 保留原样式
依赖：python-pptx

架构说明（T001 + 3.3.3 模块化拆分）：
    主引擎 PptRenderer 继承 BaseRenderer，负责模板渲染调度：
    - 继承 aippt.render.base_renderer.BaseRenderer（双引擎统一抽象层）
    - 具体能力下沉到 aippt 子模块：
      - aippt.text_replacer：文本替换 / shape 定位 / 字号自适应
      - aippt.chart_replacer：图表数据源替换（4 类图表）
      - aippt.table_filler：表格动态行扩展
    - 100% 向后兼容：保留 render(slot_data, ...) 旧签名，新增 render_outline() 适配统一接口
"""
import json
import argparse
from pathlib import Path
from copy import deepcopy
from typing import Any, Optional
from pptx import Presentation
from pptx.slide import Slide
from pptx.shapes.base import BaseShape
from pptx.util import Pt, Inches

from aippt.config import DEFAULT_REMOVE_COPYRIGHT, DEFAULT_AUTO_FIT
from aippt.logger import logger
from aippt.render import BaseRenderer, RenderArgs, RenderResult
from aippt.text_replacer import (
    iter_all_shapes as _iter_all_shapes_impl,
    find_shape as _find_shape_impl,
    replace_text as _replace_text_impl,
    auto_fit as _auto_fit_impl,
)
from aippt.chart_replacer import replace_chart_data as _replace_chart_data_impl
from aippt.table_filler import (
    fill_dynamic_table as _fill_dynamic_table_impl,
    set_cell_text as _set_cell_text_impl,
)
from aippt.image_replacer import replace_images as _replace_images_impl

_ANIM_MODULES_READY = True
_ANIM_IMPORT_ERROR = ""


class PptRenderer(BaseRenderer):
    """PPT 渲染引擎：模板 + meta + 结构化内容 → 成品 PPT

    继承 BaseRenderer（T001 双引擎抽象层）。
    MODE = "template"，与 AutoLayoutRenderer（MODE="auto"）通过 mode 区分。

    向后兼容说明：
        - render(slot_data, output_path, ...) 旧签名保留，所有现有调用方无需修改
        - 新增 render_outline(outline_data, output_path, render_args) 适配统一接口
        - BaseRenderer.render() 抽象方法由 render_outline 实现
    """

    #: 引擎模式标识
    MODE: str = "template"

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
        notes_map: Optional[dict[int, str]] = None,
        animation_theme: Optional[str] = None,
    ) -> str:
        """
        渲染 PPT：模板 + meta + slot_data → 成品 PPT

        :param slot_data: 槽位数据，结构 {"页码字符串": {"槽位名": "值"}}
            含 chart_data/table_data 时自动触发图表/表格替换
        :param output_path: 输出文件路径
        :param remove_copyright: 是否自动删除版权页（默认 True）
        :param auto_fit: 是否启用长文本字号自适应（默认 True）
        :param transitions: 全局转场配置 "auto"/"none"/dict/None
        :param animations: 全局动画配置 "auto"/"none"/dict/None
        :param notes_map: 演讲者备注，键为 page_id（从1开始），值为备注文本
        :param animation_theme: 动画主题名（business/tech/formal），None 不启用
        :return: 输出文件路径
        :raises FileNotFoundError: 模板或 meta 文件不存在
        """
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
                # chart_data/table_data/image_data 由专门逻辑处理，跳过文本替换
                if slot_name in ('chart_data', 'table_data', 'image_data'):
                    continue
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

            # 处理图表/表格形状（检测到 chart_data/table_data 时触发）
            if 'chart_data' in page_input or 'table_data' in page_input:
                for shape in self._iter_all_shapes(slide.shapes):
                    if 'chart_data' in page_input and hasattr(shape, 'has_chart') and shape.has_chart:
                        self._replace_chart_data(shape, page_input['chart_data'])
                        stats['replaced'] += 1
                    elif 'table_data' in page_input and hasattr(shape, 'has_table') and shape.has_table:
                        self._fill_dynamic_table(shape, page_input['table_data'])
                        stats['replaced'] += 1

            # 处理图片替换（检测到 image_data 时触发，T301 新增）
            if 'image_data' in page_input and isinstance(page_input['image_data'], dict):
                replaced_imgs = self._replace_images(slide, page_input['image_data'])
                stats['replaced'] += replaced_imgs

        # 演讲者备注注入：在删除版权页之前进行，page_id 与原始页码对齐
        # notes_map 键为 page_id（从 1 开始），值为该页备注文本；默认 None 不注入
        if notes_map:
            for idx, slide in enumerate(prs.slides):
                page_id = idx + 1
                if page_id in notes_map:
                    self._inject_notes(slide, notes_map[page_id])

        removed_pages: list[int] = []
        if remove_copyright and self.meta.get('removable_pages'):
            removed_pages = self._remove_slides(prs, self.meta['removable_pages'])

        # 动画主题：在 _inject_effects 之前，根据主题为每页设置默认 transition/animations
        # 优先级：单页显式 transitions/animations（dict）> 主题 page_overrides > 主题 global_transition
        if animation_theme:
            transitions, animations = self._apply_animation_theme(
                prs, animation_theme, transitions, animations
            )

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
        # 缓存渲染统计供 render_outline 复用
        self._last_stats = stats
        self._last_removed_pages = removed_pages
        return str(output_path)

    # ==================== 统一抽象接口实现（T001）====================

    def render_outline(
        self,
        outline_data: dict[str, Any],
        output_path: str,
        render_args: Optional[RenderArgs] = None,
    ) -> RenderResult:
        """BaseRenderer 统一接口实现：outline → business_data → slot_data → 渲染

        本方法为 PptRenderer 适配 BaseRenderer 抽象接口的入口，内部复用现有
        SceneAdapter.adapt() 完成 outline → slot_data 转换，再调用 render()。

        :param outline_data: outline.json 原始字典，必须含 scene 字段
        :param output_path: 输出 PPTX 路径
        :param render_args: 渲染参数容器，None 使用默认值
        :return: RenderResult 渲染结果
        :raises ValueError: outline_data 缺少 scene 字段
        """
        args = self.normalize_args(render_args)
        scene = outline_data.get("scene")
        if not scene:
            raise ValueError("outline_data 缺少 scene 字段，template 模式必须指定场景")

        # 延迟导入避免循环依赖
        from ppt_scene_adapter import SceneAdapter
        from aippt_outline import outline_to_business_data

        adapter = SceneAdapter()
        business_data = outline_to_business_data(outline_data)
        slot_data = adapter.adapt(scene, business_data, self.meta)

        # 调用现有 render() 完成实际渲染（100% 复用现有逻辑）
        self.render(
            slot_data,
            output_path,
            remove_copyright=args.remove_copyright,
            auto_fit=args.auto_fit,
            transitions=args.transitions,
            animations=args.animations,
            notes_map=args.notes_map,
            animation_theme=args.animation_theme,
        )

        # 收集渲染统计（render() 已缓存到 self._last_stats）
        stats = getattr(self, "_last_stats", {"replaced": 0, "missed": 0, "skipped": 0})
        removed = getattr(self, "_last_removed_pages", [])

        from pptx import Presentation as _Prs
        total_pages = len(_Prs(output_path).slides)

        return RenderResult(
            output_path=output_path,
            mode=self.MODE,
            total_pages=total_pages,
            replaced=stats.get("replaced", 0),
            missed=stats.get("missed", 0),
            skipped=stats.get("skipped", 0),
            removed_pages=removed,
            warnings=[],
            meta={"template_path": self.template_path},
        )

    def _apply_animation_theme(
        self,
        prs: Any,
        theme_name: str,
        transitions: Any,
        animations: Any,
    ) -> tuple[Any, Any]:
        """根据动画主题为每页构建默认 transition/animations，并与显式配置合并

        优先级：显式 transitions/animations（dict）> 主题 page_overrides > 主题 global_transition

        :param prs: Presentation 对象
        :param theme_name: 主题名（business/tech/formal）
        :param transitions: 已有的 transitions 配置（"auto"/"none"/dict/None）
        :param animations: 已有的 animations 配置（"auto"/"none"/dict/None）
        :return: 合并后的 (transitions, animations)
        """
        try:
            from aippt.animation_themes import (
                get_theme, build_page_transition_spec, build_page_animations_spec,
                SLIDE_TYPE_TO_PAGE_TYPE,
            )
        except Exception as e:
            logger.warning("动画主题模块加载失败，跳过主题注入: %s", e)
            return transitions, animations

        try:
            theme = get_theme(theme_name)
        except KeyError as e:
            logger.warning("%s，跳过主题注入", e)
            return transitions, animations

        chapters = self.meta.get('chapters', [])
        total = len(prs.slides)
        overrides = theme.get("page_overrides", {})
        global_transition = theme.get("global_transition")

        # 构建主题默认的 transitions/animations dict
        theme_transitions: dict[str, dict] = {}
        theme_animations: dict[str, list] = {}
        for page_num in range(1, total + 1):
            slide_type = self._infer_slide_type(page_num, chapters, total)
            page_type = SLIDE_TYPE_TO_PAGE_TYPE.get(slide_type, "numbered_list")
            page_cfg = overrides.get(page_type, {})

            # transition：page_overrides > global_transition
            t_name = page_cfg.get("transition") or global_transition
            if t_name and t_name != "none":
                t_spec = build_page_transition_spec(t_name)
                if t_spec:
                    theme_transitions[str(page_num)] = t_spec

            # animations：仅 page_overrides 中有配置的页面才注入
            a_cfg = page_cfg.get("animations")
            if a_cfg:
                a_spec = build_page_animations_spec(page_type, a_cfg)
                if a_spec:
                    theme_animations[str(page_num)] = a_spec

        # 合并：显式 dict 覆盖主题默认（单页 outline 配置优先级最高）
        if isinstance(transitions, dict):
            theme_transitions.update(transitions)
        if isinstance(animations, dict):
            theme_animations.update(animations)

        logger.info("动画主题 %s 已应用: %d 页转场, %d 页动画",
                    theme_name, len(theme_transitions), len(theme_animations))
        return (theme_transitions if theme_transitions else transitions,
                theme_animations if theme_animations else animations)

    def _inject_notes(self, slide: Any, notes_text: str) -> None:
        """
        注入演讲者备注到 slide

        使用 slide.notes_slide.notes_text_frame.text 写入备注文本；
        若 slide 无 notes_slide，python-pptx 会自动创建。

        :param slide: python-pptx 的 Slide 对象
        :param notes_text: 备注文本
        """
        try:
            slide.notes_slide.notes_text_frame.text = notes_text
        except Exception as e:
            logger.warning("演讲者备注注入失败: %s", e)

    def _inject_effects(
        self,
        prs: Any,
        transitions: Any,
        animations: Any,
    ) -> None:
        from aippt.animation_scheduler import inject_page_effects
        from aippt.animation_themes import SLIDE_TYPE_TO_PAGE_TYPE
        chapters = self.meta.get('chapters', [])
        total = len(prs.slides)
        theme_name = getattr(self, 'animation_theme', None)

        for page_num, slide in enumerate(prs.slides, 1):
            slide_type = self._infer_slide_type(page_num, chapters, total)
            page_type = SLIDE_TYPE_TO_PAGE_TYPE.get(slide_type, "numbered_list")

            t_name = None
            if isinstance(transitions, dict) and str(page_num) in transitions:
                t_spec = transitions[str(page_num)]
                if isinstance(t_spec, dict):
                    t_name = t_spec.get("type")

            a_spec = None
            if isinstance(animations, dict) and str(page_num) in animations:
                a_spec = animations[str(page_num)]

            inject_page_effects(slide, page_type,
                theme_name=theme_name,
                page_animations=a_spec,
                page_transition=t_name,
            )

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

    def _iter_all_shapes(self, shapes: Any) -> Any:
        """递归遍历所有 shape，含 GROUP 内的子 shape（委托 aippt.text_replacer）"""
        yield from _iter_all_shapes_impl(shapes)

    def _find_shape(self, shapes: Any, slot_info: dict[str, Any],
                    used: set[int]) -> Optional[BaseShape]:
        """根据槽位信息定位目标 shape（委托 aippt.text_replacer）"""
        return _find_shape_impl(shapes, slot_info, used)

    def _replace_text(self, shape: BaseShape, new_text: str) -> None:
        """替换 shape 的文本，保留首个 run 的格式（委托 aippt.text_replacer）"""
        _replace_text_impl(shape, new_text)

    def _auto_fit(
        self,
        shape: BaseShape,
        new_text: str,
        original_text: str,
        slot_info: Optional[dict[str, Any]] = None,
        page_str: Optional[str] = None,
        slot_name: Optional[str] = None,
    ) -> None:
        """长文本字号自适应（委托 aippt.text_replacer）"""
        _auto_fit_impl(shape, new_text, original_text, slot_info, page_str, slot_name)

    # ==================== 图表与表格动态扩展（委托子模块）====================

    def _replace_chart_data(self, shape: BaseShape, chart_data_dict: dict[str, Any]) -> None:
        """替换图表数据，100% 保留模板样式（委托 aippt.chart_replacer）

        支持 4 类图表：bar（柱状图）、line（折线图）、pie（饼图）、radar（雷达图）
        多系列自动适配：M>N 仅替换前 N 系列；M<N 多余系列清空数据；M==N 直接替换

        :param shape: GraphicFrame 形状，含 chart
        :param chart_data_dict: {"categories": [...], "series": [{"name": "...", "data": [...]}]}
        """
        _replace_chart_data_impl(shape, chart_data_dict)

    def _fill_dynamic_table(self, shape: BaseShape, table_data: dict[str, Any]) -> None:
        """动态填充表格数据，保留第 0 行表头样式模板（委托 aippt.table_filler）

        - 根据数据行数动态增减行（克隆最后一行保持样式一致性）
        - 填入表头数据（覆盖模板占位文本，保留样式）
        - 填入数据行，每行单元格继承样式（字体/对齐/背景色）
        - 自动行高适配

        :param shape: GraphicFrame 形状，含 table
        :param table_data: {"headers": [...], "rows": [[...], ...]}
        """
        _fill_dynamic_table_impl(shape, table_data)

    def _set_cell_text(self, cell: Any, text: str) -> None:
        """设置单元格文本，保留首个 run 的格式（委托 aippt.table_filler）

        :param cell: pptx table Cell 对象
        :param text: 要填入的文本
        """
        _set_cell_text_impl(cell, text)

    def _replace_images(self, slide: Any, image_data: dict[str, dict[str, Any]]) -> int:
        """替换 slide 中的图片 shape（委托 aippt.image_replacer，T301 新增）

        支持本地路径与 URL，等比覆盖（cover）/等比包含（contain）两种填充模式。
        匹配策略：按 picture shape 出现顺序依次替换 image_data 中的项。

        :param slide: python-pptx Slide 对象
        :param image_data: {slot_name: {"path"/"url": "...", "fit": "cover"/"contain"}}
        :return: 成功替换的图片数量
        """
        return _replace_images_impl(slide, image_data)

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

    def duplicate_slide(self, prs: Any, slide_index: int) -> Any:
        """复制指定 slide，返回新 slide 对象

        :param prs: Presentation 对象
        :param slide_index: 源 slide 索引（0-based）
        :return: 新创建的 slide 对象
        """
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
def cmd_render(args: Any) -> None:
    """CLI render 子命令处理"""
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
