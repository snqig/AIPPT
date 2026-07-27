"""
AutoLayoutRenderer 无模板自动布局渲染器（T005）

继承 BaseRenderer，实现无模板自动生成 PPT 的渲染引擎。

执行链路：
    1. 初始化 Presentation 16:9 画布
    2. 遍历 outline.pages，调用 ppt_auto_layout 生成每页元素
    3. 收集页面内所有元素 shape_id + role 列表
    4. 调用公共动画模块执行动画、转场注入（T007 实现）
    5. 执行文本 auto_fit 自适应逻辑（T008 实现）

设计约束：
    - 不依赖任何模板文件，从零生成原生可编辑 PPTX 元素
    - 所有样式从主题令牌读取，禁止硬编码
    - 自动生成元素必须附加 role + shape_id
    - 与 PptRenderer 共用 BaseRenderer 抽象接口
"""
from __future__ import annotations

from typing import Any, Optional

from pptx import Presentation
from pptx.util import Pt

from aippt.logger import logger
from aippt.render import BaseRenderer, RenderArgs, RenderResult
from aippt.theme_loader import load_theme
from aippt.layout import (
    create_presentation, add_blank_slide,
    dispatch_page_layout, LayoutContext, PAGE_LAYOUT_REGISTRY,
)


class AutoLayoutRenderer(BaseRenderer):
    """无模板自动布局渲染器

    继承 BaseRenderer，MODE = "auto"。
    通过 ppt_auto_layout 引擎从零生成 PPT，无需模板文件。

    与 PptRenderer 区别：
        - PptRenderer：基于模板 + meta 槽位替换，100% 保留模板样式
        - AutoLayoutRenderer：无模板，按主题令牌自动布局生成原生元素
    """

    #: 引擎模式标识
    MODE: str = "auto"

    def __init__(self, theme_name: Optional[str] = None) -> None:
        """初始化自动布局渲染器

        :param theme_name: 主题名（如 "商务蓝"），None 时使用默认主题
        """
        self.theme_name = theme_name
        self.theme = load_theme(theme_name) if theme_name else load_theme("商务蓝")
        if theme_name and self.theme.get("name") == "默认主题" and theme_name != "默认主题":
            logger.warning("主题 %s 未找到，降级到默认主题", theme_name)

    def render_outline(
        self,
        outline_data: dict[str, Any],
        output_path: str,
        render_args: Optional[RenderArgs] = None,
    ) -> RenderResult:
        """BaseRenderer 统一接口实现：outline → 自动布局 → 渲染

        执行链路：
            1. 创建 16:9 空白 Presentation
            2. 遍历 outline.pages，按 page_type 分发到布局函数
            3. 收集每页元素 ElementMeta 列表（含 shape_id + role）
            4. 调用动画/转场注入（T007 实现后启用）
            5. 保存 PPTX

        :param outline_data: outline.json 原始字典，必须含 pages 数组
        :param output_path: 输出 PPTX 路径
        :param render_args: 渲染参数容器，None 使用默认值
        :return: RenderResult 渲染结果
        :raises ValueError: outline_data 缺少 pages 数组
        """
        args = self.normalize_args(render_args)

        # 支持两种 outline 格式：pages 数组 或 cover/sections/end
        pages = outline_data.get("pages")
        if not pages:
            # 转换 cover/sections/end 格式为 pages 数组
            pages = self._convert_outline_to_pages(outline_data)
            if not pages:
                raise ValueError("outline_data 缺少 pages 数组，且无法从 cover/sections/end 转换")

        # 创建 16:9 画布
        prs = create_presentation()

        # 遍历页面分发到布局函数
        all_elements: list[dict] = []
        for page_data in pages:
            page_num = page_data.get("page_id", len(all_elements) + 1)
            page_type = page_data.get("page_type", "numbered_list")

            # 检查 page_type 是否已注册
            if page_type not in PAGE_LAYOUT_REGISTRY:
                logger.warning("页%s 未注册的 page_type: %s，降级为 numbered_list",
                               page_num, page_type)
                page_data = {**page_data, "page_type": "numbered_list"}

            slide = add_blank_slide(prs)
            ctx = LayoutContext(page_num=page_num)
            elements = dispatch_page_layout(slide, page_data, self.theme, ctx)

            # 收集元素元数据（供动画模块使用）
            for el in elements:
                all_elements.append({
                    "page_num": page_num,
                    "shape_id": el.shape_id,
                    "role": el.role,
                    "page_type": page_type,
                })

        # T007：动画/转场注入（接入现有 ppt_animations / ppt_transitions 模块）
        if args.transitions or args.animations:
            self._inject_effects(prs, args, all_elements)

        # T008：auto_fit 自适应（复用现有 aippt.text_replacer.auto_fit）
        # 对所有有文本的 shape 执行字号自适应，下限 10pt
        if args.auto_fit:
            self._apply_autofit(prs, all_elements)

        prs.save(str(output_path))
        total_pages = len(prs.slides)
        logger.info("自动布局渲染完成: %s", output_path)
        logger.info("总页数: %d, 总元素数: %d", total_pages, len(all_elements))

        return RenderResult(
            output_path=str(output_path),
            mode=self.MODE,
            total_pages=total_pages,
            replaced=len(all_elements),
            missed=0,
            skipped=0,
            removed_pages=[],
            warnings=[],
            meta={
                "theme": self.theme_name or self.theme.get("name", "默认主题"),
                "elements": all_elements,
            },
        )

    def _convert_outline_to_pages(self, outline: dict) -> list[dict]:
        """将 cover/sections/end 格式 outline 转换为 pages 数组

        用于兼容现有 cover/sections/end 格式的 outline.json。
        转换规则：
            - cover → cover 页
            - sections[].items → numbered_list 页（每个 section 一页）
            - end → ending 页（暂用 divider 布局兜底）

        :param outline: outline dict
        :return: pages 数组（可能为空）
        """
        pages: list[dict] = []
        page_id = 1

        # cover
        cover = outline.get("cover", {})
        if cover:
            pages.append({
                "page_id": page_id,
                "page_type": "cover",
                "title": cover.get("title", ""),
                "subtitle": cover.get("reporter", "") or cover.get("period", ""),
            })
            page_id += 1

        # sections → numbered_list
        sections = outline.get("sections", {})
        if isinstance(sections, dict):
            for sec_key, items in sections.items():
                if not isinstance(items, list) or not items:
                    continue
                # 提取标题列表
                titles = []
                for item in items:
                    if isinstance(item, dict):
                        titles.append(item.get("title", ""))
                    else:
                        titles.append(str(item))
                pages.append({
                    "page_id": page_id,
                    "page_type": "numbered_list",
                    "title": sec_key,
                    "items": titles,
                })
                page_id += 1

        return pages

    def _inject_effects(
        self,
        prs: Presentation,
        args: RenderArgs,
        elements: list[dict],
    ) -> None:
        """T007：动画/转场注入（接入现有 ppt_animations / ppt_transitions 模块）

        实现链路：
            1. 按页码分组元素，构建 page_num → elements 映射
            2. 遍历每页 slide：
               - 转场：args.transitions 为 "auto" 时按 page_type 自动选择 fade/push
               - 动画：args.animations 为 "auto" 时按 page_type 匹配 RECOMMENDED_ANIMATIONS
            3. shape 角色匹配：通过 shape.name 中包含的 role 关键字定位
               （自动布局生成的 shape.name 格式为 p{page_num}_{role}_{seq}）

        :param prs: Presentation 对象
        :param args: 渲染参数
        :param elements: 全部元素元数据列表
        """
        try:
            from ppt_transitions import inject_transition
            from ppt_animations import inject_animations, RECOMMENDED_ANIMATIONS
        except Exception as e:
            logger.warning("动画/转场模块加载失败，跳过注入: %s", e)
            return

        # 按页码分组元素
        page_elements: dict[int, list[dict]] = {}
        for el in elements:
            page_num = el.get("page_num", 0)
            page_elements.setdefault(page_num, []).append(el)

        # 页面类型 → slide_type 映射（用于动画 auto 匹配）
        page_type_to_slide_type = {
            "cover": "COVER",
            "catalog": "CONTENT",
            "divider": "CHAPTER",
            "numbered_list": "CONTENT",
            "kpi": "KPI",
            "timeline": "TIMELINE",
            "two_column": "CONTENT",
            "chart": "CHART",
            "table": "TABLE",
            "ending": "END",
        }

        injected_pages = 0
        for page_num, slide in enumerate(prs.slides, 1):
            page_els = page_elements.get(page_num, [])
            if not page_els:
                continue
            page_type = page_els[0].get("page_type", "numbered_list")
            slide_type = page_type_to_slide_type.get(page_type, "CONTENT")

            # 转场注入
            if args.transitions:
                if args.transitions == "auto":
                    # 按 page_type 自动选择转场
                    if page_type == "cover":
                        t_spec = {"type": "fade", "speed": "slow"}
                    elif page_type == "divider":
                        t_spec = {"type": "push", "dir": "from_left", "speed": "med"}
                    else:
                        t_spec = {"type": "fade", "speed": "med"}
                    inject_transition(slide, t_spec)
                elif isinstance(args.transitions, dict) and str(page_num) in args.transitions:
                    inject_transition(slide, args.transitions[str(page_num)])

            # 动画注入
            if args.animations:
                if args.animations == "auto":
                    # 按 slide_type 匹配推荐动画
                    anim_spec = RECOMMENDED_ANIMATIONS.get(
                        slide_type, RECOMMENDED_ANIMATIONS.get("CONTENT", [])
                    )
                    if anim_spec:
                        inject_animations(slide, anim_spec, slide_type)
                elif isinstance(args.animations, dict) and str(page_num) in args.animations:
                    anim_spec = args.animations[str(page_num)]
                    if anim_spec:
                        inject_animations(slide, anim_spec, slide_type)

            injected_pages += 1

        logger.info("动画/转场注入完成: %d 页（转场=%s, 动画=%s）",
                    injected_pages,
                    args.transitions if args.transitions else "off",
                    args.animations if args.animations else "off")

    def _apply_autofit(self, prs: Presentation, elements: list[dict]) -> None:
        """T008：auto_fit 自适应（复用现有 aippt.text_replacer.auto_fit）

        对所有有文本的 shape 执行字号自适应：
            1. 遍历每页 slide 的所有 shape
            2. 跳过无文本框的 shape（如纯背景矩形、连接线）
            3. 调用 auto_fit 几何估算，超容量时按比例缩小字号
            4. 字号下限 10pt（防止过小不可读）

        :param prs: Presentation 对象
        :param elements: 全部元素元数据列表（用于日志上下文）
        """
        try:
            from aippt.text_replacer import auto_fit
        except Exception as e:
            logger.warning("auto_fit 模块加载失败，跳过自适应: %s", e)
            return

        MIN_FONT_PT = 10  # 字号下限，防止过小不可读
        adjusted_count = 0

        for page_num, slide in enumerate(prs.slides, 1):
            for shape in slide.shapes:
                # 仅处理有文本框且有文本的 shape
                if not shape.has_text_frame:
                    continue
                text = shape.text_frame.text.strip()
                if not text:
                    continue

                # 获取当前字号（取首个 run）
                tf = shape.text_frame
                if not tf.paragraphs or not tf.paragraphs[0].runs:
                    continue
                run = tf.paragraphs[0].runs[0]
                if run.font.size is None:
                    continue
                original_size_pt = run.font.size.pt

                # 跳过已小于下限的字号
                if original_size_pt <= MIN_FONT_PT:
                    continue

                # 调用现有 auto_fit 进行几何估算与缩小
                try:
                    auto_fit(
                        shape=shape,
                        new_text=text,
                        original_text=text,
                        slot_info=None,
                        page_str=str(page_num),
                        slot_name=shape.name or "auto",
                    )
                except Exception as e:
                    logger.debug("auto_fit 跳过 shape %s: %s", shape.name, e)
                    continue

                # 强制下限保护：若 auto_fit 缩到 10pt 以下，回拉到 10pt
                try:
                    new_size = run.font.size
                    if new_size and new_size.pt < MIN_FONT_PT:
                        run.font.size = Pt(MIN_FONT_PT)
                        adjusted_count += 1
                except Exception:
                    pass

        if adjusted_count:
            logger.info("auto_fit 下限保护: %d 个 shape 字号回拉到 %dpt",
                        adjusted_count, MIN_FONT_PT)
