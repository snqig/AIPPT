"""
模板母版布局 Profile 模块

遍历 Presentation 的母版（slide_master）下的所有 slide_layout，
提取每个 layout 的占位符结构（idx/type/名称/几何/默认字体字号），
返回 {layout_idx: {name, placeholders: [...]}} 结构化字典。

用于模板能力画像，辅助槽位匹配与渲染预警，100% 向后兼容现有槽位替换架构。
"""
from typing import Any

from aippt.logger import logger


def _extract_default_font(ph) -> tuple[Any, Any]:
    """
    提取占位符的默认字体名与字号（pt）
    优先取段落级 defRPr（paragraph.font），回退到首个 run 的字体
    """
    font_name: Any = None
    font_size: Any = None
    try:
        if not ph.has_text_frame:
            return None, None
        tf = ph.text_frame
        if tf.paragraphs:
            p_font = tf.paragraphs[0].font
            font_name = p_font.name
            if p_font.size is not None:
                font_size = p_font.size.pt
        # 回退：取首个带字体信息的 run
        if font_name is None or font_size is None:
            for para in tf.paragraphs:
                for run in para.runs:
                    if font_name is None and run.font.name:
                        font_name = run.font.name
                    if font_size is None and run.font.size is not None:
                        font_size = run.font.size.pt
                    if font_name is not None and font_size is not None:
                        break
                if font_name is not None and font_size is not None:
                    break
    except Exception as e:
        logger.debug("提取占位符默认字体失败: %s", e)
    return font_name, font_size


def _extract_placeholder_info(ph) -> dict[str, Any]:
    """提取单个占位符的结构化信息"""
    info: dict[str, Any] = {
        "idx": None,
        "type": None,
        "name": None,
        "left": None,
        "top": None,
        "width": None,
        "height": None,
        "default_font": None,
        "default_size": None,
    }
    try:
        info["name"] = ph.name
    except Exception:
        pass
    try:
        info["idx"] = ph.placeholder_format.idx
    except Exception:
        pass
    try:
        ph_type = ph.placeholder_format.type
        info["type"] = str(ph_type) if ph_type is not None else None
    except Exception:
        pass
    # 几何信息（EMU）
    try:
        info["left"] = ph.left
        info["top"] = ph.top
        info["width"] = ph.width
        info["height"] = ph.height
    except Exception:
        pass
    # 默认字体与字号
    try:
        font_name, font_size = _extract_default_font(ph)
        info["default_font"] = font_name
        info["default_size"] = font_size
    except Exception:
        pass
    return info


def profile_layouts(prs) -> dict[int, dict[str, Any]]:
    """
    遍历 Presentation 母版下的所有 slide_layout，提取占位符结构

    :param prs: python-pptx Presentation 对象
    :return: {layout_idx: {"name": str, "placeholders": [{...}]}}
    """
    result: dict[int, dict[str, Any]] = {}

    # 兼容多母版：遍历 prs.slide_masters，单母版场景等价于 prs.slide_master
    try:
        masters = prs.slide_masters
    except Exception:
        try:
            masters = [prs.slide_master]
        except Exception as e:
            logger.warning("无法访问 Presentation 母版: %s", e)
            return result

    for master in masters:
        try:
            layouts = master.slide_layouts
        except Exception as e:
            logger.debug("访问母版 layouts 失败: %s", e)
            continue

        for layout_pos, layout in enumerate(layouts):
            # SlideLayout 在 python-pptx 1.0.2 无公开 .idx 属性，使用集合内序号作为 layout_idx
            layout_idx = layout_pos
            # 跨母版避免 idx 冲突覆盖（保留先出现的）
            if layout_idx in result:
                continue

            layout_info: dict[str, Any] = {
                "name": None,
                "placeholders": [],
            }
            try:
                layout_info["name"] = layout.name
            except Exception:
                pass

            try:
                placeholders = layout.placeholders
            except Exception as e:
                logger.debug("layout[%s] 占位符访问失败: %s", layout_idx, e)
                placeholders = []

            for ph in placeholders:
                try:
                    layout_info["placeholders"].append(_extract_placeholder_info(ph))
                except Exception as e:
                    logger.debug("layout[%s] 占位符提取异常: %s", layout_idx, e)

            result[layout_idx] = layout_info

    logger.debug("profile_layouts 完成，共提取 %d 个布局", len(result))
    return result
