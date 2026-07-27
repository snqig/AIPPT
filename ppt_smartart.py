"""
PPT SmartArt 文本替换模块
功能：通过 dgm 命名空间操作 SmartArt XML，替换数据模型中的文本节点
依赖：python-pptx + lxml

参考：ECMA-376 DrawingML Diagram（dgm 命名空间）
      http://schemas.openxmlformats.org/drawingml/2006/diagram

设计说明：
- SmartArt 的文本存储在独立的 diagramData part（diagramData*.xml）中，
  通过 slide 上 graphicFrame 的 dgm:relIds/@r:dm 关系定位。
- 数据模型根节点为 <dgm:sldData>，其下 <dgm:pt> 节点承载文本：
    · 任务规范：dgm:pt 含 val 文本属性 —— 直接替换 @val
    · 真实文件：文本多在 dgm:pt/dgm:t/a:txBody/.../a:t —— 同步替换 a:t
  两条路径均覆盖，保证对真实 SmartArt 与任务描述结构都生效。
- 仅修改文本，保留 SmartArt 结构、布局、配色、形状层级。
"""
from lxml import etree

from aippt.logger import logger


# ==================== 命名空间常量 ====================
NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def list_smartart_text(slide):
    """
    列出 slide 中所有 SmartArt 文本节点内容，便于调试

    遍历 slide 上所有含 dgm 命名空间的 GraphicFrame，提取其数据模型中的
    文本节点（dgm:pt/@val 及 dgm:pt 下 a:t 文本）。

    :param slide: python-pptx 的 Slide 对象
    :return: list[str] 所有 SmartArt 文本节点的文本内容
    """
    texts: list[str] = []
    for root, _part in _iter_smartart_data_roots(slide):
        if root is None:
            continue
        for pt in root.iter("{%s}pt" % NS_DGM):
            # 路径1：dgm:pt 的 val 属性
            val = pt.get("val")
            if val and val.strip():
                texts.append(val)
            # 路径2：dgm:pt 下层 a:t 元素文本（真实 SmartArt 常见结构）
            for t_elem in pt.iter("{%s}t" % NS_A):
                if t_elem.text and t_elem.text.strip():
                    texts.append(t_elem.text)
    return texts


def replace_smartart_text(slide, replacements: dict):
    """
    替换 slide 中 SmartArt 数据模型的文本节点

    按 replacements 字典 {old_text: new_text} 替换文本节点：
      - dgm:pt/@val 属性值精确匹配则替换
      - dgm:pt 下层 a:t 元素文本精确匹配则替换
    保留 SmartArt 结构、布局、配色、形状层级。

    :param slide: python-pptx 的 Slide 对象
    :param replacements: dict {old_text: new_text}
    :return: int 实际替换的文本节点数
    """
    if not replacements:
        return 0

    replaced_count = 0
    for root, part in _iter_smartart_data_roots(slide):
        if root is None:
            continue

        # 路径1：dgm:pt 的 val 属性
        for pt in root.iter("{%s}pt" % NS_DGM):
            val = pt.get("val")
            if val is not None and val in replacements:
                new_val = replacements[val]
                pt.set("val", str(new_val))
                replaced_count += 1
                logger.info("SmartArt @val 替换: %r → %r", val, new_val)

        # 路径2：dgm:pt 下层 a:t 元素文本（真实 SmartArt 结构）
        for t_elem in root.iter("{%s}t" % NS_A):
            old_text = t_elem.text
            if old_text is not None and old_text in replacements:
                new_text = replacements[old_text]
                t_elem.text = str(new_text)
                replaced_count += 1
                logger.info("SmartArt a:t 替换: %r → %r", old_text, new_text)

        # 写回 part（XmlPart 修改 _element 即自动序列化；普通 Part 写回 _blob）
        _save_part_xml(part, root)

    if replaced_count == 0:
        logger.debug("SmartArt 未匹配到任何待替换文本（提供 %d 个 key）", len(replacements))
    else:
        logger.info("SmartArt 共替换 %d 个文本节点", replaced_count)
    return replaced_count


# ==================== 内部工具函数 ====================
def _iter_graphic_frames(slide):
    """
    遍历 slide 中所有 GraphicFrame（含组合内嵌套）

    GraphicFrame 是承载表格、图表、SmartArt 的形状容器，
    其 XML 元素 localname 为 "graphicFrame"。
    """
    shapes = slide.shapes
    for shape in shapes:
        el = getattr(shape, "_element", None)
        if el is not None and etree.QName(el).localname == "graphicFrame":
            yield shape
        # 递归处理组合形状（GROUP）
        if getattr(shape, "shape_type", None) == 6:  # MSO_SHAPE_TYPE.GROUP
            try:
                yield from _iter_graphic_frames(shape)
            except Exception:
                # 部分形状不支持 .shapes 访问，忽略
                pass


def _has_dgm_namespace(el):
    """检查元素树中是否含 dgm 命名空间节点"""
    for child in el.iter():
        if child.tag.startswith("{%s}" % NS_DGM):
            return True
    return False


def _iter_smartart_data_roots(slide):
    """
    遍历 slide 中所有 SmartArt，yield 每个数据模型的 (root_element, part)

    流程：
      1. 遍历 graphicFrame，检查是否含 dgm 命名空间
      2. 从 dgm:relIds 提取 @r:dm 关系 ID，定位 diagramData part
      3. 兜底：若 relIds 缺失，扫描 slide 关系中所有 diagramData 类型 part
    """
    seen_parts: set[int] = set()

    # 主路径：从 graphicFrame 的 dgm:relIds 定位
    has_smartart_frame = False
    for shape in _iter_graphic_frames(slide):
        el = shape._element
        if not _has_dgm_namespace(el):
            continue
        has_smartart_frame = True

        rel_ids = el.find(".//{%s}relIds" % NS_DGM)
        dm_rid = None
        if rel_ids is not None:
            dm_rid = rel_ids.get("{%s}dm" % NS_R)

        if dm_rid:
            part = _get_related_part(slide, dm_rid)
            if part is not None and id(part) not in seen_parts:
                seen_parts.add(id(part))
                root = _get_part_root(part)
                if root is not None:
                    yield root, part

    # 兜底路径：无 relIds 或无 SmartArt frame 时，扫描 slide 的 diagramData 关系
    if not has_smartart_frame or not seen_parts:
        try:
            rels = slide.part.rels
        except Exception:
            rels = {}
        for rel in rels.values():
            if getattr(rel, "is_external", False):
                continue
            if rel.reltype and rel.reltype.endswith("/diagramData"):
                part = rel.target_part
                if id(part) not in seen_parts:
                    seen_parts.add(id(part))
                    root = _get_part_root(part)
                    if root is not None:
                        yield root, part


def _get_related_part(slide, rId):
    """通过 rId 获取 slide 关联的 part，兼容不同 python-pptx 版本"""
    try:
        return slide.part.part_related_by(rId)
    except AttributeError:
        try:
            return slide.part.related_part(rId)
        except Exception:
            return None
    except Exception:
        return None


def _get_part_root(part):
    """
    获取 part 的 XML 根元素：
      - XmlPart：返回已解析的 _element（修改后保存时自动序列化）
      - 普通 Part：解析 blob 返回新树（需调用 _save_part_xml 写回）
    """
    elem = getattr(part, "_element", None)
    if elem is not None:
        return elem
    blob = getattr(part, "blob", None)
    if not blob:
        return None
    try:
        return etree.fromstring(blob)
    except Exception as e:
        logger.warning("解析 SmartArt 数据 part XML 失败: %s", e)
        return None


def _save_part_xml(part, root):
    """
    将修改后的 XML 写回 part

    - XmlPart：_element 已原地修改，无需额外操作（保存时自动序列化）
    - 普通 Part：序列化后写回 _blob
    """
    # XmlPart 的 _element 已被原地修改，跳过
    if getattr(part, "_element", None) is not None:
        return
    try:
        new_xml = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        part._blob = new_xml
    except Exception as e:
        logger.warning("写回 SmartArt 数据 part 失败: %s", e)


if __name__ == "__main__":
    # 自测：打印模块导出的公开函数
    print("replace_smartart_text:", hasattr(__import__("ppt_smartart"), "replace_smartart_text"))
    print("list_smartart_text:", hasattr(__import__("ppt_smartart"), "list_smartart_text"))
