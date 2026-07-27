"""
弹性约束计算工具库（T604）

功能：
    1. 文本高度预估：基于字号、字数、文本框宽度估算所需高度
    2. 纵向元素弹性均分：在固定容器内，按内容权重分配高度
    3. 自动换行支持：长文本按字符宽度估算换行点

设计原则：
    - 轻量纯函数，无外部依赖（仅用标准库）
    - 所有尺寸使用 EMU（python-pptx 原生单位）
    - 预估误差目标 ≤10%，配合 auto_fit 兜底防溢出
    - 可通过开关关闭，回退到静态等高模式

使用方式：
    from aippt.layout.elastic_constraint import (
        estimate_text_height, distribute_vertical, wrap_text
    )

    # 1. 预估文本所需高度
    h = estimate_text_height("一段较长文本...", font_size_pt=16, box_width_emu=...)

    # 2. 纵向弹性均分多个条目
    heights = distribute_vertical(
        container_h=Emu(4*914400),
        items=[{"text": "短文本"}, {"text": "很长很长的文本..."}],
        min_h_emu=Emu(0.5*914400),
    )
"""
from __future__ import annotations

from typing import Any

from pptx.util import Emu, Pt


# EMU 转换常量
EMU_PER_INCH = 914400
EMU_PER_PT = 12700  # 1 pt = 12700 EMU

# 字符宽度系数（相对字号）
# 中文字符宽度 ≈ 1.0 * font_size
# 英文字符宽度 ≈ 0.55 * font_size
# 混排时按字符比例加权
CN_CHAR_WIDTH_RATIO = 1.0
EN_CHAR_WIDTH_RATIO = 0.55

# 行高系数（相对字号）
LINE_HEIGHT_RATIO = 1.2

# 默认最小行高（EMU），防止条目被压扁
DEFAULT_MIN_ITEM_H_EMU = int(0.4 * EMU_PER_INCH)


def estimate_text_height(
    text: str,
    font_size_pt: float,
    box_width_emu: int,
    line_height_ratio: float = LINE_HEIGHT_RATIO,
) -> int:
    """估算文本在指定宽度下所需的高度（EMU）

    算法：
        1. 按字符类型（中文/英文）计算总字符宽度
        2. 除以行宽得到行数
        3. 行数 × 行高 = 总高度

    :param text: 文本内容
    :param font_size_pt: 字号 pt
    :param box_width_emu: 文本框宽度 EMU
    :param line_height_ratio: 行高系数（默认 1.2）
    :return: 估算高度 EMU
    """
    if not text:
        return 0

    font_size_emu = int(font_size_pt * EMU_PER_PT)
    box_width_pt = box_width_emu / EMU_PER_PT

    # 计算单字符宽度（pt）
    cn_char_w_pt = font_size_pt * CN_CHAR_WIDTH_RATIO
    en_char_w_pt = font_size_pt * EN_CHAR_WIDTH_RATIO

    # 累加字符宽度（区分中英文）
    total_w_pt = 0.0
    for ch in text:
        if _is_cjk(ch):
            total_w_pt += cn_char_w_pt
        elif ch in "\t":
            total_w_pt += en_char_w_pt * 4  # Tab 按 4 字符宽
        elif ch == "\n":
            total_w_pt = 0  # 强制换行，重置累计
            total_w_pt += 0  # 换行符本身不占宽
        else:
            total_w_pt += en_char_w_pt

    # 估算行数
    if total_w_pt <= 0:
        return 0
    char_w_per_line = box_width_pt
    if char_w_per_line <= 0:
        return font_size_emu * line_height_ratio  # 兜底单行

    # 显式换行符计数
    explicit_lines = text.count("\n") + 1
    # 按字符宽度估算的行数
    wrapped_lines = max(1, int(total_w_pt / char_w_per_line) + (1 if total_w_pt % char_w_per_line else 0))

    total_lines = max(explicit_lines, wrapped_lines)
    line_h_emu = font_size_emu * line_height_ratio
    return int(total_lines * line_h_emu)


def distribute_vertical(
    container_h_emu: int,
    items: list[dict[str, Any]],
    min_h_emu: int = DEFAULT_MIN_ITEM_H_EMU,
    gap_emu: int = 0,
    font_size_pt: float = 16.0,
    box_width_emu: int = 0,
) -> list[int]:
    """纵向弹性均分多个条目高度

    算法：
        1. 预估每个条目所需高度
        2. 若总高度 ≤ 容器高度，按预估高度分配，剩余空间均分给所有条目
        3. 若总高度 > 容器高度，按权重等比缩小（下限 min_h_emu）
        4. 条目数过多导致下限总和 > 容器高度时，全部按下限分配（允许溢出，由 auto_fit 兜底）

    :param container_h_emu: 容器总高度 EMU
    :param items: 条目列表，每个条目 dict 含 text 字段（用于预估高度）
    :param min_h_emu: 单条目最小高度 EMU
    :param gap_emu: 条目间距 EMU
    :param font_size_pt: 字号 pt（用于预估）
    :param box_width_emu: 文本框宽度 EMU（用于预估换行）
    :return: 每个条目的高度列表（EMU），长度等于 items
    """
    if not items:
        return []

    n = len(items)
    total_gap = gap_emu * (n - 1) if n > 1 else 0
    available_h = max(0, container_h_emu - total_gap)

    # 1. 预估每个条目高度
    estimated = []
    for item in items:
        text = str(item.get("text", ""))
        if box_width_emu > 0:
            h = estimate_text_height(text, font_size_pt, box_width_emu)
        else:
            # 无宽度信息时，按单行高度 + 字符数粗估
            h = int(font_size_pt * EMU_PER_PT * LINE_HEIGHT_RATIO)
            if len(text) > 20:
                h *= 2  # 长文本粗估 2 行
        # 加上下内边距（每条目上下各 0.1 inch）
        h += int(0.2 * EMU_PER_INCH)
        estimated.append(max(h, min_h_emu))

    total_estimated = sum(estimated)

    if total_estimated <= available_h:
        # 总高度小于容器，按预估分配，剩余空间均分
        surplus = available_h - total_estimated
        bonus_per_item = surplus // n
        return [h + bonus_per_item for h in estimated]
    else:
        # 总高度超过容器，按权重等比缩小
        scale = available_h / total_estimated
        scaled = [max(int(h * scale), min_h_emu) for h in estimated]
        # 检查下限总和是否溢出
        if sum(scaled) > available_h:
            # 下限总和仍溢出，全部按下限（允许溢出，由 auto_fit 兜底）
            return [min_h_emu] * n
        # 将剩余空间补偿给非下限条目
        deficit = available_h - sum(scaled)
        if deficit > 0:
            # 按原权重分配剩余
            weights = [h - min_h_emu for h in scaled]
            total_weight = sum(weights)
            if total_weight > 0:
                for i in range(n):
                    if weights[i] > 0:
                        scaled[i] += int(deficit * weights[i] / total_weight)
        return scaled


def wrap_text(
    text: str,
    font_size_pt: float,
    box_width_emu: int,
) -> list[str]:
    """按字符宽度估算换行点，返回分行后的字符串列表

    用于显示前的文本预换行，配合 estimate_text_height 使用。

    :param text: 原始文本
    :param font_size_pt: 字号 pt
    :param box_width_emu: 文本框宽度 EMU
    :return: 分行后的字符串列表（不含换行符）
    """
    if not text:
        return []

    box_width_pt = box_width_emu / EMU_PER_PT
    cn_w = font_size_pt * CN_CHAR_WIDTH_RATIO
    en_w = font_size_pt * EN_CHAR_WIDTH_RATIO

    lines: list[str] = []
    current_line = ""
    current_w = 0.0

    for ch in text:
        if ch == "\n":
            lines.append(current_line)
            current_line = ""
            current_w = 0.0
            continue

        ch_w = cn_w if _is_cjk(ch) else en_w
        if current_w + ch_w > box_width_pt and current_line:
            lines.append(current_line)
            current_line = ch
            current_w = ch_w
        else:
            current_line += ch
            current_w += ch_w

    if current_line:
        lines.append(current_line)

    return lines


def _is_cjk(ch: str) -> bool:
    """判断字符是否为 CJK 字符（中日韩统一表意文字 + 全角符号）

    :param ch: 单字符
    :return: True 为 CJK 字符
    """
    if not ch:
        return False
    cp = ord(ch)
    # CJK 统一表意文字
    if 0x4E00 <= cp <= 0x9FFF:
        return True
    # CJK 扩展 A
    if 0x3400 <= cp <= 0x4DBF:
        return True
    # CJK 兼容表意文字
    if 0xF900 <= cp <= 0xFAFF:
        return True
    # 全角标点符号
    if 0x3000 <= cp <= 0x303F:
        return True
    # 全角 ASCII / 半角片假名
    if 0xFF00 <= cp <= 0xFFEF:
        return True
    return False


# ==================== 便捷接口：弹性均分布局函数 ====================
def elastic_distribute_items(
    items: list[dict[str, Any]],
    area_top_emu: int,
    area_height_emu: int,
    area_width_emu: int,
    font_size_pt: float = 16.0,
    min_item_h_inch: float = 0.4,
    gap_inch: float = 0.15,
) -> list[tuple[int, int]]:
    """便捷接口：在指定区域内弹性均分条目，返回每个条目的 (top_emu, height_emu)

    :param items: 条目列表
    :param area_top_emu: 区域顶部 y（EMU）
    :param area_height_emu: 区域总高度 EMU
    :param area_width_emu: 区域宽度 EMU（用于换行预估）
    :param font_size_pt: 字号 pt
    :param min_item_h_inch: 单条目最小高度 inch
    :param gap_inch: 条目间距 inch
    :return: [(top_emu, height_emu), ...] 长度等于 items
    """
    min_h_emu = int(min_item_h_inch * EMU_PER_INCH)
    gap_emu = int(gap_inch * EMU_PER_INCH)

    heights = distribute_vertical(
        container_h_emu=area_height_emu,
        items=items,
        min_h_emu=min_h_emu,
        gap_emu=gap_emu,
        font_size_pt=font_size_pt,
        box_width_emu=area_width_emu,
    )

    result: list[tuple[int, int]] = []
    cur_top = area_top_emu
    for h in heights:
        result.append((cur_top, h))
        cur_top += h + gap_emu
    return result
