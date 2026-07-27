"""
设计稿解析器原型（T501 辅助模块）

功能：
    使用计算机视觉（OpenCV/Pillow）分析设计稿图片，自动提取设计令牌。
    用于将任意设计稿（不限于 guizang-ppt-skill）注入主题系统。

支持的提取能力：
    1. 配色方案：主色调、辅助色、文本颜色、背景色（K-means 聚类）
    2. 字体大小预估：通过文字区域检测预估字号层级
    3. 间距规范：测量元素间距、对齐 8px 刻度原则

使用场景：
    - 当设计稿来源不是 guizang-ppt-skill（无明文令牌）时使用
    - 当需要从截图反推设计规范时使用
    - 与 design_tokens.py（明文抽取）互补

设计约束：
    - 输入：单页或多页设计稿图片（PNG/JPG）
    - 输出：初步的设计令牌字典（可能需人工微调）
    - 准确率：规范设计稿 ≥80%，复杂设计稿 60-80%
"""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Any, Optional

import cv2
import numpy as np
from PIL import Image

from aippt.logger import logger


# ==================== 配色提取 ====================
def extract_color_palette(
    image_path: str,
    num_colors: int = 8,
    min_saturation: float = 0.1,
    min_value: float = 0.15,
) -> dict[str, Any]:
    """从图片提取配色方案（K-means 聚类）

    提取策略：
        1. 图片转 HSV 色彩空间
        2. 过滤过低饱和度/亮度的像素（避免提取大量灰阶）
        3. K-means 聚类找出主要色簇
        4. 按频率排序，区分主色/辅助色/背景色

    :param image_path: 图片路径
    :param num_colors: 提取颜色数量（默认 8）
    :param min_saturation: 最小饱和度阈值（0-1）
    :param min_value: 最小亮度阈值（0-1）
    :return: 配色字典
        - primary: 主色 hex
        - secondary: 辅助色 hex
        - background: 背景色 hex（出现频率最高）
        - text_primary: 主文字色 hex（最暗的非饱和色）
        - text_secondary: 次要文字色 hex
        - all_colors: 全部色簇列表 [{hex, rgb, ratio}]
    """
    img = cv2.imread(image_path)
    if img is None:
        raise FileNotFoundError(f"无法读取图片: {image_path}")

    # BGR → HSV
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)

    # 过滤过低饱和度/亮度的像素
    sat_mask = hsv[:, :, 1] / 255.0 >= min_saturation
    val_mask = hsv[:, :, 2] / 255.0 >= min_value
    mask = sat_mask & val_mask

    # 收集彩色像素
    colored_pixels = img[mask]
    if len(colored_pixels) < 100:
        # 彩色像素过少，回退到全图分析
        colored_pixels = img.reshape(-1, 3)

    # K-means 聚类
    num_clusters = min(num_colors, len(colored_pixels))
    pixels_float = np.float32(colored_pixels)
    _, labels, centers = cv2.kmeans(
        pixels_float, num_clusters, None,
        (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 20, 1.0),
        3, cv2.KMEANS_PP_CENTERS,
    )

    # 统计每个色簇占比
    label_counts = Counter(labels.flatten())
    total = sum(label_counts.values())

    clusters: list[dict[str, Any]] = []
    for label, count in label_counts.most_common():
        bgr = centers[label]
        rgb = (int(bgr[2]), int(bgr[1]), int(bgr[0]))  # BGR → RGB
        hex_color = "#{:02X}{:02X}{:02X}".format(*rgb)
        ratio = count / total
        # 计算亮度和饱和度（BGR → HSV）
        bgr_pixel = np.uint8([[bgr]])
        hsv_pixel = cv2.cvtColor(bgr_pixel, cv2.COLOR_BGR2HSV)[0, 0]
        h, s, v = int(hsv_pixel[0]), int(hsv_pixel[1]), int(hsv_pixel[2])
        clusters.append({
            "hex": hex_color,
            "rgb": rgb,
            "ratio": round(ratio, 4),
            "hsv": (h, s, v),
            "brightness": v,  # 亮度 0-255
            "saturation": s,  # 饱和度 0-255
        })

    # 分类：背景色（占比最高）、主色（饱和度最高的彩色）、文字色（最暗）
    background = clusters[0]["hex"] if clusters else "#FFFFFF"

    # 主色：饱和度 > 50 且占比前 3
    saturated = [c for c in clusters if c["saturation"] > 50][:3]
    primary = saturated[0]["hex"] if saturated else (
        clusters[1]["hex"] if len(clusters) > 1 else "#000000"
    )
    secondary = saturated[1]["hex"] if len(saturated) > 1 else (
        clusters[2]["hex"] if len(clusters) > 2 else "#666666"
    ) if len(clusters) > 2 else "#666666"

    # 文字色：最暗的颜色
    darkest = min(clusters, key=lambda c: c["brightness"]) if clusters else None
    text_primary = darkest["hex"] if darkest and darkest["brightness"] < 100 else "#1F2937"

    # 次要文字色：中等亮度
    mid_brightness = [c for c in clusters if 80 < c["brightness"] < 180]
    text_secondary = mid_brightness[0]["hex"] if mid_brightness else "#6B7280"

    return {
        "primary": primary,
        "secondary": secondary,
        "background": background,
        "text_primary": text_primary,
        "text_secondary": text_secondary,
        "all_colors": clusters,
    }


# ==================== 字体大小预估 ====================
def estimate_font_sizes(
    image_path: str,
    min_area: int = 100,
    max_areas: int = 20,
) -> dict[str, Any]:
    """预估设计稿中的字体大小层级

    实现思路：
        1. 转灰度图，二值化
        2. 使用形态学操作连接文字行
        3. findContours 找文字区域
        4. 按区域高度聚类，识别字号层级

    :param image_path: 图片路径
    :param min_area: 最小文字区域面积（像素）
    :param max_areas: 最多取前 N 个大区域
    :return: 字号层级字典
        - levels: [{level, height_px, height_pt, role_guess, sample_count}]
        - min_font_pt: 最小字号（pt）
        - max_font_pt: 最大字号（pt）
    """
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    if img is None:
        raise FileNotFoundError(f"无法读取图片: {image_path}")

    # 二值化（自适应阈值）
    binary = cv2.adaptiveThreshold(
        img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV, 11, 2,
    )

    # 形态学操作：水平膨胀连接文字行
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (30, 1))
    closed = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel_h)

    # 查找轮廓
    contours, _ = cv2.findContours(
        closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE,
    )

    # 过滤过小区域，按面积排序取前 N
    areas = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        area = w * h
        if area >= min_area and h >= 8:  # 高度至少 8px
            areas.append({"x": x, "y": y, "w": w, "h": h, "area": area})

    areas.sort(key=lambda a: a["area"], reverse=True)
    areas = areas[:max_areas]

    if not areas:
        return {"levels": [], "min_font_pt": 0, "max_font_pt": 0}

    # 按高度聚类（10px 精度）
    height_buckets: dict[int, list[dict]] = {}
    for a in areas:
        bucket = (a["h"] // 10) * 10
        height_buckets.setdefault(bucket, []).append(a)

    # 按高度降序排列，分配角色
    sorted_buckets = sorted(height_buckets.items(), key=lambda x: -x[0])
    levels: list[dict[str, Any]] = []
    role_guesses = ["display", "h1", "h2", "h3", "body", "meta", "kicker"]

    for idx, (height_px, samples) in enumerate(sorted_buckets):
        # 像素高度 → pt（96 DPI）
        height_pt = round(height_px / 96 * 72, 1)
        role = role_guesses[idx] if idx < len(role_guesses) else f"level_{idx}"
        levels.append({
            "level": idx + 1,
            "height_px": height_px,
            "height_pt": height_pt,
            "role_guess": role,
            "sample_count": len(samples),
        })

    heights_pt = [lv["height_pt"] for lv in levels]
    return {
        "levels": levels,
        "min_font_pt": min(heights_pt) if heights_pt else 0,
        "max_font_pt": max(heights_pt) if heights_pt else 0,
    }


# ==================== 间距测量 ====================
def measure_spacing(
    image_path: str,
    grid_size: int = 8,
) -> dict[str, Any]:
    """测量设计稿中的间距规范

    实现思路：
        1. 检测水平/垂直空白带（连续白色行/列）
        2. 统计空白带宽度分布
        3. 对齐 grid_size（默认 8px）刻度

    :param image_path: 图片路径
    :param grid_size: 栅格刻度（像素），默认 8
    :return: 间距字典
        - horizontal_gaps: 水平间距分布 [{width_px, width_inch, count}]
        - vertical_gaps: 垂直间距分布
        - dominant_gap_px: 主导间距（像素）
        - safe_margin_inch: 估算安全边距（英寸）
    """
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    if img is None:
        raise FileNotFoundError(f"无法读取图片: {image_path}")

    h, w = img.shape

    # 二值化（>240 视为空白）
    _, binary = cv2.threshold(img, 240, 255, cv2.THRESH_BINARY)

    # 水平空白带：每行全白
    row_sums = np.sum(binary == 255, axis=1)
    blank_rows = row_sums == w

    # 垂直空白带：每列全白
    col_sums = np.sum(binary == 255, axis=0)
    blank_cols = col_sums == h

    # 统计水平空白带宽度
    def measure_blank_runs(blank_mask: np.ndarray) -> list[int]:
        runs: list[int] = []
        current = 0
        for v in blank_mask:
            if v:
                current += 1
            else:
                if current > 0:
                    runs.append(current)
                current = 0
        if current > 0:
            runs.append(current)
        return runs

    h_runs = measure_blank_runs(blank_rows)
    v_runs = measure_blank_runs(blank_cols)

    # 过滤过小间距（< 4px 视为字间距），对齐 8px 刻度
    def bucket_runs(runs: list[int], grid: int) -> list[dict[str, Any]]:
        if not runs:
            return []
        filtered = [r for r in runs if r >= 4]
        if not filtered:
            return []
        # 对齐 grid 刻度
        bucketed: dict[int, int] = {}
        for r in filtered:
            bucket = max(grid, (r // grid) * grid)
            bucketed[bucket] = bucketed.get(bucket, 0) + 1
        # 按频率排序
        sorted_buckets = sorted(bucketed.items(), key=lambda x: -x[1])
        return [
            {
                "width_px": bw,
                "width_inch": round(bw / 96, 3),
                "count": cnt,
            }
            for bw, cnt in sorted_buckets[:8]
        ]

    h_gaps = bucket_runs(h_runs, grid_size)
    v_gaps = bucket_runs(v_runs, grid_size)

    # 主导间距：频率最高的
    dominant_gap_px = h_gaps[0]["width_px"] if h_gaps else grid_size

    # 估算安全边距：上下左右边缘空白带宽度
    top_margin = 0
    for v in blank_rows:
        if v:
            top_margin += 1
        else:
            break
    bottom_margin = 0
    for v in reversed(blank_rows):
        if v:
            bottom_margin += 1
        else:
            break
    left_margin = 0
    for v in blank_cols:
        if v:
            left_margin += 1
        else:
            break
    right_margin = 0
    for v in reversed(blank_cols):
        if v:
            right_margin += 1
        else:
            break

    safe_margin_inch = round(
        np.mean([top_margin, bottom_margin, left_margin, right_margin]) / 96, 3
    )

    return {
        "horizontal_gaps": h_gaps,
        "vertical_gaps": v_gaps,
        "dominant_gap_px": dominant_gap_px,
        "dominant_gap_inch": round(dominant_gap_px / 96, 3),
        "safe_margin_inch": safe_margin_inch,
        "margins_px": {
            "top": top_margin,
            "bottom": bottom_margin,
            "left": left_margin,
            "right": right_margin,
        },
    }


# ==================== 综合解析入口 ====================
def parse_design_image(image_path: str) -> dict[str, Any]:
    """综合解析设计稿图片，提取完整设计令牌

    一次性调用配色提取、字体预估、间距测量，返回统一字典。

    :param image_path: 设计稿图片路径
    :return: 完整设计令牌字典
        - source: 图片路径
        - colors: 配色方案
        - fonts: 字号层级
        - spacing: 间距规范
        - warnings: 解析警告列表
    """
    path = Path(image_path)
    if not path.exists():
        raise FileNotFoundError(f"设计稿图片不存在: {image_path}")

    warnings: list[str] = []

    try:
        colors = extract_color_palette(image_path)
    except Exception as e:
        logger.warning("配色提取失败: %s", e)
        colors = {}
        warnings.append(f"配色提取失败: {e}")

    try:
        fonts = estimate_font_sizes(image_path)
    except Exception as e:
        logger.warning("字体预估失败: %s", e)
        fonts = {"levels": [], "min_font_pt": 0, "max_font_pt": 0}
        warnings.append(f"字体预估失败: {e}")

    try:
        spacing = measure_spacing(image_path)
    except Exception as e:
        logger.warning("间距测量失败: %s", e)
        spacing = {}
        warnings.append(f"间距测量失败: {e}")

    return {
        "source": str(path),
        "image_size": Image.open(image_path).size,
        "colors": colors,
        "fonts": fonts,
        "spacing": spacing,
        "warnings": warnings,
    }


def parse_design_images_to_json(
    image_paths: list[str],
    output_path: str,
) -> dict[str, Any]:
    """批量解析多张设计稿，合并为统一令牌 JSON

    多图合并策略：
        - 配色：取所有图的主色众数
        - 字号：合并所有字号层级，重新聚类
        - 间距：取所有图的主导间距众数

    :param image_paths: 设计稿图片路径列表
    :param output_path: 输出 JSON 路径
    :return: 合并后的设计令牌字典
    """
    all_results = [parse_design_image(p) for p in image_paths]

    # 合并配色：取第一张图作为基础（通常是封面）
    merged_colors = all_results[0]["colors"] if all_results else {}

    # 合并字号层级：取所有图的层级并集
    all_levels: list[dict] = []
    for r in all_results:
        all_levels.extend(r.get("fonts", {}).get("levels", []))

    # 合并间距：取所有图的主导间距众数
    dominant_gaps = [
        r.get("spacing", {}).get("dominant_gap_inch", 0.2)
        for r in all_results
    ]
    merged_dominant_gap = max(set(dominant_gaps), key=dominant_gaps.count) if dominant_gaps else 0.2

    merged = {
        "sources": [r["source"] for r in all_results],
        "colors": merged_colors,
        "fonts": {
            "levels": all_levels,
            "min_font_pt": min(
                (r.get("fonts", {}).get("min_font_pt", 0) for r in all_results),
                default=0,
            ),
            "max_font_pt": max(
                (r.get("fonts", {}).get("max_font_pt", 0) for r in all_results),
                default=0,
            ),
        },
        "spacing": {
            "dominant_gap_inch": merged_dominant_gap,
            "safe_margin_inch": all_results[0].get("spacing", {}).get("safe_margin_inch", 0.5)
            if all_results else 0.5,
        },
        "warnings": [w for r in all_results for w in r.get("warnings", [])],
    }

    Path(output_path).write_text(
        json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8",
    )
    logger.info("设计稿解析完成: %s（%d 张图片）", output_path, len(image_paths))
    return merged
