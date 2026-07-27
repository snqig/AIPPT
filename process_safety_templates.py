"""
安全教育模板批量处理工具
功能：
  1. 为 models/安全教育/ 下所有 .pptx 生成 meta.json（标记末尾2页为 removable_pages）
  2. 生成 2x2 多页缩略图
  3. 更新 preview_manifest.json 和 templates_index.json

用法：
  python process_safety_templates.py                    # 全量处理
  python process_safety_templates.py --meta-only        # 仅生成 meta
  python process_safety_templates.py --thumbnail-only   # 仅生成缩略图
  python process_safety_templates.py --force            # 覆盖已存在
"""
import argparse
import json
import sys
import time
from pathlib import Path
from collections import defaultdict

from ppt_meta_tool import generate_single_meta
from generate_thumbnails import generate_thumbnail
from import_templates import update_templates_index
from aippt.logger import logger


CATEGORY = "安全教育"
MODELS_DIR = Path("models")
SRC_DIR = MODELS_DIR / CATEGORY
REMOVABLE_TAIL = 2


def generate_all_meta(force=False):
    """为所有安全教育模板生成 meta.json，标记末尾2页为可删除"""
    pptx_files = sorted(SRC_DIR.glob("*.pptx"))
    logger.info("扫描到 %d 个安全教育模板", len(pptx_files))

    success = 0
    skipped = 0
    failures = []

    for idx, pptx in enumerate(pptx_files, 1):
        meta_path = pptx.with_suffix(".pptx.meta.json")
        # meta 文件名：xxx.pptx → xxx.pptx.meta.json
        # 但 generate_single_meta 内部用 pptx_path.stem，所以 meta 文件名是 xxx.meta.json
        # 需要确认命名规则
        meta_path = pptx.parent / f"{pptx.name}.meta.json"

        if meta_path.exists() and not force:
            skipped += 1
            if idx % 20 == 0:
                logger.info("[%d/%d] 已跳过 %d 个", idx, len(pptx_files), skipped)
            continue

        logger.info("[%d/%d] 生成 meta: %s", idx, len(pptx_files), pptx.name)

        meta_result = generate_single_meta(pptx, CATEGORY)
        if isinstance(meta_result, tuple):
            meta, meta_error = meta_result
        else:
            meta, meta_error = meta_result, None

        if meta is None:
            logger.error("  meta 生成失败: %s", meta_error)
            failures.append({"file": pptx.name, "error": f"meta: {meta_error}"})
            continue

        # 标记末尾 N 页为可删除
        total_p = meta.get("total_pages", 0)
        existing = meta.get("removable_pages", [])
        tail = list(range(total_p - REMOVABLE_TAIL + 1, total_p + 1))
        meta["removable_pages"] = sorted(set(existing + tail))

        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        success += 1
        if idx % 20 == 0:
            logger.info("  进度: %d/%d 成功, %d 跳过", success, idx, skipped)

    logger.info("meta 生成完成: 成功 %d, 跳过 %d, 失败 %d", success, skipped, len(failures))
    if failures:
        for fail in failures[:5]:
            logger.warning("  失败: %s - %s", fail["file"], fail["error"])
        if len(failures) > 5:
            logger.warning("  ... 还有 %d 个失败", len(failures) - 5)

    return success, skipped, failures


def generate_all_thumbnails(force=False):
    """为所有安全教育模板生成 2x2 多页缩略图"""
    pptx_files = sorted(SRC_DIR.glob("*.pptx"))
    logger.info("开始生成缩略图: %d 个模板", len(pptx_files))

    success = 0
    skipped = 0
    failures = []

    for idx, pptx in enumerate(pptx_files, 1):
        thumb_path = pptx.with_suffix(".pptx.png")

        if thumb_path.exists() and not force:
            skipped += 1
            if idx % 10 == 0:
                logger.info("[%d/%d] 已跳过 %d 个", idx, len(pptx_files), skipped)
            continue

        logger.info("[%d/%d] 生成缩略图: %s", idx, len(pptx_files), pptx.name)

        ok = generate_thumbnail(pptx, thumb_path, layout="2x2")
        if ok:
            success += 1
        else:
            failures.append(str(pptx))

        # 每10个打印进度
        if idx % 10 == 0:
            logger.info("  进度: %d/%d, 成功 %d, 跳过 %d", idx, len(pptx_files), success, skipped)

    logger.info("缩略图生成完成: 成功 %d, 跳过 %d, 失败 %d", success, skipped, len(failures))
    if failures:
        for f in failures[:5]:
            logger.warning("  失败: %s", f)
        if len(failures) > 5:
            logger.warning("  ... 还有 %d 个失败", len(failures) - 5)

    return success, skipped, failures


def update_preview_manifest():
    """更新 preview_manifest.json（合并安全教育条目）"""
    manifest_path = MODELS_DIR / "preview_manifest.json"

    # 加载现有 manifest
    existing = []
    if manifest_path.exists():
        with open(manifest_path, "r", encoding="utf-8") as f:
            existing = json.load(f)

    # 移除旧的安全教育条目
    existing = [e for e in existing if e.get("category") != CATEGORY]

    # 添加新的安全教育条目
    pptx_files = sorted(SRC_DIR.glob("*.pptx"))
    for pptx in pptx_files:
        meta_path = pptx.parent / f"{pptx.name}.meta.json"
        thumb_path = pptx.with_suffix(".pptx.png")

        meta = {}
        if meta_path.exists():
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)

        existing.append({
            "template_id": meta.get("template_id", ""),
            "category": CATEGORY,
            "name": pptx.stem,
            "source_file": pptx.name,
            "pptx_path": str(pptx.relative_to(MODELS_DIR)).replace("\\", "/"),
            "meta_path": str(meta_path.relative_to(MODELS_DIR)).replace("\\", "/"),
            "screenshot": str(thumb_path.relative_to(MODELS_DIR)).replace("\\", "/") if thumb_path.exists() else None,
            "total_pages": meta.get("total_pages", 0),
            "matched_keyword": "安全教育",
        })

    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    logger.info("preview_manifest.json 已更新: 共 %d 条", len(existing))


def main():
    parser = argparse.ArgumentParser(description="安全教育模板批量处理工具")
    parser.add_argument("--meta-only", action="store_true", help="仅生成 meta.json")
    parser.add_argument("--thumbnail-only", action="store_true", help="仅生成缩略图")
    parser.add_argument("--force", action="store_true", help="覆盖已存在的文件")
    args = parser.parse_args()

    do_meta = not args.thumbnail_only
    do_thumb = not args.meta_only

    if do_meta:
        logger.info("=" * 60)
        logger.info("Step 1: 生成 meta.json（标记末尾 %d 页可删除）", REMOVABLE_TAIL)
        logger.info("=" * 60)
        generate_all_meta(force=args.force)

    if do_thumb:
        logger.info("=" * 60)
        logger.info("Step 2: 生成 2x2 多页缩略图")
        logger.info("=" * 60)
        generate_all_thumbnails(force=args.force)

    logger.info("=" * 60)
    logger.info("Step 3: 更新索引")
    logger.info("=" * 60)
    update_preview_manifest()
    update_templates_index(str(MODELS_DIR))

    logger.info("=" * 60)
    logger.info("全部完成！")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
