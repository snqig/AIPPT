"""
templates_index.json 迁移脚本（P1-2.2.3）

为旧版索引补充新字段，保持 100% 向后兼容：
  - color_scheme:  色系（默认 "蓝色系"）
  - industry:      适用行业（默认 ["通用"]）
  - page_range:    页数范围描述（根据 total_pages 计算）
  - quality_score: 质量评分（默认 80）

用法：
  python migrate_index.py                       # 默认迁移 models/templates_index.json
  python migrate_index.py --index 自定义路径
  python migrate_index.py --dry-run             # 只打印不写入
"""
import argparse
import json
from pathlib import Path
from typing import Any

from aippt.logger import logger

# 默认值定义
DEFAULT_COLOR_SCHEME = "蓝色系"
DEFAULT_INDUSTRY = ["通用"]
DEFAULT_QUALITY_SCORE = 80


def calc_page_range(total_pages: int) -> str:
    """
    根据 total_pages 计算页数范围描述。

    规则：以 5 页为一档，向下取整作为下界，向上取整作为上界。
    示例：34 -> "30-35页"；21 -> "20-25页"；13 -> "10-15页"。
    """
    if not isinstance(total_pages, int) or total_pages <= 0:
        return "10-15页"
    lower = (total_pages - 1) // 5 * 5 + 1   # 1,6,11,16...
    upper = lower + 4
    return f"{lower}-{upper}页"


def migrate_entry(entry: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    """
    为单个模板条目补充新字段，返回 (更新后条目, 已补字段列表)。
    若字段已存在则保留原值（不覆盖）。
    """
    patched: list[str] = []
    if "color_scheme" not in entry:
        entry["color_scheme"] = DEFAULT_COLOR_SCHEME
        patched.append("color_scheme")
    if "industry" not in entry:
        entry["industry"] = list(DEFAULT_INDUSTRY)
        patched.append("industry")
    if "page_range" not in entry:
        entry["page_range"] = calc_page_range(entry.get("total_pages", 0))
        patched.append("page_range")
    if "quality_score" not in entry:
        entry["quality_score"] = DEFAULT_QUALITY_SCORE
        patched.append("quality_score")
    return entry, patched


def migrate_index(index_path: Path, dry_run: bool = False) -> dict[str, Any]:
    """
    迁移索引文件，返回统计信息。
    """
    if not index_path.exists():
        logger.error("索引文件不存在: %s", index_path)
        return {"ok": False, "error": "文件不存在"}

    with open(index_path, "r", encoding="utf-8") as f:
        index = json.load(f)

    templates = index.get("templates", [])
    total = len(templates)
    patched_count = 0
    field_stats: dict[str, int] = {}

    for entry in templates:
        _, patched = migrate_entry(entry)
        if patched:
            patched_count += 1
            for fld in patched:
                field_stats[fld] = field_stats.get(fld, 0) + 1

    logger.info("索引迁移：%d 个模板，%d 个被补字段", total, patched_count)
    for fld, cnt in field_stats.items():
        logger.info("  - %s: %d 条补充默认值", fld, cnt)

    if dry_run:
        logger.info("dry-run 模式：不写入文件")
    else:
        with open(index_path, "w", encoding="utf-8") as f:
            json.dump(index, f, ensure_ascii=False, indent=2)
        logger.info("已写回: %s", index_path)

    return {
        "ok": True,
        "total": total,
        "patched": patched_count,
        "field_stats": field_stats,
        "dry_run": dry_run,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="templates_index.json 迁移工具")
    parser.add_argument("--index", default="models/templates_index.json",
                        help="索引文件路径（默认: models/templates_index.json）")
    parser.add_argument("--dry-run", action="store_true", help="只打印不写入")
    args = parser.parse_args()

    index_path = Path(args.index)
    result = migrate_index(index_path, dry_run=args.dry_run)
    if not result.get("ok"):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
