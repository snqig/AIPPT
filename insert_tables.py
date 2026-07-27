"""图表 / 表格动态替换测试工具

提供独立 CLI，传入 PPTX + 测试数据 JSON，快速验证图表数据源替换与表格动态行扩展效果，
输出渲染前后对比报告（结构化 JSON），便于调试与回归验证。

设计原则：
  - 不依赖完整 meta.json，直接扫描 PPTX 内首个图表/表格形状并替换
  - 100% 复用 PptRenderer._replace_chart_data / _fill_dynamic_table，保证与生产逻辑一致
  - 输出结构化 JSON，适配 opencode 模型调用
  - 支持单图表、单表格、批量三种模式

子命令：
  test-chart   : 替换首个图表的数据源
  test-table   : 替换首个表格的数据（含动态行扩展）
  test-all     : 同时测试图表与表格
  list         : 列出 PPTX 内所有图表/表格位置与类型

用法示例：
  python insert_tables.py test-chart --input template.pptx --data chart_data.json --output out.pptx
  python insert_tables.py test-table --input template.pptx --data table_data.json --output out.pptx
  python insert_tables.py list --input template.pptx
"""
import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any, Optional

from pptx import Presentation
from pptx.shapes.base import BaseShape

from aippt.logger import logger
from ppt_renderer import PptRenderer


# ==================== 扫描工具 ====================

def _iter_graphic_frames(prs: Presentation):
    """遍历所有幻灯片中的 GraphicFrame（图表/表格载体）

    python-pptx 1.0.2 中 MSO_SHAPE_TYPE 无 GRAPHIC_FRAME 枚举，
    直接通过 has_chart / has_table 属性识别图表与表格形状。
    """
    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            is_chart = hasattr(shape, "has_chart") and shape.has_chart
            is_table = hasattr(shape, "has_table") and shape.has_table
            if is_chart or is_table:
                yield slide_idx, shape


def list_charts_and_tables(pptx_path: Path) -> dict[str, Any]:
    """列出 PPTX 内所有图表与表格的位置、类型、当前数据规模

    :param pptx_path: PPTX 文件路径
    :return: {"charts": [...], "tables": [...], "total": int}
    """
    prs = Presentation(str(pptx_path))
    charts: list[dict[str, Any]] = []
    tables: list[dict[str, Any]] = []

    for slide_idx, shape in _iter_graphic_frames(prs):
        if hasattr(shape, "has_chart") and shape.has_chart:
            try:
                chart = shape.chart
                ct_str = str(chart.chart_type).upper()
                if "BAR" in ct_str or "COLUMN" in ct_str:
                    category = "bar"
                elif "LINE" in ct_str:
                    category = "line"
                elif "PIE" in ct_str:
                    category = "pie"
                elif "RADAR" in ct_str:
                    category = "radar"
                else:
                    category = "unknown"

                try:
                    plot = chart.plots[0]
                    series_count = len(list(plot.series))
                    cat_count = len(list(plot.categories))
                except Exception:
                    series_count = 0
                    cat_count = 0

                charts.append({
                    "slide": slide_idx,
                    "chart_type": category,
                    "raw_type": str(chart.chart_type),
                    "series_count": series_count,
                    "category_count": cat_count,
                })
            except Exception as e:
                charts.append({"slide": slide_idx, "error": str(e)})
        elif hasattr(shape, "has_table") and shape.has_table:
            try:
                table = shape.table
                tables.append({
                    "slide": slide_idx,
                    "rows": len(table.rows),
                    "cols": len(table.columns),
                })
            except Exception as e:
                tables.append({"slide": slide_idx, "error": str(e)})

    return {
        "charts": charts,
        "tables": tables,
        "total": len(charts) + len(tables),
    }


# ==================== 测试执行 ====================

def _build_temp_meta(template_path: Path, page_slots: dict[str, list[dict]]) -> Path:
    """为 PptRenderer 构造最小化临时 meta.json（仅含 page_slots）

    PptRenderer 初始化需要 meta.json，这里生成仅包含必要字段的临时文件。
    """
    import tempfile
    meta = {
        "template_id": f"insert_tables_test__{template_path.stem}",
        "category": "测试",
        "total_pages": 0,
        "chapters": [],
        "page_slots": page_slots,
        "removable_pages": [],
    }
    tmp = Path(tempfile.mkdtemp()) / "insert_tables_test.meta.json"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    return tmp


def run_chart_replacement(
    input_pptx: Path,
    chart_data: dict[str, Any],
    output_pptx: Path,
) -> dict[str, Any]:
    """测试图表数据源替换

    :param input_pptx: 输入 PPTX
    :param chart_data: {"categories": [...], "series": [{"name": "...", "data": [...]}]}
    :param output_pptx: 输出 PPTX
    :return: 对比报告
    """
    # 渲染前：扫描首个图表
    before = list_charts_and_tables(input_pptx)
    if not before["charts"]:
        return {
            "test_pass": False,
            "error": "输入 PPTX 中未找到图表形状",
            "before": before,
        }

    first_chart = before["charts"][0]
    target_slide = str(first_chart["slide"])

    # 构造临时 meta，让 PptRenderer 在目标页执行 chart_data 替换
    page_slots = {
        target_slide: [{"slot": "chart_data", "type": "chart"}]
    }
    tmp_meta = _build_temp_meta(input_pptx, page_slots)
    slot_data = {target_slide: {"chart_data": chart_data}}

    try:
        renderer = PptRenderer(str(input_pptx), str(tmp_meta))
        t0 = time.perf_counter()
        renderer.render(
            slot_data,
            str(output_pptx),
            remove_copyright=False,
            auto_fit=False,
            transitions=None,
            animations=None,
        )
        elapsed = round(time.perf_counter() - t0, 3)

        # 渲染后：重新扫描验证
        after = list_charts_and_tables(output_pptx)
        after_chart = after["charts"][0] if after["charts"] else {}

        return {
            "test_pass": True,
            "target": first_chart,
            "before": {
                "chart_type": first_chart.get("chart_type"),
                "series_count": first_chart.get("series_count"),
                "category_count": first_chart.get("category_count"),
            },
            "after": {
                "chart_type": after_chart.get("chart_type"),
                "series_count": after_chart.get("series_count"),
                "category_count": after_chart.get("category_count"),
            },
            "input_data": {
                "categories_count": len(chart_data.get("categories", [])),
                "series_count": len(chart_data.get("series", [])),
            },
            "elapsed_seconds": elapsed,
            "output": str(output_pptx),
        }
    except Exception as e:
        logger.exception("图表测试失败")
        return {
            "test_pass": False,
            "error": str(e),
            "before": before,
        }


def run_table_replacement(
    input_pptx: Path,
    table_data: dict[str, Any],
    output_pptx: Path,
) -> dict[str, Any]:
    """测试表格动态行扩展

    :param input_pptx: 输入 PPTX
    :param table_data: {"headers": [...], "rows": [[...], ...]}
    :param output_pptx: 输出 PPTX
    :return: 对比报告
    """
    before = list_charts_and_tables(input_pptx)
    if not before["tables"]:
        return {
            "test_pass": False,
            "error": "输入 PPTX 中未找到表格形状",
            "before": before,
        }

    first_table = before["tables"][0]
    target_slide = str(first_table["slide"])

    page_slots = {
        target_slide: [{"slot": "table_data", "type": "table"}]
    }
    tmp_meta = _build_temp_meta(input_pptx, page_slots)
    slot_data = {target_slide: {"table_data": table_data}}

    try:
        renderer = PptRenderer(str(input_pptx), str(tmp_meta))
        t0 = time.perf_counter()
        renderer.render(
            slot_data,
            str(output_pptx),
            remove_copyright=False,
            auto_fit=False,
            transitions=None,
            animations=None,
        )
        elapsed = round(time.perf_counter() - t0, 3)

        after = list_charts_and_tables(output_pptx)
        after_table = after["tables"][0] if after["tables"] else {}

        return {
            "test_pass": True,
            "target": first_table,
            "before": {
                "rows": first_table.get("rows"),
                "cols": first_table.get("cols"),
            },
            "after": {
                "rows": after_table.get("rows"),
                "cols": after_table.get("cols"),
            },
            "input_data": {
                "headers_count": len(table_data.get("headers", [])),
                "data_rows": len(table_data.get("rows", [])),
            },
            "rows_delta": (after_table.get("rows", 0) - first_table.get("rows", 0)),
            "elapsed_seconds": elapsed,
            "output": str(output_pptx),
        }
    except Exception as e:
        logger.exception("表格测试失败")
        return {
            "test_pass": False,
            "error": str(e),
            "before": before,
        }


# ==================== CLI ====================

def _load_data_json(path: Path) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="图表/表格动态替换测试工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list", help="列出 PPTX 内所有图表/表格")
    p_list.add_argument("--input", required=True, help="PPTX 文件路径")

    p_chart = sub.add_parser("test-chart", help="测试图表数据源替换")
    p_chart.add_argument("--input", required=True, help="输入 PPTX")
    p_chart.add_argument("--data", required=True, help="图表数据 JSON 文件")
    p_chart.add_argument("--output", required=True, help="输出 PPTX")

    p_table = sub.add_parser("test-table", help="测试表格动态行扩展")
    p_table.add_argument("--input", required=True, help="输入 PPTX")
    p_table.add_argument("--data", required=True, help="表格数据 JSON 文件")
    p_table.add_argument("--output", required=True, help="输出 PPTX")

    p_all = sub.add_parser("test-all", help="同时测试图表与表格")
    p_all.add_argument("--input", required=True, help="输入 PPTX")
    p_all.add_argument("--chart-data", help="图表数据 JSON")
    p_all.add_argument("--table-data", help="表格数据 JSON")
    p_all.add_argument("--output", required=True, help="输出 PPTX")

    args = parser.parse_args()

    if args.cmd == "list":
        result = list_charts_and_tables(Path(args.input))
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "test-chart":
        data = _load_data_json(Path(args.data))
        result = run_chart_replacement(Path(args.input), data, Path(args.output))
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get("test_pass") else 1

    if args.cmd == "test-table":
        data = _load_data_json(Path(args.data))
        result = run_table_replacement(Path(args.input), data, Path(args.output))
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get("test_pass") else 1

    if args.cmd == "test-all":
        report: dict[str, Any] = {"test_pass": True, "items": []}
        # 图表测试（中间产物，最终输出合并到最后）
        if args.chart_data:
            chart_data = _load_data_json(Path(args.chart_data))
            tmp_out = Path(args.output).with_suffix(".chart.tmp.pptx")
            r = run_chart_replacement(Path(args.input), chart_data, tmp_out)
            report["items"].append({"type": "chart", "result": r})
            if not r.get("test_pass"):
                report["test_pass"] = False
            next_input = tmp_out
        else:
            next_input = Path(args.input)

        if args.table_data:
            table_data = _load_data_json(Path(args.table_data))
            r = test_table_replacement(next_input, table_data, Path(args.output))
            report["items"].append({"type": "table", "result": r})
            if not r.get("test_pass"):
                report["test_pass"] = False

        # 清理临时文件
        tmp_chart = Path(args.output).with_suffix(".chart.tmp.pptx")
        if tmp_chart.exists():
            try:
                tmp_chart.unlink()
            except Exception:
                pass

        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if report.get("test_pass") else 1

    return 1


if __name__ == "__main__":
    sys.exit(main())
