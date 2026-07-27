"""3.2.4 性能基准测试

统计单份生成耗时、内存占用，建立性能基线，防止迭代中性能退化。
核心渲染性能基线：0.15~0.92 秒/份（v2.0.0 实测）。

标记为 slow：需实际渲染 PPTX，使用 `-m "not slow"` 跳过。
"""
import gc
import json
import time
import tracemalloc
from pathlib import Path

import pytest
from pptx import Presentation

from aippt.logger import logger
from ppt_scene_adapter import SceneAdapter
from ppt_renderer import PptRenderer

ROOT = Path(__file__).parent.parent

# 性能基线（秒）：v2.0.0 实测单份生成 0.15~0.92 秒
PERF_BASELINE_MIN_SEC = 0.10
PERF_BASELINE_MAX_SEC = 3.50  # 留 4x 余量，避免硬件差异导致误报
PERF_BASELINE_MEMORY_MB = 100  # 单份生成内存上限基线


def _load_business_data(scene: str) -> dict:
    """加载场景对应的业务数据"""
    biz_map = {
        "工作总结": "business_worksummary.json",
        "年终总结": "business_annual.json",
        "工作汇报": "business_report.json",
    }
    biz_file = ROOT / biz_map.get(scene, "business_worksummary.json")
    if not biz_file.exists():
        pytest.skip(f"业务数据文件不存在: {biz_file}")
    with open(biz_file, "r", encoding="utf-8") as f:
        return json.load(f)


def _find_template_for_scene(scene: str) -> tuple[str, str, str]:
    """为场景找到一个可用模板（template_id, pptx_path, meta_path）"""
    adapter = SceneAdapter()
    templates = adapter.list_templates(category=scene)
    if not templates:
        pytest.skip(f"场景 {scene} 无可用模板")
    t = templates[0]
    meta, _ = adapter.get_template_meta(template_id=t["template_id"])
    pptx_path = adapter.get_template_pptx(meta)
    return t["template_id"], pptx_path, meta.get("path", "")


# ==================== 单份渲染性能测试 ====================

@pytest.mark.slow
def test_single_render_performance():
    """单份生成耗时基线：应在 0.10~3.50 秒内完成"""
    scene = "工作总结"
    biz_data = _load_business_data(scene)
    template_id, pptx_path, meta_rel = _find_template_for_scene(scene)
    meta_path = ROOT / "models" / meta_rel if not Path(meta_rel).is_absolute() else Path(meta_rel)

    adapter = SceneAdapter()
    meta, _ = adapter.get_template_meta(template_id=template_id)
    slot_data = adapter.adapt(scene, biz_data, meta)

    renderer = PptRenderer(pptx_path, str(meta_path))
    out_path = str(ROOT / "perf_test_output.pptx")

    # 强制 GC 后计时
    gc.collect()
    t0 = time.perf_counter()
    renderer.render(slot_data, out_path, remove_copyright=True, auto_fit=True)
    elapsed = time.perf_counter() - t0

    logger.info("性能测试: 单份渲染耗时 %.3f 秒（基线 %.2f~%.2f）",
                elapsed, PERF_BASELINE_MIN_SEC, PERF_BASELINE_MAX_SEC)

    assert elapsed >= PERF_BASELINE_MIN_SEC * 0.5, f"耗时异常偏低: {elapsed:.3f}s"
    assert elapsed <= PERF_BASELINE_MAX_SEC, (
        f"性能退化: 单份渲染耗时 {elapsed:.3f}s 超过基线 {PERF_BASELINE_MAX_SEC}s"
    )

    # 清理
    try:
        Path(out_path).unlink()
    except Exception:
        pass


@pytest.mark.slow
def test_render_memory_usage():
    """单份生成内存占用基线：应低于 100MB"""
    scene = "工作总结"
    biz_data = _load_business_data(scene)
    template_id, pptx_path, meta_rel = _find_template_for_scene(scene)
    meta_path = ROOT / "models" / meta_rel if not Path(meta_rel).is_absolute() else Path(meta_rel)

    adapter = SceneAdapter()
    meta, _ = adapter.get_template_meta(template_id=template_id)
    slot_data = adapter.adapt(scene, biz_data, meta)

    renderer = PptRenderer(pptx_path, str(meta_path))
    out_path = str(ROOT / "perf_test_memory.pptx")

    gc.collect()
    tracemalloc.start()
    renderer.render(slot_data, out_path, remove_copyright=True, auto_fit=True)
    current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    peak_mb = peak / (1024 * 1024)
    logger.info("内存测试: 峰值占用 %.2f MB（基线 %d MB）", peak_mb, PERF_BASELINE_MEMORY_MB)

    assert peak_mb <= PERF_BASELINE_MEMORY_MB, (
        f"内存超标: 峰值 {peak_mb:.2f}MB 超过基线 {PERF_BASELINE_MEMORY_MB}MB"
    )

    try:
        Path(out_path).unlink()
    except Exception:
        pass


# ==================== 批量渲染性能测试 ====================

@pytest.mark.slow
def test_batch_render_performance():
    """批量渲染 5 份平均耗时基线：单份平均应低于 3.5 秒"""
    scene = "工作总结"
    biz_data = _load_business_data(scene)
    template_id, pptx_path, meta_rel = _find_template_for_scene(scene)
    meta_path = ROOT / "models" / meta_rel if not Path(meta_rel).is_absolute() else Path(meta_rel)

    adapter = SceneAdapter()
    meta, _ = adapter.get_template_meta(template_id=template_id)
    slot_data = adapter.adapt(scene, biz_data, meta)

    batch_count = 5
    gc.collect()
    t0 = time.perf_counter()
    for i in range(batch_count):
        renderer = PptRenderer(pptx_path, str(meta_path))
        out_path = str(ROOT / f"perf_batch_{i}.pptx")
        renderer.render(slot_data, out_path, remove_copyright=True, auto_fit=True)
    total_elapsed = time.perf_counter() - t0
    avg = total_elapsed / batch_count

    logger.info("批量性能: %d 份总耗时 %.3fs, 平均 %.3fs/份", batch_count, total_elapsed, avg)

    assert avg <= PERF_BASELINE_MAX_SEC, (
        f"批量性能退化: 平均 {avg:.3f}s/份 超过基线 {PERF_BASELINE_MAX_SEC}s"
    )

    # 清理
    for i in range(batch_count):
        try:
            (ROOT / f"perf_batch_{i}.pptx").unlink()
        except Exception:
            pass


# ==================== 校验性能测试 ====================

def test_validate_performance():
    """校验引擎性能：单次大纲校验应低于 100ms"""
    from aippt.validators import validate_outline

    outline = {
        "scene": "年终总结",
        "total_pages": 6,
        "pages": [
            {"page_id": 1, "page_type": "cover", "title": "2025年度工作总结", "subtitle": "复盘"},
            {"page_id": 2, "page_type": "catalog", "title": "目录", "items": ["业绩", "成果", "规划"]},
            {"page_id": 3, "page_type": "divider", "section_no": "01", "title": "业绩概览"},
            {"page_id": 4, "page_type": "kpi", "title": "核心指标",
             "kpi_items": [{"label": "用户量", "value": "128万", "trend": "+35%"}]},
            {"page_id": 5, "page_type": "numbered_list", "title": "成果",
             "items": [{"subtitle": "增长", "desc": "提升35%"}]},
            {"page_id": 6, "page_type": "ending", "title": "感谢", "subtitle": "欢迎指正"},
        ],
    }

    gc.collect()
    t0 = time.perf_counter()
    for _ in range(20):
        validate_outline(outline)
    elapsed = time.perf_counter() - t0
    avg_ms = (elapsed / 20) * 1000

    logger.info("校验性能: 平均 %.2f ms/次", avg_ms)
    assert avg_ms <= 100, f"校验性能退化: {avg_ms:.2f}ms 超过 100ms 基线"
