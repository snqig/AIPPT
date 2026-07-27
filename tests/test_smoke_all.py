"""阶段四：全量模板冒烟测试（pytest 版本）

迁移自 smoke_test_all.py，对所有模板逐一执行渲染，统计成功率、页数偏差、异常情况。
标记为 slow：全量测试耗时较长，使用 `-m "not slow"` 跳过。
"""
import json
import time
from pathlib import Path

import pytest
from pptx import Presentation

from aippt.logger import logger
from ppt_scene_adapter import SceneAdapter

ROOT = Path(__file__).parent.parent

# 场景 → 业务数据文件
SCENE_DATA = {
    "工作总结": "business_worksummary.json",
    "个人简历": "business_resume.json",
    "自我介绍": "business_intro.json",
    "年终总结": "business_annual.json",
    "工作汇报": "business_report.json",
    "工作计划": "business_plan.json",
    "述职报告": "business_duty.json",
    "开题报告": "business_thesis.json",
    "公司简介": "business_company.json",
    "职业规划": "business_career.json",
}


@pytest.mark.slow
def test_smoke_all_templates(tmp_path):
    """全量模板冒烟测试：逐一渲染所有模板并校验页数

    保留原有 smoke_test_all.py 测试逻辑：
      - 遍历所有模板（adapter.list_templates）
      - 按场景匹配业务数据文件
      - 渲染并校验输出页数（与 removable_pages 之后的期望页数对比）
      - 统计成功/页数异常/失败/跳过
    """
    adapter = SceneAdapter()
    templates = adapter.list_templates()
    logger.info(f"全量冒烟测试：共 {len(templates)} 套模板")

    out_dir = tmp_path / "output_smoke"
    out_dir.mkdir(exist_ok=True)

    results = []
    total_time = 0.0

    for t in templates:
        tid = t["template_id"]
        category = t["category"]
        name = t["name"]
        total_pages = t["total_pages"]

        biz_file = SCENE_DATA.get(category)
        if not biz_file:
            results.append({"tid": tid, "category": category, "name": name,
                            "status": "skip", "reason": "无业务数据"})
            continue

        biz_path = ROOT / biz_file
        if not biz_path.exists():
            results.append({"tid": tid, "category": category, "name": name,
                            "status": "skip", "reason": f"业务数据文件不存在: {biz_file}"})
            continue

        with open(biz_path, "r", encoding="utf-8") as f:
            biz_data = json.load(f)

        out_path = str(out_dir / f"smoke_{category}_{name}.pptx")
        logger.info(f"渲染 [{category}] {name} ({total_pages}页)...")

        t0 = time.time()
        try:
            adapter.render(category, biz_data, template_id=tid, output_path=out_path)
            elapsed = time.time() - t0
            total_time += elapsed

            # 校验输出文件
            prs = Presentation(out_path)
            cur_pages = len(prs.slides)
            meta, _ = adapter.get_template_meta(path=t["path"])
            expected = total_pages - len(meta.get("removable_pages", []))
            page_ok = cur_pages == expected

            results.append({
                "tid": tid, "category": category, "name": name,
                "total_pages": total_pages, "cur_pages": cur_pages,
                "expected": expected, "page_ok": page_ok,
                "elapsed": elapsed, "status": "success",
            })
            logger.info(f"完成 {cur_pages}/{expected}页, {elapsed:.2f}s")
        except Exception as e:
            elapsed = time.time() - t0
            total_time += elapsed
            results.append({
                "tid": tid, "category": category, "name": name,
                "elapsed": elapsed, "status": f"failed: {str(e)[:50]}",
            })
            logger.error(f"失败 [{category}] {name}: {str(e)[:80]}")

    # 汇总
    success = sum(1 for r in results if r["status"] == "success" and r.get("page_ok"))
    page_warn = sum(1 for r in results if r["status"] == "success" and not r.get("page_ok"))
    failed = sum(1 for r in results if r["status"] != "success" and r["status"] != "skip")
    skipped = sum(1 for r in results if r["status"] == "skip")

    logger.info(
        f"成功 {success} / 页数异常 {page_warn} / 失败 {failed} / 跳过 {skipped} / 总耗时 {total_time:.2f}s"
    )

    # 断言：至少有结果，且无渲染失败
    assert len(results) > 0, "未收集到任何测试结果"
    assert failed == 0, (
        f"有 {failed} 套模板渲染失败: "
        f"{[r for r in results if r['status'] not in ('success', 'skip')]}"
    )
