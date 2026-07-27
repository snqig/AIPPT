"""阶段四：边界场景测试（pytest 版本）

迁移自 edge_test.py，测试长文本溢出、空字段、多内容项超过槽位、缺字段等异常场景。
标记为 slow：边界场景需实际渲染 PPTX，耗时较长，使用 `-m "not slow"` 跳过。

每个边界场景对应一个独立测试函数，保留原有 edge_test.py 的用例逻辑。
"""
import copy
import json
from pathlib import Path

import pytest
from pptx import Presentation

from aippt.logger import logger
from ppt_scene_adapter import SceneAdapter

ROOT = Path(__file__).parent.parent


@pytest.fixture(scope="module")
def base_data():
    """加载基础业务数据（business_worksummary.json），供各边界用例深拷贝后修改"""
    biz_path = ROOT / "business_worksummary.json"
    with open(biz_path, "r", encoding="utf-8") as f:
        return json.load(f)


@pytest.fixture(scope="module")
def adapter():
    """共享 SceneAdapter 实例（模块级复用，避免重复初始化）"""
    return SceneAdapter()


def _render_and_check(adapter, data, out_path: str, case_name: str):
    """渲染并校验输出文件存在（辅助函数）

    :param adapter: SceneAdapter 实例
    :param data: 业务数据
    :param out_path: 输出 pptx 路径
    :param case_name: 用例名称（用于日志）
    """
    logger.info(f"渲染边界用例: {case_name}")
    adapter.render("工作总结", data, output_path=out_path)
    assert Path(out_path).exists(), f"用例 {case_name} 渲染后应生成 pptx 文件"
    prs = Presentation(out_path)
    logger.info(f"用例 {case_name} 生成成功, {len(prs.slides)}页")


# ==================== 用例1: 长文本溢出 ====================

@pytest.mark.slow
def test_long_text_overflow(base_data, adapter, tmp_path):
    """用例1: 长文本溢出测试（超长标题/描述）"""
    data = copy.deepcopy(base_data)
    # 超长标题（200字）
    data["sections"]["work_content"][0]["title"] = (
        "核心产品迭代升级完成V2.0版本开发用户活跃度同比增长35%"
        "新增注册用户12万供应链优化采购成本降低15%库存周转率提升22%"
    ) * 2
    # 超长描述（900字）
    data["sections"]["work_content"][0]["desc"] = (
        "本期完成核心产品V2.0版本全量开发与发布，涵盖用户增长、供应链优化、"
        "客户拓展、市场推广、团队建设、客户满意、运营优化、品牌升级等八大模块"
    ) * 10
    out_path = str(tmp_path / "edge_long_text.pptx")
    _render_and_check(adapter, data, out_path, "long_text_overflow")


# ==================== 用例2: 空字段 ====================

@pytest.mark.slow
def test_empty_data(base_data, adapter, tmp_path):
    """用例2: 空字段测试（cover/sections/end 均为空对象）"""
    data = copy.deepcopy(base_data)
    data["cover"] = {}
    data["sections"] = {}
    data["end"] = {}
    out_path = str(tmp_path / "edge_empty_data.pptx")
    _render_and_check(adapter, data, out_path, "empty_data")


# ==================== 用例3: 部分字段缺失 ====================

@pytest.mark.slow
def test_partial_fields(base_data, adapter, tmp_path):
    """用例3: 部分字段缺失测试（仅保留 cover 和首个 section，删除 end）"""
    data = copy.deepcopy(base_data)
    data["sections"] = {"work_content": base_data["sections"]["work_content"]}
    data.pop("end", None)
    out_path = str(tmp_path / "edge_partial_fields.pptx")
    _render_and_check(adapter, data, out_path, "partial_fields")


# ==================== 用例4: 内容项超过槽位 ====================

@pytest.mark.slow
def test_overflow_items(base_data, adapter, tmp_path):
    """用例4: 内容项超过槽位测试（每 section 填充 20 个 item，远超模板槽位）"""
    data = copy.deepcopy(base_data)
    for key in data["sections"]:
        data["sections"][key] = [
            {"title": f"项目{i + 1}", "desc": f"这是第{i + 1}个项目的工作内容描述，包含详细的工作成果与数据分析。"}
            for i in range(20)
        ]
    out_path = str(tmp_path / "edge_overflow_items.pptx")
    _render_and_check(adapter, data, out_path, "overflow_items")


# ==================== 用例5: 空字符串值 ====================

@pytest.mark.slow
def test_blank_strings(base_data, adapter, tmp_path):
    """用例5: 空字符串值测试（cover/section 字段为空字符串）"""
    data = copy.deepcopy(base_data)
    data["cover"]["title"] = ""
    data["cover"]["reporter"] = ""
    data["sections"]["work_content"][0]["title"] = ""
    data["sections"]["work_content"][0]["desc"] = ""
    out_path = str(tmp_path / "edge_blank_strings.pptx")
    _render_and_check(adapter, data, out_path, "blank_strings")


# ==================== 用例6: None 值字段 ====================

@pytest.mark.slow
def test_none_values(base_data, adapter, tmp_path):
    """用例6: None 值字段测试（invalid format，cover/section 字段为 None）"""
    data = copy.deepcopy(base_data)
    data["cover"]["title"] = None
    data["cover"]["reporter"] = None
    data["sections"]["work_content"][0]["title"] = None
    data["sections"]["work_content"][0]["desc"] = None
    out_path = str(tmp_path / "edge_none_values.pptx")
    _render_and_check(adapter, data, out_path, "none_values")


# ==================== 用例7: 极短文本 ====================

@pytest.mark.slow
def test_short_text(base_data, adapter, tmp_path):
    """用例7: 极短文本测试（单字标题/描述）"""
    data = copy.deepcopy(base_data)
    data["cover"]["title"] = "总结"
    data["cover"]["reporter"] = "李"
    data["sections"]["work_content"][0]["title"] = "迭代"
    data["sections"]["work_content"][0]["desc"] = "完成"
    out_path = str(tmp_path / "edge_short_text.pptx")
    _render_and_check(adapter, data, out_path, "short_text")
