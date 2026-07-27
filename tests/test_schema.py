"""schemas/outline.schema.json 单元测试

覆盖：
  - 合法 outline 通过 schema
  - chart_data 字段校验
  - table rows 字段校验
  - transition 枚举校验
"""
import pytest

from aippt.logger import logger
from aippt.validators import validate_schema


def _make_pages(extra_pages=None):
    """构造 4 页基础 outline（满足 minItems=4）"""
    pages = [
        {"page_id": 1, "page_type": "cover", "title": "封面"},
        {"page_id": 2, "page_type": "divider", "title": "分隔"},
        {"page_id": 3, "page_type": "divider", "title": "分隔2"},
        {"page_id": 4, "page_type": "ending", "title": "结束"},
    ]
    if extra_pages:
        pages[1:1] = extra_pages  # 插入到封面之后、结尾之前
    return pages


def test_outline_schema_valid(sample_outline):
    """合法 outline 通过 schema"""
    result = validate_schema(sample_outline, "outline")
    assert result.is_valid, (
        f"合法 outline 应通过 schema 校验，错误: {[e.to_dict() for e in result.errors]}"
    )


def test_outline_schema_chart_data():
    """chart_data 字段校验：合法/非法（缺 series）"""
    # 合法 chart_data
    valid = {
        "scene": "工作总结",
        "pages": _make_pages([
            {"page_id": 5, "page_type": "chart", "title": "图表",
             "chart_type": "bar",
             "chart_data": {
                 "categories": ["Q1", "Q2"],
                 "series": [{"name": "收入", "data": [100, 200]}],
             }},
        ]),
    }
    result = validate_schema(valid, "outline")
    assert result.is_valid, (
        f"合法 chart_data 应通过，错误: {[e.to_dict() for e in result.errors]}"
    )

    # 非法 chart_data（缺少 series 必填字段）
    invalid = {
        "scene": "工作总结",
        "pages": _make_pages([
            {"page_id": 5, "page_type": "chart", "title": "图表",
             "chart_type": "bar",
             "chart_data": {"categories": ["Q1", "Q2"]}},
        ]),
    }
    result = validate_schema(invalid, "outline")
    assert not result.is_valid, "缺少 series 的 chart_data 应校验失败"
    logger.info(f"chart_data 缺 series 错误: {[e.code for e in result.errors]}")


def test_outline_schema_table_rows():
    """table rows 字段校验：合法/非法（空数组违反 minItems=1）"""
    # 合法 table
    valid = {
        "scene": "工作总结",
        "pages": _make_pages([
            {"page_id": 5, "page_type": "table", "title": "表格",
             "headers": ["指标", "数值"],
             "rows": [["用户", "100w"], ["收入", "200w"]]},
        ]),
    }
    result = validate_schema(valid, "outline")
    assert result.is_valid, (
        f"合法 table 应通过，错误: {[e.to_dict() for e in result.errors]}"
    )

    # 非法 table（rows 为空数组，违反 minItems: 1）
    invalid = {
        "scene": "工作总结",
        "pages": _make_pages([
            {"page_id": 5, "page_type": "table", "title": "表格",
             "headers": ["指标", "数值"],
             "rows": []},
        ]),
    }
    result = validate_schema(invalid, "outline")
    assert not result.is_valid, "空 rows 数组应校验失败（minItems=1）"
    logger.info(f"空 rows 错误: {[e.code for e in result.errors]}")


def test_outline_schema_transition_enum():
    """transition 枚举校验：合法值通过，非法值失败"""
    # 合法 transition
    valid = {
        "scene": "工作总结",
        "pages": [
            {"page_id": 1, "page_type": "cover", "title": "封面",
             "transition": "fade"},
            {"page_id": 2, "page_type": "divider", "title": "分隔"},
            {"page_id": 3, "page_type": "divider", "title": "分隔2"},
            {"page_id": 4, "page_type": "ending", "title": "结束"},
        ],
    }
    result = validate_schema(valid, "outline")
    assert result.is_valid, (
        f"合法 transition 应通过，错误: {[e.to_dict() for e in result.errors]}"
    )

    # 非法 transition（不在枚举内）
    invalid = {
        "scene": "工作总结",
        "pages": [
            {"page_id": 1, "page_type": "cover", "title": "封面",
             "transition": "invalid_transition"},
            {"page_id": 2, "page_type": "divider", "title": "分隔"},
            {"page_id": 3, "page_type": "divider", "title": "分隔2"},
            {"page_id": 4, "page_type": "ending", "title": "结束"},
        ],
    }
    result = validate_schema(invalid, "outline")
    assert not result.is_valid, "非法 transition 应校验失败"
    logger.info(f"非法 transition 错误: {[e.code for e in result.errors]}")
