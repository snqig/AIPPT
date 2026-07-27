"""aippt/validators.py 单元测试

覆盖：
  - validate_outline：合法 outline / 缺 scene / 非法 page_type
  - validate_animations：A001-A005 错误码
  - auto_fix_outline：page_id 重排 / cover 页 by_bullet 自动关闭
"""
import pytest

from aippt.logger import logger
from aippt.validators import (
    auto_fix_outline,
    validate_animations,
    validate_outline,
)


# ==================== validate_outline ====================

def test_validate_outline_valid(sample_outline):
    """合法 outline 通过校验（无 error 级别错误）"""
    result = validate_outline(sample_outline)
    error_codes = [e.code for e in result.errors]
    assert result.is_valid, (
        f"合法 outline 应通过校验，实际错误: {error_codes}"
    )


def test_validate_outline_missing_scene():
    """缺 scene 字段报 F100"""
    outline = {
        "total_pages": 5,
        "pages": [
            {"page_id": 1, "page_type": "cover", "title": "封面"},
            {"page_id": 2, "page_type": "numbered_list", "title": "列表",
             "items": ["项1", "项2"]},
            {"page_id": 3, "page_type": "divider", "title": "分隔"},
            {"page_id": 4, "page_type": "ending", "title": "结束"},
            {"page_id": 5, "page_type": "ending", "title": "结束2"},
        ],
    }
    result = validate_outline(outline)
    error_codes = [e.code for e in result.errors]
    assert "F100" in error_codes, (
        f"缺 scene 字段应报 F100，实际错误码: {error_codes}"
    )


def test_validate_outline_invalid_page_type():
    """非法 page_type 报 F102"""
    outline = {
        "scene": "工作总结",
        "total_pages": 5,
        "pages": [
            {"page_id": 1, "page_type": "cover", "title": "封面"},
            {"page_id": 2, "page_type": "invalid_type", "title": "非法类型"},
            {"page_id": 3, "page_type": "divider", "title": "分隔"},
            {"page_id": 4, "page_type": "ending", "title": "结束"},
            {"page_id": 5, "page_type": "ending", "title": "结束2"},
        ],
    }
    result = validate_outline(outline)
    error_codes = [e.code for e in result.errors]
    assert "F102" in error_codes, (
        f"非法 page_type 应报 F102，实际错误码: {error_codes}"
    )


# ==================== validate_animations ====================

def test_validate_animations_invalid_transition():
    """非法 transition 报 A001"""
    outline = {
        "scene": "工作总结",
        "pages": [
            {"page_id": 1, "page_type": "cover", "title": "封面",
             "transition": "invalid_transition"},
        ],
    }
    result = validate_animations(outline)
    error_codes = [e.code for e in result.errors]
    assert "A001" in error_codes, (
        f"非法 transition 应报 A001，实际错误码: {error_codes}"
    )


def test_validate_animations_invalid_entry():
    """非法 entry 报 A002"""
    outline = {
        "scene": "工作总结",
        "pages": [
            {"page_id": 1, "page_type": "numbered_list", "title": "列表",
             "animations": {"entry": "invalid_entry"}},
        ],
    }
    result = validate_animations(outline)
    error_codes = [e.code for e in result.errors]
    assert "A002" in error_codes, (
        f"非法 entry 应报 A002，实际错误码: {error_codes}"
    )


def test_validate_animations_string_animations():
    """animations 为字符串报 A003"""
    outline = {
        "scene": "工作总结",
        "pages": [
            {"page_id": 1, "page_type": "numbered_list", "title": "列表",
             "animations": "fly_in"},
        ],
    }
    result = validate_animations(outline)
    error_codes = [e.code for e in result.errors]
    assert "A003" in error_codes, (
        f"animations 为字符串应报 A003，实际错误码: {error_codes}"
    )


def test_validate_animations_string_by_bullet():
    """by_bullet 为字符串报 A004"""
    outline = {
        "scene": "工作总结",
        "pages": [
            {"page_id": 1, "page_type": "numbered_list", "title": "列表",
             "animations": {"by_bullet": "true"}},
        ],
    }
    result = validate_animations(outline)
    error_codes = [e.code for e in result.errors]
    assert "A004" in error_codes, (
        f"by_bullet 为字符串应报 A004，实际错误码: {error_codes}"
    )


def test_validate_animations_by_bullet_on_cover():
    """cover 页开 by_bullet 报 A005（warning，可自动关闭）"""
    outline = {
        "scene": "工作总结",
        "pages": [
            {"page_id": 1, "page_type": "cover", "title": "封面",
             "animations": {"by_bullet": True}},
        ],
    }
    result = validate_animations(outline)
    warning_codes = [w.code for w in result.warnings]
    assert "A005" in warning_codes, (
        f"cover 页 by_bullet=True 应报 A005 警告，实际警告码: {warning_codes}"
    )
    # A005 是 warning 级别，不影响 is_valid
    assert result.is_valid, "A005 为 warning，不应影响 is_valid"


# ==================== auto_fix_outline ====================

def test_auto_fix_outline_page_id():
    """page_id 重排：不连续/非从1开始 → 1,2,3..."""
    outline = {
        "scene": "工作总结",
        "pages": [
            {"page_id": 5, "page_type": "cover", "title": "封面"},
            {"page_id": 3, "page_type": "numbered_list", "title": "列表",
             "items": ["a", "b"]},
            {"page_id": 1, "page_type": "ending", "title": "结束"},
        ],
    }
    fixed, result = auto_fix_outline(outline)
    page_ids = [p["page_id"] for p in fixed["pages"]]
    assert page_ids == [1, 2, 3], (
        f"page_id 应重排为 [1, 2, 3]，实际: {page_ids}"
    )
    # 应有修复记录
    fix_desc = " ".join(result.fixed)
    assert "page_id" in fix_desc, f"应有 page_id 修复记录，实际: {result.fixed}"
    logger.info(f"page_id 重排修复: {result.fixed}")


def test_auto_fix_outline_by_bullet_auto_close():
    """cover 页 by_bullet 自动关闭（A005 兜底修复）"""
    outline = {
        "scene": "工作总结",
        "pages": [
            {"page_id": 1, "page_type": "cover", "title": "封面",
             "animations": {"by_bullet": True}},
        ],
    }
    fixed, result = auto_fix_outline(outline)
    by_bullet = fixed["pages"][0]["animations"]["by_bullet"]
    assert by_bullet is False, (
        f"cover 页 by_bullet 应自动关闭为 False，实际: {by_bullet}"
    )
    fix_desc = " ".join(result.fixed)
    assert "by_bullet" in fix_desc and "自动关闭" in fix_desc, (
        f"应有 by_bullet 自动关闭修复记录，实际: {result.fixed}"
    )
    logger.info(f"by_bullet 自动关闭修复: {result.fixed}")
