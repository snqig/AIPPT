"""超强动画方案单元测试

覆盖：
  - 6 套主题（business/tech/formal/cinematic/dynamic-impact/minimal-plus）可加载
  - P1 透传字段：intensity / sequence / bullet_delay_ms / dir / delay_ms
  - intensity → duration 倍率映射（low=0.7x, med=1.0x, high=1.4x）
  - sequence → spec 透传（staggered 模式）
  - bullet_delay_ms → spec 透传
  - 校验层 A007 对非法新字段的降级（validators.auto_fix_outline）
  - _build_by_bullet_nodes 的 staggered 模式（后续段落 with_prev + 紧凑 delay）
"""
import pytest

from aippt.animation_themes import (
    ANIMATION_THEMES,
    get_theme,
)
from aippt.animation_scheduler import (
    build_entry_specs,
    schedule_slide_animations,
    schedule_transition,
    list_animation_themes,
    BY_BULLET_PAGE_TYPES,
)
from aippt.validators import (
    validate_animations,
    auto_fix_outline,
    ANIM_INTENSITY_ENUM,
    ANIM_SEQUENCE_ENUM,
    ANIM_DIR_ENUM,
    ANIM_MORPH_OPTION_ENUM,
)


# ==================== 主题加载 ====================

class TestSuperThemes:
    """6 套主题可用性"""

    def test_all_six_themes_present(self):
        themes = list_animation_themes()
        for name in ("business", "tech", "formal", "cinematic", "dynamic-impact", "minimal-plus"):
            assert name in themes, f"缺少主题: {name}"

    def test_new_themes_have_global_transition(self):
        for name in ("cinematic", "dynamic-impact", "minimal-plus"):
            theme = get_theme(name)
            assert "global_transition" in theme, f"{name} 缺少 global_transition"
            assert "page_overrides" in theme, f"{name} 缺少 page_overrides"

    def test_cinematic_uses_morph(self):
        """cinematic 主题全局转场应为 morph（byObject）"""
        theme = get_theme("cinematic")
        assert theme["global_transition"] == "morph"
        assert theme.get("global_transition_option") == "byObject"

    def test_new_themes_cover_all_page_types(self):
        """新增 3 套主题应覆盖全部 12 类页面类型"""
        all_page_types = {
            "cover", "catalog", "divider", "numbered_list", "kpi", "timeline",
            "two_column", "skill_percent", "preset_titles", "chart", "table", "ending",
        }
        for name in ("cinematic", "dynamic-impact", "minimal-plus"):
            theme = get_theme(name)
            covered = set(theme.get("page_overrides", {}).keys())
            missing = all_page_types - covered
            assert not missing, f"{name} 未覆盖页面类型: {missing}"


# ==================== P1 透传字段 ====================

class TestTransparentFields:
    """超强方案 P1 透传字段"""

    def test_intensity_high_applies_duration_factor(self):
        """intensity=high → duration × 1.4"""
        specs = schedule_slide_animations(
            "numbered_list",
            theme_name=None,
            page_animations={
                "entry": "fly_in",
                "by_bullet": False,
                "intensity": "high",
            },
        )
        title_spec = next(s for s in specs if s["shape"] == "title")
        # title 默认 duration 800，high → 800 * 1.4 = 1120
        assert title_spec["duration_ms"] == 1120, f"intensity=high 未生效: {title_spec}"

    def test_intensity_low_applies_duration_factor(self):
        """intensity=low → duration × 0.7"""
        specs = schedule_slide_animations(
            "numbered_list",
            theme_name=None,
            page_animations={
                "entry": "fade",
                "by_bullet": False,
                "intensity": "low",
            },
        )
        title_spec = next(s for s in specs if s["shape"] == "title")
        # title 默认 800, low → 800 * 0.7 = 560
        assert title_spec["duration_ms"] == 560, f"intensity=low 未生效: {title_spec}"

    def test_intensity_med_keeps_duration(self):
        """intensity=med → duration × 1.0（不变）"""
        specs = schedule_slide_animations(
            "numbered_list",
            theme_name=None,
            page_animations={
                "entry": "fade",
                "by_bullet": False,
                "intensity": "med",
            },
        )
        title_spec = next(s for s in specs if s["shape"] == "title")
        assert title_spec["duration_ms"] == 800

    def test_bullet_delay_ms_passthrough(self):
        """bullet_delay_ms 透传到 body spec"""
        specs = schedule_slide_animations(
            "numbered_list",
            theme_name=None,
            page_animations={
                "entry": "fly_in",
                "by_bullet": True,
                "bullet_delay_ms": 600,
            },
        )
        body_spec = next(s for s in specs if s["shape"] == "body")
        assert body_spec.get("text_build") == "by_bullet"
        assert body_spec.get("bullet_delay_ms") == 600

    def test_dir_passthrough(self):
        """dir 入场方向透传到所有 spec"""
        specs = schedule_slide_animations(
            "numbered_list",
            theme_name=None,
            page_animations={
                "entry": "fly_in",
                "by_bullet": False,
                "dir": "from_bottom",
            },
        )
        assert all(s.get("dir") == "from_bottom" for s in specs), \
            f"dir 未透传: {specs}"

    def test_sequence_staggered_passthrough(self):
        """sequence=staggered 透传到 body spec"""
        specs = schedule_slide_animations(
            "numbered_list",
            theme_name=None,
            page_animations={
                "entry": "fly_in",
                "by_bullet": True,
                "sequence": "staggered",
            },
        )
        body_spec = next(s for s in specs if s["shape"] == "body")
        assert body_spec.get("sequence") == "staggered"

    def test_sequence_sequential_not_in_spec(self):
        """sequence=sequential（默认）不写入 spec（减少噪声）"""
        specs = schedule_slide_animations(
            "numbered_list",
            theme_name=None,
            page_animations={
                "entry": "fly_in",
                "by_bullet": True,
                "sequence": "sequential",
            },
        )
        body_spec = next(s for s in specs if s["shape"] == "body")
        assert "sequence" not in body_spec, f"sequential 不应写入 spec: {body_spec}"

    def test_intensity_provides_default_bullet_delay(self):
        """intensity=high 且未设 bullet_delay_ms → 默认 400ms"""
        specs = schedule_slide_animations(
            "numbered_list",
            theme_name=None,
            page_animations={
                "entry": "fly_in",
                "by_bullet": True,
                "intensity": "high",
            },
        )
        body_spec = next(s for s in specs if s["shape"] == "body")
        assert body_spec.get("bullet_delay_ms") == 400

    def test_explicit_bullet_delay_overrides_intensity_default(self):
        """显式 bullet_delay_ms 优先于 intensity 默认值"""
        specs = schedule_slide_animations(
            "numbered_list",
            theme_name=None,
            page_animations={
                "entry": "fly_in",
                "by_bullet": True,
                "intensity": "high",
                "bullet_delay_ms": 300,
            },
        )
        body_spec = next(s for s in specs if s["shape"] == "body")
        assert body_spec.get("bullet_delay_ms") == 300


# ==================== cinematic 主题端到端 ====================

class TestCinematicThemeE2E:
    """cinematic 主题端到端：主题配置 → spec 构建"""

    def test_cinematic_numbered_list_uses_fly_in_from_bottom(self):
        """cinematic 主题 numbered_list 页应配置 fly_in + from_bottom + by_bullet"""
        theme = get_theme("cinematic")
        cfg = theme["page_overrides"]["numbered_list"]
        anim = cfg["animations"]
        assert anim["entry"] == "fly_in"
        assert anim.get("dir") == "from_bottom"
        assert anim.get("by_bullet") is True

    def test_cinematic_schedule_produces_dir_and_bullet_delay(self):
        """cinematic 主题调度后，body spec 应含 dir + bullet_delay_ms"""
        specs = schedule_slide_animations("numbered_list", theme_name="cinematic")
        body_spec = next(s for s in specs if s["shape"] == "body")
        assert body_spec.get("dir") == "from_bottom"
        assert body_spec.get("bullet_delay_ms") is not None
        assert body_spec.get("text_build") == "by_bullet"

    def test_cinematic_kpi_uses_fade_transition(self):
        """cinematic 主题 KPI 页转场应为 fade（override 全局 morph）"""
        t_spec = schedule_transition("kpi", theme_name="cinematic")
        assert t_spec is not None
        # transition spec 应反映 fade（非 morph）
        assert t_spec.get("name") == "fade" or "fade" in str(t_spec)

    def test_cinematic_cover_uses_morph_transition(self):
        """cinematic 主题 cover 页转场应为 morph"""
        t_spec = schedule_transition("cover", theme_name="cinematic")
        assert t_spec is not None
        t_str = str(t_spec)
        assert "morph" in t_str, f"cover 应使用 morph: {t_spec}"


# ==================== 校验层 A007 降级 ====================

class TestA007Validation:
    """A007 新字段校验与降级"""

    def _make_outline_with_anim(self, animations: dict) -> dict:
        return {
            "scene": "工作总结",
            "total_pages": 1,
            "pages": [{
                "page_id": 1,
                "page_type": "numbered_list",
                "title": "测试",
                "items": ["条目1", "条目2"],
                "animations": animations,
            }],
        }

    def test_a007_invalid_intensity_warns(self):
        outline = self._make_outline_with_anim({
            "entry": "fly_in", "intensity": "super_high",
        })
        r = validate_animations(outline)
        a007_warnings = [w for w in r.warnings if w.code == "A007" and "intensity" in w.path]
        assert len(a007_warnings) == 1

    def test_a007_invalid_sequence_warns(self):
        outline = self._make_outline_with_anim({
            "entry": "fly_in", "sequence": "random",
        })
        r = validate_animations(outline)
        a007_warnings = [w for w in r.warnings if w.code == "A007" and "sequence" in w.path]
        assert len(a007_warnings) == 1

    def test_a007_invalid_dir_warns(self):
        outline = self._make_outline_with_anim({
            "entry": "fly_in", "dir": "from_nowhere",
        })
        r = validate_animations(outline)
        a007_warnings = [w for w in r.warnings if w.code == "A007" and "dir" in w.path]
        assert len(a007_warnings) == 1

    def test_a007_invalid_bullet_delay_type_warns(self):
        outline = self._make_outline_with_anim({
            "entry": "fly_in", "by_bullet": True, "bullet_delay_ms": "fast",
        })
        r = validate_animations(outline)
        a007_warnings = [w for w in r.warnings if w.code == "A007" and "bullet_delay_ms" in w.path]
        assert len(a007_warnings) == 1

    def test_a007_out_of_range_delay_warns(self):
        outline = self._make_outline_with_anim({
            "entry": "fly_in", "delay_ms": 99999,
        })
        r = validate_animations(outline)
        a007_warnings = [w for w in r.warnings if w.code == "A007" and "delay_ms" in w.path]
        assert len(a007_warnings) == 1

    def test_auto_fix_removes_invalid_intensity(self):
        outline = self._make_outline_with_anim({
            "entry": "fly_in",
            "by_bullet": True,
            "intensity": "super_high",
            "sequence": "random",
            "dir": "from_nowhere",
            "bullet_delay_ms": "fast",
        })
        fixed, result = auto_fix_outline(outline)
        anim = fixed["pages"][0]["animations"]
        # 非法字段全部移除
        assert "intensity" not in anim
        assert "sequence" not in anim
        assert "dir" not in anim
        assert "bullet_delay_ms" not in anim
        # 合法字段保留
        assert anim["entry"] == "fly_in"
        assert anim["by_bullet"] is True
        # 修复记录
        assert len(result.fixed) >= 4

    def test_auto_fix_keeps_valid_new_fields(self):
        """合法的新字段应保留，不被误删"""
        outline = self._make_outline_with_anim({
            "entry": "fly_in",
            "by_bullet": True,
            "intensity": "high",
            "sequence": "staggered",
            "dir": "from_bottom",
            "bullet_delay_ms": 600,
        })
        fixed, result = auto_fix_outline(outline)
        anim = fixed["pages"][0]["animations"]
        assert anim["intensity"] == "high"
        assert anim["sequence"] == "staggered"
        assert anim["dir"] == "from_bottom"
        assert anim["bullet_delay_ms"] == 600
        # 不应有修复记录
        fix_msgs = [f for f in result.fixed if "intensity" in f or "sequence" in f or "dir" in f]
        assert fix_msgs == []

    def test_valid_new_fields_pass_validation(self):
        """合法的新字段不应产生 A007 warning"""
        outline = self._make_outline_with_anim({
            "entry": "fly_in",
            "by_bullet": True,
            "intensity": "high",
            "sequence": "staggered",
            "dir": "from_bottom",
            "bullet_delay_ms": 600,
            "delay_ms": 200,
        })
        r = validate_animations(outline)
        a007_warnings = [w for w in r.warnings if w.code == "A007"]
        assert a007_warnings == [], f"合法字段不应产生 A007: {a007_warnings}"

    def test_transition_option_invalid_downgrade(self):
        """非法 transition_option 降级"""
        outline = {
            "scene": "工作总结",
            "total_pages": 1,
            "pages": [{
                "page_id": 1,
                "page_type": "cover",
                "title": "测试",
                "transition": "morph",
                "transition_option": "bySentence",  # 非法
            }],
        }
        fixed, result = auto_fix_outline(outline)
        assert "transition_option" not in fixed["pages"][0]
        assert any("transition_option" in f for f in result.fixed)


# ==================== build_entry_specs 直接测试 ====================

class TestBuildEntrySpecs:
    """build_entry_specs 函数级测试"""

    def test_intensity_factor_applied_to_duration(self):
        specs = build_entry_specs(
            "numbered_list", "fade", by_bullet=False,
            intensity_factor=1.4,
        )
        title_spec = next(s for s in specs if s["shape"] == "title")
        # 800 * 1.4 = 1120
        assert title_spec["duration_ms"] == 1120

    def test_sequence_staggered_written_to_body_spec(self):
        specs = build_entry_specs(
            "numbered_list", "fly_in", by_bullet=True,
            sequence="staggered",
        )
        body_spec = next(s for s in specs if s["shape"] == "body")
        assert body_spec.get("sequence") == "staggered"

    def test_sequence_sequential_not_written(self):
        specs = build_entry_specs(
            "numbered_list", "fly_in", by_bullet=True,
            sequence="sequential",
        )
        body_spec = next(s for s in specs if s["shape"] == "body")
        assert "sequence" not in body_spec

    def test_group_first_role_uses_role_trigger(self):
        """group 内首角色用 ROLE_TRIGGER，不再回退为 role 名"""
        # catalog 的 group ("number", "body")，number 的 ROLE_TRIGGER="with_prev"
        specs = build_entry_specs("catalog", "fade", by_bullet=False)
        # 找到 number 角色 spec
        number_spec = next(s for s in specs if s["shape"] == "number")
        assert number_spec["trigger"] == "with_prev", \
            f"group 首角色 trigger 应为 with_prev（非 'number'）: {number_spec}"

    def test_by_bullet_only_for_body_desc_roles(self):
        """by_bullet 仅对 body/desc 角色生效，title 不应有 text_build"""
        specs = build_entry_specs(
            "numbered_list", "fly_in", by_bullet=True,
        )
        title_spec = next(s for s in specs if s["shape"] == "title")
        body_spec = next(s for s in specs if s["shape"] == "body")
        assert "text_build" not in title_spec
        assert body_spec.get("text_build") == "by_bullet"
