"""build_scenario 에러 경로 단위테스트.

데이터 의존성을 피하기 위해 _load 단계 실패 케이스 위주.
정상 빌드 경로는 별도 통합 테스트에서 (master/catalog/store 필요).
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ppt_builder.api import BuildError, build_scenario
from ppt_builder.models import ScenarioInput


def test_unknown_skeleton_id_raises_build_error():
    s = ScenarioInput(
        scenario_name="x",
        skeleton_id="__definitely_not_a_real_skeleton__",
    )
    with pytest.raises(BuildError) as e:
        build_scenario(s)
    assert "skeleton_id" in str(e.value)
    assert "not found" in str(e.value) or "available" in str(e.value)


def test_safe_filename_handles_unicode_and_punct():
    from ppt_builder.api import _safe_filename

    assert _safe_filename("우리 회사 전략 — 2026") == "우리_회사_전략___2026"
    assert _safe_filename("***bad***") == "bad"
    assert _safe_filename("") == "scenario"
    assert len(_safe_filename("x" * 200)) == 60


def test_build_result_summary_line_no_charts():
    from ppt_builder.api import BuildResult

    r = BuildResult(
        pptx=Path("out/foo.pptx"),
        plan=[{}, {}, {}],
        edits=[],
        n_mode_a=2,
        n_n1_lite=1,
    )
    line = r.summary_line()
    assert "foo.pptx" in line
    assert "3 slides" in line
    assert "2 mode_a" in line
    assert "1 n1_lite" in line
    assert "charts" not in line


def test_build_result_summary_line_with_charts():
    from ppt_builder.api import BuildResult

    r = BuildResult(
        pptx=Path("out/foo.pptx"),
        plan=[{}],
        edits=[],
        n_mode_a=1,
        n_n1_lite=0,
        chart_injected={"evidence": True, "analysis": False},
    )
    line = r.summary_line()
    assert "charts: 1/2" in line
