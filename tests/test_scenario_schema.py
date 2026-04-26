"""ScenarioInput 스키마 단위테스트.

검증 항목:
  - 정상 입력 통과
  - skeleton_id / narrative_sequence 둘 다 없으면 거부
  - content_by_role의 role이 narrative_sequence에 없으면 거부
  - chart_data의 role이 narrative_sequence에 없으면 거부
  - ChartSpec.series values 길이 != categories 길이 거부
  - 잘못된 role 문자열 거부
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from pydantic import ValidationError

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ppt_builder.models import (
    ChartSeriesSpec,
    ChartSpec,
    ScenarioInput,
    ScenarioMetadata,
)


def test_valid_full_input():
    s = ScenarioInput(
        scenario_name="A test",
        narrative_sequence=["opening", "situation", "recommendation", "closing"],
        content_by_role={
            "opening": ["A title"],
            "situation": ["s1", "s2"],
            "recommendation": ["r1"],
            "closing": ["wrap up"],
        },
        chart_data={
            "situation": ChartSpec(
                categories=["Q1", "Q2", "Q3"],
                series=[ChartSeriesSpec(name="Sales", values=[100.0, 120.0, 90.0])],
            )
        },
        metadata=ScenarioMetadata(title="Test", client="ACME"),
    )
    assert s.scenario_name == "A test"
    assert s.chart_data["situation"].categories == ["Q1", "Q2", "Q3"]


def test_skeleton_id_only_is_ok():
    s = ScenarioInput(
        scenario_name="skel only",
        skeleton_id="consulting_proposal_30",
    )
    assert s.skeleton_id == "consulting_proposal_30"


def test_narrative_only_is_ok():
    s = ScenarioInput(
        scenario_name="narr only",
        narrative_sequence=["opening", "closing"],
    )
    assert s.narrative_sequence == ["opening", "closing"]


def test_neither_skeleton_nor_narrative_rejected():
    with pytest.raises(ValidationError) as e:
        ScenarioInput(scenario_name="bad")
    assert "skeleton_id or narrative_sequence" in str(e.value)


def test_content_role_not_in_narrative_rejected():
    with pytest.raises(ValidationError) as e:
        ScenarioInput(
            scenario_name="x",
            narrative_sequence=["opening"],
            content_by_role={"opening": ["t"], "situation": ["s"]},
        )
    assert "content_by_role has role 'situation'" in str(e.value)


def test_chart_role_not_in_narrative_rejected():
    with pytest.raises(ValidationError) as e:
        ScenarioInput(
            scenario_name="x",
            narrative_sequence=["opening"],
            chart_data={
                "evidence": ChartSpec(
                    categories=["a"],
                    series=[ChartSeriesSpec(name="x", values=[1.0])],
                )
            },
        )
    assert "chart_data has role 'evidence'" in str(e.value)


def test_chart_series_length_mismatch_rejected():
    with pytest.raises(ValidationError) as e:
        ChartSpec(
            categories=["Q1", "Q2", "Q3"],
            series=[ChartSeriesSpec(name="bad", values=[1.0, 2.0])],
        )
    assert "has 2 values but categories has 3" in str(e.value)


def test_invalid_role_rejected():
    with pytest.raises(ValidationError):
        ScenarioInput(
            scenario_name="x",
            narrative_sequence=["bogus_role"],  # type: ignore[list-item]
        )


def test_extra_fields_forbidden():
    with pytest.raises(ValidationError):
        ScenarioInput(
            scenario_name="x",
            skeleton_id="s",
            unknown_field="oops",  # type: ignore[call-arg]
        )


def test_empty_categories_rejected():
    with pytest.raises(ValidationError):
        ChartSpec(
            categories=[],
            series=[ChartSeriesSpec(name="x", values=[1.0])],
        )


def test_chart_series_color_valid_hex_accepted():
    s = ChartSeriesSpec(name="x", values=[1.0], color="#D04A02")
    assert s.color == "#D04A02"
    # 소문자 입력은 대문자로 정규화
    s2 = ChartSeriesSpec(name="x", values=[1.0], color="#d04a02")
    assert s2.color == "#D04A02"


def test_chart_series_color_invalid_hex_rejected():
    bad_values = ["D04A02", "#D04A0", "#GG0000", "rgb(208,74,2)", "#D04A02FF"]
    for bad in bad_values:
        with pytest.raises(ValidationError):
            ChartSeriesSpec(name="x", values=[1.0], color=bad)


def test_chart_series_color_none_is_default():
    s = ChartSeriesSpec(name="x", values=[1.0])
    assert s.color is None


def test_existing_scenario_dict_loads():
    """benchmark_5_scenarios.SCENARIO_CONTENT 형식 호환 확인."""
    raw = {
        "scenario_name": "D철강 2030 넷제로 전환 로드맵",
        "skeleton_id": "transformation_roadmap_10",
        "content_by_role": {
            "opening": ["D철강 2030 넷제로 전환 로드맵"],
            "situation": ["2024년 배출량 1,200만톤"],
            "complication": ["EU CBAM 2026 시행"],
            "analysis": ["현 배출 1,200만톤 vs 2030 목표 600만톤"],
            "recommendation": ["전기로 전환 + 수소환원제철"],
            "roadmap": [
                "Phase 1 (2026-2027) 전기로 도입",
                "Phase 2 (2028-2030) 수소환원 파일럿",
            ],
            "benefit": ["탄소비용 절감 4,200억원"],
            "risk": ["수소 단가 변동"],
            "closing": ["2030 600만톤"],
        },
    }
    # narrative_sequence 없이 skeleton_id만으로도 통과해야 함
    s = ScenarioInput(**raw)
    assert s.skeleton_id == "transformation_roadmap_10"
    assert "roadmap" in s.content_by_role
