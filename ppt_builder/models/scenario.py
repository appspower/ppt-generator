"""Phase B 운영 입력 스키마 — Mode A + N1-Lite 하이브리드 빌드 함수의 1차 계약.

[docs/PROJECT_DIRECTION.md](../../docs/PROJECT_DIRECTION.md) §3 ROLE_MODE_MAP과
일치하는 13개 narrative role을 Literal 타입으로 강제.

기존 `schema.py`의 `PresentationSchema`(컴포넌트 자동 조합)와는 별개 — 이쪽은
헌법에서 거부된 N1-Full 스타일이므로 새 빌드 흐름은 본 모듈만 사용.

핵심
----
- ScenarioInput: 빌드 함수 입력
- ChartSpec: chart_data 항목 (categories + series) — replace_chart_data 자동 호출
- ScenarioMetadata: 표지/메타데이터

검증
----
- narrative_sequence와 skeleton_id 중 최소 하나 필요
- content_by_role의 모든 role이 narrative_sequence에 등장
- ChartSpec.series의 values 길이 == categories 길이
"""

from __future__ import annotations

import re
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

_HEX_COLOR_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


NarrativeRole = Literal[
    "opening",
    "agenda",
    "divider",
    "situation",
    "complication",
    "evidence",
    "analysis",
    "recommendation",
    "roadmap",
    "benefit",
    "risk",
    "closing",
    "appendix",
]

ALL_ROLES: tuple[str, ...] = (
    "opening", "agenda", "divider",
    "situation", "complication", "evidence",
    "analysis", "recommendation", "roadmap",
    "benefit", "risk", "closing", "appendix",
)


class ChartSeriesSpec(BaseModel):
    """차트 단일 시리즈.

    color
    -----
    hex `#RRGGBB` 형식으로 시리즈 색상 명시. None이면 마스터 차트 원본 색 유지.
    명시된 시리즈만 색이 변경됨 — 자동 palette는 적용 안 함 (사용자 의도 보호).
    """
    model_config = ConfigDict(extra="forbid")

    name: str = Field(min_length=1)
    values: list[float] = Field(min_length=1)
    color: str | None = Field(
        default=None,
        description="hex #RRGGBB 형식 (예: '#D04A02'). None이면 차트 원본 색 유지.",
    )

    @field_validator("color")
    @classmethod
    def _validate_hex(cls, v: str | None) -> str | None:
        if v is None:
            return v
        if not _HEX_COLOR_RE.match(v):
            raise ValueError(
                f"color must be hex '#RRGGBB' format, got: {v!r}"
            )
        return v.upper()


class ChartSpec(BaseModel):
    """차트 슬라이드의 실제 데이터.

    빌드 단계에서 매칭되는 role의 첫 차트 슬라이드에 `replace_chart_data`로
    주입. 차트 shape의 flat_idx는 빌드가 자동 탐지(`has_chart` + 첫 GraphicFrame).
    """
    model_config = ConfigDict(extra="forbid")

    categories: list[str] = Field(min_length=1)
    series: list[ChartSeriesSpec] = Field(min_length=1)

    @model_validator(mode="after")
    def _check_lengths(self) -> "ChartSpec":
        n = len(self.categories)
        for s in self.series:
            if len(s.values) != n:
                raise ValueError(
                    f"chart series '{s.name}' has {len(s.values)} values "
                    f"but categories has {n}"
                )
        return self


class ScenarioMetadata(BaseModel):
    """표지/푸터에 들어갈 부가 정보 (옵션)."""
    model_config = ConfigDict(extra="forbid")

    title: str = ""
    client: str = ""
    author: str = ""
    date: str = ""


class ScenarioInput(BaseModel):
    """빌드 함수 입력. CLI/라이브러리 진입점이 받는 형식.

    narrative_sequence와 skeleton_id 중 최소 하나가 있어야 한다. 둘 다 있으면
    narrative_sequence가 우선 (CLI override 의도).

    chart_data 키 형식
    -----------------
    NarrativeRole 그대로(예: "evidence") → 해당 role 슬라이드의 첫 차트에 주입.
    같은 role이 여러 step에 등장하면 첫 번째 step에만 적용 (단순 default).
    """
    model_config = ConfigDict(extra="forbid")

    scenario_name: str = Field(min_length=1)
    skeleton_id: str | None = None
    narrative_sequence: list[NarrativeRole] = Field(default_factory=list)
    content_by_role: dict[NarrativeRole, list[str]] = Field(default_factory=dict)
    chart_data: dict[NarrativeRole, ChartSpec] = Field(default_factory=dict)
    metadata: ScenarioMetadata = Field(default_factory=ScenarioMetadata)

    @model_validator(mode="after")
    def _check_input(self) -> "ScenarioInput":
        if not self.skeleton_id and not self.narrative_sequence:
            raise ValueError(
                "ScenarioInput requires either skeleton_id or narrative_sequence."
            )
        # narrative가 명시되었을 때만 cross-check
        if self.narrative_sequence:
            seq_set = set(self.narrative_sequence)
            for role in self.content_by_role:
                if role not in seq_set:
                    raise ValueError(
                        f"content_by_role has role '{role}' not in "
                        f"narrative_sequence={list(self.narrative_sequence)}"
                    )
            for role in self.chart_data:
                if role not in seq_set:
                    raise ValueError(
                        f"chart_data has role '{role}' not in "
                        f"narrative_sequence={list(self.narrative_sequence)}"
                    )
        return self


__all__ = [
    "NarrativeRole",
    "ALL_ROLES",
    "ChartSeriesSpec",
    "ChartSpec",
    "ScenarioMetadata",
    "ScenarioInput",
]
