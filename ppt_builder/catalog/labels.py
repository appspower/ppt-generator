"""Track 1 Stage 3 — 라벨 스키마 (3축 multi-label).

근거
----
- Phase 1B (자동 archetype 추론): 14종 60.6% 커버
- Phase 1C (24 detector): multi-label 62%
- Track 2 (스켈레톤): NarrativeRole 13종
- 사용자 35 hand-curated templates (`metadata.json`): density/intent
- HJ 179장 선별 패턴: 34종 (참고)

3 축
----
L1 Macro: chart / table / card / diagram / cover  (정확 1개)
L2 Specific archetype: 22종 enum  (multi-label 1~3개)
L3 Narrative role: NarrativeRole 13종  (multi-label 1~2개)
"""
from __future__ import annotations

from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field

from .schemas import NarrativeRole


# --- L1 Macro -------------------------------------------------------
class MacroLabel(str, Enum):
    CHART = "chart"
    TABLE = "table"
    CARD = "card"
    DIAGRAM = "diagram"
    COVER = "cover"
    UNKNOWN = "unknown"


# --- L2 Specific archetype ------------------------------------------
class ArchetypeLabel(str, Enum):
    # table/grid
    TABLE_NATIVE = "table_native"
    DENSE_GRID = "dense_grid"
    MATRIX_2X2 = "matrix_2x2"
    MATRIX_3X3 = "matrix_3x3"
    MATRIX_NXN = "matrix_NxN"

    # cards
    CARDS_2COL = "cards_2col"
    CARDS_3COL = "cards_3col"
    CARDS_4COL = "cards_4col"
    CARDS_5PLUS = "cards_5plus"
    VERTICAL_LIST = "vertical_list"

    # diagrams
    ORGCHART = "orgchart"
    HUB_SPOKE = "hub_spoke"
    FLOWCHART = "flowchart"
    ROADMAP = "roadmap"
    TIMELINE_H = "timeline_h"
    GANTT = "gantt"
    SWIMLANE = "swimlane"
    FUNNEL = "funnel"
    VENN = "venn"

    # text/visual
    CHART_NATIVE = "chart_native"
    COVER_DIVIDER = "cover_divider"
    SINGLE_BLOCK = "single_block"
    LEFT_TITLE_RIGHT_BODY = "left_title_right_body"

    # fallback
    UNKNOWN = "unknown"


# --- 검출 신뢰도 ----------------------------------------------------
class Confidence(BaseModel):
    """라벨별 신뢰도 (0~1) 와 근거."""
    label: str
    score: float = Field(ge=0.0, le=1.0)
    source: Literal["1B_grid", "1C_detector", "manual", "claude_vision"] = "1B_grid"
    reason: str | None = None


class SlideLabels(BaseModel):
    """단일 슬라이드의 multi-label 결과."""
    slide_index: int
    macro: MacroLabel = MacroLabel.UNKNOWN
    archetype: list[ArchetypeLabel] = Field(default_factory=list)
    narrative_role: list[NarrativeRole] = Field(default_factory=list)

    macro_confidence: float = 0.0
    archetype_confidences: list[Confidence] = Field(default_factory=list)
    role_confidences: list[Confidence] = Field(default_factory=list)

    overall_confidence: float = 0.0     # 후속 검수 우선순위
    needs_review: bool = False
    review_reason: str | None = None
