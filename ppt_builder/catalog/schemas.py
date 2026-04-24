"""Phase A2 카탈로그 스키마 — Pydantic 모델.

설계 원칙:
  - narrative_role을 primary 축으로 (Outline-First Planning 지원)
  - max_chars 필수 (56점 실패 #1 overflow 해결)
  - deck skeleton을 Pydantic으로 (56점 실패 #5 덱 리듬 해결)
"""
from __future__ import annotations

from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field


class NarrativeRole(str, Enum):
    """SCQA/Pyramid 기반 덱 내 슬라이드의 서사 역할.

    기반: McKinsey Pyramid Principle + Minto SCQA
    """
    OPENING = "opening"              # 표지, 제목
    AGENDA = "agenda"                # 목차
    SITUATION = "situation"          # 현재 상황, 배경
    COMPLICATION = "complication"    # 문제, 이슈, 갭
    EVIDENCE = "evidence"            # 데이터, 증거, 사실
    ANALYSIS = "analysis"            # 분석, 해석
    RECOMMENDATION = "recommendation"  # 권고, 제안
    ROADMAP = "roadmap"              # 실행 계획, 일정
    BENEFIT = "benefit"              # 기대 효과, 가치
    RISK = "risk"                    # 리스크, 제약
    CLOSING = "closing"              # 결론, 요약, Q&A
    DIVIDER = "divider"              # 섹션 구분
    APPENDIX = "appendix"            # 참고, 부록
    UNKNOWN = "unknown"              # 미분류


class TagAxes(BaseModel):
    """슬라이드 다축 태그 — retrieval 쿼리 기반."""

    narrative_role: NarrativeRole = NarrativeRole.UNKNOWN
    section: Literal[
        "opening", "agenda", "context", "analysis",
        "recommendation", "timeline", "conclusion", "appendix", "unknown",
    ] = "unknown"
    structure: Literal[
        "grid", "flow", "hierarchy", "cards", "chart",
        "text_heavy", "cover", "divider", "mixed", "unknown",
    ] = "unknown"
    density: Literal["low", "med", "high", "unknown"] = "unknown"
    intent: list[Literal[
        "define", "compare", "sequence", "decompose",
        "emphasize", "showcase", "unknown",
    ]] = Field(default_factory=lambda: ["unknown"])
    visual: list[Literal[
        "with_chart", "with_image", "with_icons",
        "with_table", "text_only",
    ]] = Field(default_factory=list)


class SlotSchema(BaseModel):
    """편집 가능한 슬롯 — max_chars로 overflow 방지."""

    slot_id: str                       # 슬라이드 내 유일 (e.g., "title", "col_1")
    role: Literal[
        "title", "subtitle", "body", "caption",
        "label", "data", "header", "footer", "kicker",
    ]
    type: Literal["text", "image", "chart"] = "text"

    # 제약 (overflow 방지)
    max_chars: int = 0                 # 95퍼센타일 측정치
    quantity: int | None = None        # 고정 수량 (예: 4열 헤더 = 4)
    quantity_hint: str | None = None   # "7 cols x up to 30 rows" 등

    # 스타일 힌트
    bold: bool = False
    italic: bool = False
    alignment: Literal["left", "center", "right", "justify"] | None = None

    # 샘플/프리뷰
    preview: list[str] = Field(default_factory=list)


class SlideMeta(BaseModel):
    """1,251장 각 슬라이드의 파싱 메타데이터 (Track 1 Stage 1 출력).

    dedup + cluster + tag 기반. 구조적 시그니처 포함.
    """

    slide_index: int                   # 0-based, 마스터 템플릿 내 인덱스
    layout_name: str = ""              # PPT 내장 layout 이름
    shape_count_total: int = 0
    shape_count_leaf: int = 0          # iter_leaf_shapes 수
    picture_count: int = 0
    chart_count: int = 0
    table_count: int = 0
    group_count: int = 0

    # 텍스트 시그니처
    text_total: str = ""               # 모든 텍스트 합
    text_hash_sim: str = ""            # SimHash (64-bit hex)
    placeholder_count: int = 0         # `~~` 개수
    korean_char_count: int = 0
    english_char_count: int = 0

    # 구조 시그니처 (클러스터링 기반)
    structure_sig: str = ""            # "paragraph×5|picture×2|table×1" 같은 형태
    has_group: bool = False
    has_smartart: bool = False

    # 덱 경계 감지 시그니처 (Track 2)
    title_text: str = ""               # 추정 제목 (가장 큰 텍스트)
    page_number_text: str = ""         # 하단 페이지 번호 텍스트
    footer_text: str = ""              # 푸터


class DeckBoundary(BaseModel):
    """Track 2 Stage B1 출력 — 1,251장 내 원본 덱 경계."""

    deck_id: str                       # "deck_001" 등
    start_index: int                   # 0-based
    end_index: int                     # 0-based, inclusive
    slide_count: int
    detection_signal: str              # 왜 이 경계인지
    confidence: float                  # 0-1


class DeckOutline(BaseModel):
    """Outline-First Planning의 중간 산출물.

    PPT 생성 시 workflow Step 1에서 이 구조를 먼저 결정.
    """

    governing_thought: str             # 덱의 한 문장 결론 (Minto)
    scqa: dict[str, str] = Field(default_factory=dict)
    #  {"situation": ..., "complication": ..., "question": ..., "answer": ...}

    skeleton_id: str                   # 적용한 스켈레톤 (skeletons.json key)
    narrative_sequence: list[NarrativeRole]  # 슬라이드별 역할 시퀀스

    target_slide_count: int


class DeckSkeleton(BaseModel):
    """Track 2 Stage B3 출력 — 추출된 표준 서사 스켈레톤."""

    skeleton_id: str                   # "pwc_proposal_30slide" 등
    use_cases: list[str]               # ["제안서", "전략 보고"]
    narrative_sequence: list[NarrativeRole]
    slide_count_range: tuple[int, int]
    frequency: int                     # 원본 덱 몇 개에서 추출됨
    example_deck_ids: list[str]
