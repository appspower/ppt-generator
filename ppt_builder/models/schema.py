"""Pydantic 슬라이드 스키마 모델.

아키텍처:
  - 슬라이드 = 레이아웃 + 컴포넌트들의 조합
  - 컴포넌트: card, text_block, badge, kicker, bullet, table, chart, process_flow
  - 레이아웃: columns, full, process, table, mixed
"""

from __future__ import annotations

from typing import Annotated, Any, Literal, Union
from pydantic import BaseModel, Discriminator, Field

from .enums import SlideType, LayoutType, ComponentType, ChartType


# ============================================================
# Component Models (조합 가능한 빌딩 블록)
# ============================================================

class ContentBlock(BaseModel):
    """콘텐츠 블록 (card 내부 등에서 사용)."""
    type: str = "text"  # text, heading, kpi
    text: str = ""
    bold: bool = False
    value: str = ""     # KPI 값


class CardComponent(BaseModel):
    """카드 컴포넌트 - 헤더/서브타이틀/본문, 배경/테두리/그림자."""
    type: Literal[ComponentType.CARD] = ComponentType.CARD
    header: str = ""
    subtitle: str = ""
    content: list[ContentBlock] = Field(default_factory=list)
    bullets: list[str] = Field(default_factory=list)
    style: str = "default"     # default, dark, accent
    col: int | None = None     # 컬럼 지정 (None이면 자동 배치)


class TextBlockComponent(BaseModel):
    """텍스트 블록 컴포넌트."""
    type: Literal[ComponentType.TEXT_BLOCK] = ComponentType.TEXT_BLOCK
    text: str
    bold: bool = False
    size: str = "body"         # body, small, large
    col: int | None = None


class BadgeComponent(BaseModel):
    """뱃지/태그 컴포넌트."""
    type: Literal[ComponentType.BADGE] = ComponentType.BADGE
    label: str
    style: str = "accent"      # accent, dark, light
    col: int | None = None


class KickerComponent(BaseModel):
    """키커 텍스트 (상단 작은 강조 텍스트)."""
    type: Literal[ComponentType.KICKER] = ComponentType.KICKER
    text: str
    col: int | None = None


class BulletComponent(BaseModel):
    """불릿/KPI 항목."""
    type: Literal[ComponentType.BULLET] = ComponentType.BULLET
    items: list[str]
    style: str = "default"     # default, kpi
    col: int | None = None


class TableData(BaseModel):
    """표 데이터."""
    headers: list[str]
    rows: list[list[str]]


class TableComponent(BaseModel):
    """표 컴포넌트."""
    type: Literal[ComponentType.TABLE] = ComponentType.TABLE
    data: TableData
    col: int | None = None


class ChartSeries(BaseModel):
    name: str
    values: list[float | int]


class ChartData(BaseModel):
    categories: list[str]
    series: list[ChartSeries]


class ChartComponent(BaseModel):
    """차트 컴포넌트."""
    type: Literal[ComponentType.CHART] = ComponentType.CHART
    chart_type: ChartType
    data: ChartData
    col: int | None = None


class ProcessStep(BaseModel):
    label: str
    description: str = ""


class ProcessFlowComponent(BaseModel):
    """프로세스 플로우 컴포넌트."""
    type: Literal[ComponentType.PROCESS_FLOW] = ComponentType.PROCESS_FLOW
    steps: list[ProcessStep] = Field(min_length=2)
    col: int | None = None


class ChevronStep(BaseModel):
    label: str
    description: str = ""


class ChevronProcessComponent(BaseModel):
    """쉐브론 화살표 프로세스."""
    type: Literal[ComponentType.CHEVRON_PROCESS] = ComponentType.CHEVRON_PROCESS
    steps: list[ChevronStep] = Field(min_length=2)
    col: int | None = None


class MatrixCell(BaseModel):
    text: str = ""
    highlight: bool = False
    style: str = "default"  # default, accent, dark


class FrameworkMatrixComponent(BaseModel):
    """N×M 프레임워크 매트릭스 (셀 하이라이트, 색상 코딩)."""
    type: Literal[ComponentType.FRAMEWORK_MATRIX] = ComponentType.FRAMEWORK_MATRIX
    row_headers: list[str] = Field(default_factory=list)
    col_headers: list[str] = Field(default_factory=list)
    cells: list[list[MatrixCell]] = Field(min_length=1)
    col: int | None = None


class NumberedCircleComponent(BaseModel):
    """번호 원형 뱃지 (컨설팅 필수 요소)."""
    type: Literal[ComponentType.NUMBERED_CIRCLE] = ComponentType.NUMBERED_CIRCLE
    items: list[str] = Field(min_length=1)
    style: str = "accent"  # accent, dark, grey
    col: int | None = None


class TakeawayBarComponent(BaseModel):
    """하단 핵심 메시지 바."""
    type: Literal[ComponentType.TAKEAWAY_BAR] = ComponentType.TAKEAWAY_BAR
    message: str
    style: str = "accent"  # accent, dark
    col: int | None = None


class QuadrantItem(BaseModel):
    number: str = "01"
    title: str = ""
    bullets: list[str] = Field(default_factory=list)
    style: str = "default"  # default, accent, dark, grey


class NumberedQuadrantComponent(BaseModel):
    """2×2 번호 정보 블록."""
    type: Literal[ComponentType.NUMBERED_QUADRANT] = ComponentType.NUMBERED_QUADRANT
    items: list[QuadrantItem] = Field(min_length=2, max_length=6)
    col: int | None = None


class HarveyBallMatrixComponent(BaseModel):
    """Harvey Ball 평가 매트릭스."""
    type: Literal[ComponentType.HARVEY_BALL_MATRIX] = ComponentType.HARVEY_BALL_MATRIX
    row_headers: list[str]
    col_headers: list[str]
    scores: list[list[int]] = Field(description="0/25/50/75/100 per cell")
    col: int | None = None


class VerticalFlowStep(BaseModel):
    label: str
    detail: str = ""
    style: str = "default"  # default, accent


class VerticalFlowComponent(BaseModel):
    """수직 화살표 체인."""
    type: Literal[ComponentType.VERTICAL_FLOW] = ComponentType.VERTICAL_FLOW
    steps: list[VerticalFlowStep] = Field(min_length=2)
    col: int | None = None


class FunnelStage(BaseModel):
    label: str
    value: str = ""
    conversion: str = ""


class FunnelComponent(BaseModel):
    """깔때기 (전환율 분석)."""
    type: Literal[ComponentType.FUNNEL] = ComponentType.FUNNEL
    stages: list[FunnelStage] = Field(min_length=2)
    col: int | None = None


class RAGItem(BaseModel):
    name: str
    status: str = "green"  # green, amber, red
    trend: str = ""  # up, flat, down
    note: str = ""


class RAGTableComponent(BaseModel):
    """R/A/G 상태 테이블."""
    type: Literal[ComponentType.RAG_TABLE] = ComponentType.RAG_TABLE
    items: list[RAGItem] = Field(min_length=1)
    col: int | None = None


class DividerComponent(BaseModel):
    """구분선."""
    type: Literal[ComponentType.DIVIDER] = ComponentType.DIVIDER
    col: int | None = None


class ImageComponent(BaseModel):
    """이미지."""
    type: Literal[ComponentType.IMAGE] = ComponentType.IMAGE
    path: str
    col: int | None = None


# --- Union type for all components ---

ComponentSchema = Annotated[
    Union[
        CardComponent,
        TextBlockComponent,
        BadgeComponent,
        KickerComponent,
        BulletComponent,
        TableComponent,
        ChartComponent,
        ProcessFlowComponent,
        ChevronProcessComponent,
        FrameworkMatrixComponent,
        NumberedCircleComponent,
        TakeawayBarComponent,
        NumberedQuadrantComponent,
        HarveyBallMatrixComponent,
        VerticalFlowComponent,
        FunnelComponent,
        RAGTableComponent,
        DividerComponent,
        ImageComponent,
    ],
    Field(discriminator="type"),
]


# ============================================================
# Slide Models
# ============================================================

class SlideSection(BaseModel):
    """stacked 레이아웃의 개별 섹션."""
    height_ratio: float = Field(default=0.5, ge=0.1, le=0.9, description="전체 콘텐츠 영역 대비 비율")
    layout: LayoutType = LayoutType.FULL
    n_cols: int = Field(default=1, ge=1, le=4)
    elements: list[ComponentSchema] = Field(min_length=1)


class ContentSlide(BaseModel):
    """본문 슬라이드 = 레이아웃 + 컴포넌트 조합.

    - 단일 레이아웃: layout + elements 사용
    - 복합 레이아웃: layout="stacked" + sections 사용
    - 사이드바: layout="sidebar" + elements (좌:sidebar_items, 우:main)
    - 전체화면: layout="fullscreen" + elements (헤더바 없음)
    """
    type: Literal[SlideType.CONTENT] = SlideType.CONTENT
    title: str = Field(description="Assertion Title (핵심어 중심, 문장형 금지)")
    header_message: str = ""
    breadcrumb: str = ""
    sidebar_items: list[str] = Field(default_factory=list, description="sidebar 레이아웃: 좌측 네비 항목들")
    sidebar_active: int = Field(default=0, description="sidebar: 활성화된 항목 인덱스")
    layout: LayoutType = LayoutType.COLUMNS
    n_cols: int = Field(default=1, ge=1, le=4)
    elements: list[ComponentSchema] = Field(default_factory=list)
    sections: list[SlideSection] = Field(default_factory=list, description="stacked 레이아웃용 섹션 분할")
    footnote: str = ""


class CoverSlide(BaseModel):
    """표지 슬라이드."""
    type: Literal[SlideType.COVER] = SlideType.COVER
    title: str
    subtitle: str = ""
    author: str = ""
    date: str = ""


class SectionDividerSlide(BaseModel):
    """섹션 구분 슬라이드."""
    type: Literal[SlideType.SECTION_DIVIDER] = SlideType.SECTION_DIVIDER
    section_number: int | None = None
    title: str


class ConclusionSlide(BaseModel):
    """종료/결론 슬라이드."""
    type: Literal[SlideType.CONCLUSION] = SlideType.CONCLUSION
    title: str = "Thank you"
    subtitle: str = ""


class TemplateSlide(BaseModel):
    """템플릿 인젝션 슬라이드 — 사전 제작된 복잡 슬라이드를 복제+치환."""
    type: Literal[SlideType.TEMPLATE] = SlideType.TEMPLATE
    template_name: str = Field(description="템플릿 이름 (hub_spoke, timeline, comparison, swimlane, kpi_dashboard, pyramid, sidebar_nav, before_after, swot, value_chain)")
    replacements: dict[str, str] = Field(default_factory=dict, description="{{placeholder}}: 실제값 매핑")


# --- Union type for all slides ---

SlideSchema = Annotated[
    Union[
        CoverSlide,
        ContentSlide,
        SectionDividerSlide,
        ConclusionSlide,
        TemplateSlide,
    ],
    Field(discriminator="type"),
]


# ============================================================
# Top-level Presentation Schema
# ============================================================

class PresentationMetadata(BaseModel):
    title: str = "Untitled Presentation"
    client: str = ""
    date: str = ""
    template: str = "default.pptx"
    accent_color: str = "#D04A02"   # 강조색 (커스터마이징 가능)


class PresentationSchema(BaseModel):
    metadata: PresentationMetadata = Field(default_factory=PresentationMetadata)
    slides: list[SlideSchema] = Field(min_length=1)
