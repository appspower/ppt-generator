"""슬라이드 및 컴포넌트 타입 열거형."""

from enum import Enum


# ============================================================
# Layout Types (레이아웃 엔진이 사용)
# ============================================================
class LayoutType(str, Enum):
    """슬라이드 레이아웃 유형."""
    COLUMNS = "columns"       # N개 컬럼 자동 배치
    FULL = "full"             # 전체 콘텐츠 영역 사용
    PROCESS = "process"       # 프로세스 플로우 (화살표)
    TABLE = "table"           # 표 레이아웃
    MIXED = "mixed"           # 혼합 (자유 배치)
    STACKED = "stacked"       # 수직 분할 (상단+하단 각각 다른 레이아웃)
    SIDEBAR = "sidebar"       # 좌측 사이드바 + 우측 콘텐츠
    FULLSCREEN = "fullscreen" # 헤더바 없이 전체화면


# ============================================================
# Component Types (컴포넌트 렌더러가 사용)
# ============================================================
class ComponentType(str, Enum):
    """컴포넌트 유형."""
    CARD = "card"             # 카드 (헤더+본문, 배경/테두리/그림자)
    TEXT_BLOCK = "text_block"  # 텍스트 블록
    BADGE = "badge"           # 뱃지/태그
    KICKER = "kicker"         # 키커 텍스트 (상단 작은 텍스트)
    BULLET = "bullet"         # 불릿/KPI 항목
    TABLE = "table"           # 표 컴포넌트
    CHART = "chart"           # 차트 컴포넌트
    PROCESS_FLOW = "process_flow"  # 프로세스 플로우
    CHEVRON_PROCESS = "chevron_process"  # 쉐브론 프로세스
    FRAMEWORK_MATRIX = "framework_matrix"  # N×M 프레임워크 매트릭스
    NUMBERED_CIRCLE = "numbered_circle"  # 번호 원형 뱃지
    TAKEAWAY_BAR = "takeaway_bar"  # 하단 핵심 메시지 바
    NUMBERED_QUADRANT = "numbered_quadrant"  # 2×2 번호 정보 블록
    HARVEY_BALL_MATRIX = "harvey_ball_matrix"  # Harvey Ball 평가 매트릭스
    VERTICAL_FLOW = "vertical_flow"  # 수직 화살표 체인
    FUNNEL = "funnel"         # 깔때기 (전환율)
    RAG_TABLE = "rag_table"   # R/A/G 상태 테이블
    DIVIDER = "divider"       # 구분선
    IMAGE = "image"           # 이미지


# ============================================================
# Slide Types (특수 슬라이드)
# ============================================================
class SlideType(str, Enum):
    """슬라이드 유형."""
    COVER = "cover"                   # 표지
    CONTENT = "content"               # 본문 (레이아웃 + 컴포넌트)
    SECTION_DIVIDER = "section_divider"  # 섹션 구분
    CONCLUSION = "conclusion"         # 종료/결론
    TEMPLATE = "template"             # 템플릿 인젝션 (복제 + 치환)


# ============================================================
# Chart Types
# ============================================================
class ChartType(str, Enum):
    BAR = "bar"
    HORIZONTAL_BAR = "horizontal_bar"
    STACKED_BAR = "stacked_bar"
    LINE = "line"
    PIE = "pie"
    WATERFALL = "waterfall"
