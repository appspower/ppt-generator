"""공통 스타일 상수 - 회사 컬러 정책 + 레퍼런스 반영.

컬러: 모노크롬(White-Grey-Black) + Orange 강조 3단계
레이아웃: Edge-to-Edge, MARGIN 0.4", 4:3 Standard
"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor


# ============================================================
# Slide Dimensions (4:3 Standard)
# ============================================================
SLIDE_WIDTH = Inches(10.0)
SLIDE_HEIGHT = Inches(7.5)

# ============================================================
# Edge-to-Edge Layout System (MARGIN 0.4")
# ============================================================
MARGIN = 0.4

# 헤더 바 (다크 배경, 완성본 기준 0.55")
HEADER_BAR_X = 0
HEADER_BAR_Y = 0
HEADER_BAR_W = 10.0
HEADER_BAR_H = 0.55

# 타이틀 (헤더 바 안)
TITLE_X = Inches(0.15)
TITLE_Y = Inches(0.08)
TITLE_W = Inches(6.5)
TITLE_H = Inches(0.4)

# 헤더 메시지 (헤더 바 바로 아래)
HEADER_X = Inches(0.3)
HEADER_Y = Inches(0.6)
HEADER_W = Inches(9.4)
HEADER_H = Inches(0.35)

# 콘텐츠 영역
CONTENT_X = Inches(0.3)
CONTENT_Y = Inches(1.0)
CONTENT_W = Inches(9.4)
CONTENT_H = Inches(6.0)

# 푸터
FOOTER_X = Inches(MARGIN)
FOOTER_Y = Inches(7.05)
FOOTER_W = Inches(10.0 - 2 * MARGIN)
FOOTER_H = Inches(0.35)

# 간격
COL_GAP = Inches(0.15)
ROW_GAP = Inches(0.12)

# ============================================================
# Color Palette - 회사 공식 컬러 (첨부 사진 기반)
# ============================================================

# --- White / Black ---
CL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)          # 255-255-255
CL_BLACK = RGBColor(0x00, 0x00, 0x00)          # 0-0-0 (본문 텍스트)
CL_BG = RGBColor(0xEB, 0xEB, 0xEB)             # 235-235-235 (배경)

# --- Orange 강조 3단계 ---
CL_ACCENT = RGBColor(0xFD, 0x51, 0x08)         # 253-81-8  (Orange - 주 강조)
CL_ACCENT_MID = RGBColor(0xFE, 0x7C, 0x39)     # 254-124-57 (Medium Orange)
CL_ACCENT_LIGHT = RGBColor(0xFF, 0xAA, 0x72)   # 255-170-114 (Light Orange)

# --- Grey 3단계 ---
CL_GREY = RGBColor(0xA1, 0xA8, 0xB3)           # 161-168-179 (Grey)
CL_GREY_MID = RGBColor(0xB5, 0xBC, 0xC4)       # 181-188-196 (Medium Grey)
CL_GREY_LIGHT = RGBColor(0xCB, 0xD1, 0xD6)     # 203-209-214 (Light Grey)

# --- 기능별 별칭 ---
CL_BODY_TEXT = CL_BLACK                         # 본문 텍스트
CL_DARK = RGBColor(0x1A, 0x1A, 0x1A)           # 다크 배경
CL_TABLE_HEADER = CL_BLACK                      # 표 헤더 (순수 검정)
CL_TABLE_ROW_ALT = RGBColor(0xF2, 0xF2, 0xF2)  # 표 교대행
CL_BORDER = RGBColor(0xDD, 0xDD, 0xDE)         # 테두리 (components.md 참조)
CL_POSITIVE = RGBColor(0x27, 0xAE, 0x60)       # 긍정/상승
CL_NEGATIVE = RGBColor(0xC0, 0x39, 0x2B)       # 부정/하락

# --- 차트 색상 순서 ---
CHART_COLORS = [
    CL_ACCENT,         # Orange (주)
    CL_ACCENT_MID,     # Medium Orange
    CL_GREY,           # Grey
    CL_GREY_MID,       # Medium Grey
    CL_GREY_LIGHT,     # Light Grey
    CL_BLACK,          # Black
]

# ============================================================
# Fonts (spec.md + components.md 기반)
# ============================================================
FONT_TITLE = "Georgia"          # 제목 (원본 유지)
FONT_BODY = "Arial"             # 본문 (라틴)
FONT_EA = "맑은 고딕"            # 동아시아 (한글)

# --- Font Sizes (4:3 슬라이드 최적화) ---
FONT_SIZE_COVER_TITLE = Pt(32)  # 표지 제목
FONT_SIZE_TITLE = Pt(16)        # 슬라이드 제목 (Assertion Title)
FONT_SIZE_HEADER = Pt(11)       # 헤더 메시지
FONT_SIZE_SUBTITLE = Pt(12)     # 부제
FONT_SIZE_BODY = Pt(10)         # 본문
FONT_SIZE_BULLET = Pt(10)       # 불릿
FONT_SIZE_SMALL = Pt(9)         # 작은 텍스트
FONT_SIZE_FOOTNOTE = Pt(7)      # 각주/출처
FONT_SIZE_TABLE_HEADER = Pt(9)  # 표 헤더
FONT_SIZE_TABLE_BODY = Pt(9)    # 표 본문
FONT_SIZE_BADGE = Pt(8)         # 뱃지
FONT_SIZE_KICKER = Pt(9)        # 키커
FONT_SIZE_SECTION = Pt(36)      # 섹션 구분 제목
FONT_SIZE_KPI = Pt(24)          # KPI 숫자


# ============================================================
# Text Height Estimation
# ============================================================

def estimate_text_height(
    text: str,
    font_pt: float,
    box_width_inches: float,
    bold: bool = False,
) -> float:
    """텍스트의 렌더링 높이를 인치로 추정한다.

    python-pptx는 실제 렌더링 높이를 알 수 없으므로,
    글자수 + 폰트크기 + 박스너비로 근사 계산한다.
    """
    if not text:
        return 0.0

    # 평균 글자 폭 (인치): 폰트 pt × 0.007 (한글은 약 2배)
    avg_char_w = font_pt * 0.007
    # 한글 비율 추정
    korean_ratio = sum(1 for c in text if ord(c) > 0x1100) / max(len(text), 1)
    avg_char_w *= (1 + korean_ratio * 0.8)
    if bold:
        avg_char_w *= 1.1

    # 줄 수 계산
    lines = text.split('\n')
    total_lines = 0
    chars_per_line = max(1, box_width_inches / avg_char_w)

    for line in lines:
        if not line.strip():
            total_lines += 0.5  # 빈 줄은 반 줄
        else:
            total_lines += max(1, len(line) / chars_per_line)

    # 줄 높이 (인치): 폰트 pt × 0.018 (줄간 포함)
    line_height = font_pt * 0.018
    return total_lines * line_height


def estimate_block_height(
    items: list[dict],
    box_width_inches: float,
) -> float:
    """콘텐츠 블록의 총 높이를 인치로 추정한다.

    items: 각 항목은 dict로, 아래 키를 가질 수 있다:
        - "text": 텍스트 문자열
        - "size": 폰트 pt (기본 9)
        - "bold": 굵게 여부 (기본 False)
        - "h": 고정 높이 오버라이드 (있으면 텍스트 추정 무시)
        - "gap": 이 항목 뒤 추가 여백 (기본 0)

    사용 예:
        h = estimate_block_height([
            {"text": "수행 주체", "size": 7, "bold": True, "gap": 0},
            {"text": "PMO + Palantir", "size": 9, "gap": 0.1},
            {"text": "도구 / 기술", "size": 7, "bold": True},
            {"text": "Jira REST\\nFoundry", "size": 8, "gap": 0.1},
        ], box_width_inches=1.5)
    """
    total = 0.0
    for item in items:
        if "h" in item:
            total += item["h"]
        elif "text" in item and item["text"]:
            total += estimate_text_height(
                item["text"],
                font_pt=item.get("size", 9),
                box_width_inches=box_width_inches,
                bold=item.get("bold", False),
            )
        total += item.get("gap", 0)
    return total


# ============================================================
# Layout Helpers
# ============================================================

def calc_columns(
    n_cols: int,
    gap: float = 0.15,
    start_x: float = 0.3,
    total_w: float = 9.4,
) -> list[tuple[float, float]]:
    """컬럼 위치와 폭을 계산한다. (인치 단위)"""
    col_w = (total_w - (n_cols - 1) * gap) / n_cols
    return [(start_x + i * (col_w + gap), col_w) for i in range(n_cols)]


def calc_rows(
    n_rows: int,
    gap: float = 0.12,
    start_y: float = 1.3,
    total_h: float = 5.7,
) -> list[tuple[float, float]]:
    """행 위치와 높이를 계산한다. (인치 단위)"""
    row_h = (total_h - (n_rows - 1) * gap) / n_rows
    return [(start_y + i * (row_h + gap), row_h) for i in range(n_rows)]
