"""SlideComposer — Layout + Zone 기반 슬라이드 조합 시스템.

기존 22개 패턴 함수와 공존하면서, 패턴/컴포넌트를 자유 조합할 수 있는
새 레이어. 기존 패턴은 그대로 사용 가능.

사용 예:
    from ppt_builder.composer import SlideComposer
    from ppt_builder.components import comp_kpi_row, comp_bar_chart_h

    composer = SlideComposer(slide)
    composer.header(SlideHeader(title="...", category="..."))

    # 레이아웃 선택 → zone별 컴포넌트 배치
    zones = composer.layout("top_bottom", split=0.3)
    comp_kpi_row(composer.canvas, kpis=[...], region=zones["top"])
    comp_bar_chart_h(composer.canvas, data=[...], region=zones["bottom"])

    composer.takeaway("핵심 인사이트")
    composer.footer(SlideFooter(...))
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from pptx.slide import Slide

from ppt_builder.primitives import Canvas, Region
from ppt_builder.patterns import SlideHeader, SlideFooter, _draw_header, _draw_footer, _draw_takeaway


# ============================================================
# Layout Definitions
# ============================================================

# 콘텐츠 영역 기본값 (헤더/인트로 이후, takeaway 이전)
DEFAULT_CONTENT = Region(0.3, 1.7, 9.4, 4.75)  # y=1.7 ~ y=6.45


def _split_h(r: Region, split: float, gap: float = 0.15) -> dict[str, Region]:
    """수평 2분할 (left/right)."""
    left_w = (r.w - gap) * split
    right_w = r.w - left_w - gap
    return {
        "left": Region(r.x, r.y, left_w, r.h),
        "right": Region(r.x + left_w + gap, r.y, right_w, r.h),
    }


def _split_v(r: Region, split: float, gap: float = 0.12) -> dict[str, Region]:
    """수직 2분할 (top/bottom)."""
    top_h = (r.h - gap) * split
    bottom_h = r.h - top_h - gap
    return {
        "top": Region(r.x, r.y, r.w, top_h),
        "bottom": Region(r.x, r.y + top_h + gap, r.w, bottom_h),
    }


def _columns(r: Region, n: int, gap: float = 0.15) -> dict[str, Region]:
    """N열 균등 분할."""
    col_w = (r.w - gap * (n - 1)) / n
    return {
        f"col_{i}": Region(r.x + i * (col_w + gap), r.y, col_w, r.h)
        for i in range(n)
    }


def _grid_2x2(r: Region, gap: float = 0.12) -> dict[str, Region]:
    """2×2 그리드."""
    cell_w = (r.w - gap) / 2
    cell_h = (r.h - gap) / 2
    return {
        "tl": Region(r.x, r.y, cell_w, cell_h),
        "tr": Region(r.x + cell_w + gap, r.y, cell_w, cell_h),
        "bl": Region(r.x, r.y + cell_h + gap, cell_w, cell_h),
        "br": Region(r.x + cell_w + gap, r.y + cell_h + gap, cell_w, cell_h),
    }


def _sidebar_left(r: Region, sidebar_w: float = 3.0, gap: float = 0.15) -> dict[str, Region]:
    """좌측 사이드바 + 우측 메인."""
    return {
        "sidebar": Region(r.x, r.y, sidebar_w, r.h),
        "main": Region(r.x + sidebar_w + gap, r.y, r.w - sidebar_w - gap, r.h),
    }


# 레이아웃 레지스트리
LAYOUTS = {
    "full": lambda r, **kw: {"main": r},
    "two_column": _split_h,
    "top_bottom": _split_v,
    "three_column": lambda r, **kw: _columns(r, 3, kw.get("gap", 0.15)),
    "four_column": lambda r, **kw: _columns(r, 4, kw.get("gap", 0.12)),
    "grid_2x2": _grid_2x2,
    "sidebar_left": _sidebar_left,
    "sidebar_right": lambda r, **kw: {
        "main": Region(r.x, r.y, r.w - kw.get("sidebar_w", 3.0) - 0.15, r.h),
        "sidebar": Region(r.x + r.w - kw.get("sidebar_w", 3.0), r.y,
                          kw.get("sidebar_w", 3.0), r.h),
    },
}


# ============================================================
# SlideComposer
# ============================================================


class SlideComposer:
    """Layout + Zone 기반 슬라이드 조합기.

    사용 흐름:
        1. composer = SlideComposer(slide)
        2. composer.header(header_spec)        # 헤더 렌더 (선택)
        3. composer.intro("설명 텍스트")         # 인트로 텍스트 (선택)
        4. zones = composer.layout("two_column", split=0.5)
        5. comp_xxx(composer.canvas, ..., region=zones["left"])
        6. comp_yyy(composer.canvas, ..., region=zones["right"])
        7. composer.takeaway("인사이트")
        8. composer.footer(footer_spec)
    """

    def __init__(self, slide: Slide):
        self.slide = slide
        self.canvas = Canvas(slide)
        self._content_y = 1.7  # 콘텐츠 시작 y (헤더/인트로 이후 조정됨)
        self._header_drawn = False
        self._takeaway_y = 6.55

    def header(self, spec: SlideHeader):
        """헤더 렌더. 콘텐츠 영역 시작 y를 자동 조정."""
        _draw_header(self.canvas, spec)
        self._header_drawn = True
        self._content_y = 1.2  # 헤더 이후 시작점

    def intro(self, text: str):
        """인트로 텍스트. 콘텐츠 영역 시작 y를 아래로 밀음."""
        self.canvas.text(
            text, x=0.3, y=self._content_y, w=9.4, h=0.30,
            size=10, color="grey_900", anchor="top",
        )
        self._content_y += 0.35

    def layout(self, name: str, **kwargs) -> dict[str, Region]:
        """레이아웃 선택 → zone dict 반환.

        Args:
            name: "full", "two_column", "top_bottom", "three_column",
                  "four_column", "grid_2x2", "sidebar_left", "sidebar_right"
            **kwargs: split (0~1), gap, sidebar_w 등 레이아웃별 파라미터

        Returns:
            {"zone_name": Region, ...}
        """
        content = Region(
            0.3, self._content_y,
            9.4, self._takeaway_y - self._content_y - 0.1,
        )

        layout_fn = LAYOUTS.get(name)
        if layout_fn is None:
            raise ValueError(f"Unknown layout: {name!r}. Available: {list(LAYOUTS)}")

        return layout_fn(content, **kwargs)

    def takeaway(self, message: str, *, y: float | None = None):
        """하단 인사이트 바."""
        _draw_takeaway(self.canvas, message, y=y or self._takeaway_y)

    def footer(self, spec: SlideFooter):
        """푸터 렌더."""
        _draw_footer(self.canvas, spec)

    # --------------------------------------------------------
    # 편의 메서드 — 패턴 함수를 zone에 배치
    # --------------------------------------------------------

    def fill_pattern(
        self,
        zone: Region,
        pattern_func,
        spec,
    ):
        """기존 패턴 함수를 zone 안에 렌더 (push_region 활용).

        주의: 기존 패턴은 절대좌표로 작성되어 있으므로,
        zone 안에 완벽히 맞지 않을 수 있음.
        region-aware로 리팩터된 패턴만 정확히 동작.
        """
        self.canvas.push_region(zone)
        try:
            pattern_func(self.slide, spec)
        finally:
            self.canvas.pop_region()


# ============================================================
# Composition Rules — 조합 지능
# ============================================================


# Zone 톤: zone별 배경색/강조 수준 프리셋
ZONE_TONES = {
    "dark": {"bg": "grey_800", "text": "white", "stripe": "grey_900"},
    "mid": {"bg": "grey_200", "text": "grey_900", "stripe": "grey_700"},
    "light": {"bg": "white", "text": "grey_900", "stripe": "grey_700"},
    "subtle": {"bg": "grey_100", "text": "grey_900", "stripe": "grey_400"},
    "accent": {"bg": "zone_alert", "text": "grey_900", "stripe": "accent"},
    "positive": {"bg": "zone_positive", "text": "grey_900", "stripe": "positive"},
    "negative": {"bg": "zone_negative", "text": "grey_900", "stripe": "negative"},
}


def apply_zone_tone(
    canvas: Canvas,
    region: Region,
    tone: str = "light",
    *,
    border: bool = True,
):
    """Zone에 배경 톤을 적용한다. 콘텐츠 렌더 전에 호출."""
    t = ZONE_TONES.get(tone, ZONE_TONES["light"])
    canvas.box(
        x=0, y=0, w=region.w, h=region.h,
        fill=t["bg"],
        border=0.5 if border else None,
        border_color="grey_mid",
        region=region,
    )


# 추천 조합 레시피 — Claude가 콘텐츠 유형을 보고 선택
COMPOSITION_RECIPES = {
    "kpi_summary_detail": {
        "description": "상단 KPI 요약 + 하단 상세 차트/해설",
        "layout": "top_bottom",
        "layout_params": {"split": 0.28},
        "zones": {
            "top": {"component": "kpi_row OR data_card_row", "tone": "light"},
            "bottom": {"component": "bar_chart OR callout_list OR bullet_list", "tone": "light"},
        },
        "when": "KPI 성과 보고 + 세부 분석이 동시에 필요할 때",
    },
    "analysis_insight": {
        "description": "좌 분석 데이터 + 우 인사이트 해설",
        "layout": "two_column",
        "layout_params": {"split": 0.5},
        "zones": {
            "left": {"component": "bar_chart OR bullet_list OR heat_map", "tone": "light"},
            "right": {"component": "callout_list OR icon_list", "tone": "light"},
        },
        "when": "데이터 분석 결과와 So What을 한 장에 담을 때",
    },
    "stats_narrative": {
        "description": "좌 핵심 지표 + 우 상세 해설",
        "layout": "sidebar_left",
        "layout_params": {"sidebar_w": 2.8},
        "zones": {
            "sidebar": {"component": "stat_column OR gauge_column", "tone": "subtle"},
            "main": {"component": "callout_list OR bullet_list", "tone": "light"},
        },
        "when": "정량 지표와 정성 해설을 좌우로 대비할 때",
    },
    "option_comparison": {
        "description": "N개 옵션 병렬 비교 (중앙 권장 강조)",
        "layout": "three_column",
        "layout_params": {},
        "zones": {
            "col_0": {"component": "option_card", "tone": "light"},
            "col_1": {"component": "option_card", "tone": "dark"},  # 권장
            "col_2": {"component": "option_card", "tone": "light"},
        },
        "when": "전략적 옵션 A/B/C 비교 시",
    },
    "status_dashboard": {
        "description": "4분면 PMO 대시보드",
        "layout": "grid_2x2",
        "layout_params": {},
        "zones": {
            "tl": {"component": "kpi_row OR gauge", "tone": "light"},
            "tr": {"component": "icon_list OR bullet_list", "tone": "subtle"},
            "bl": {"component": "bar_chart OR progress_bars", "tone": "light"},
            "br": {"component": "timeline_mini OR bullet_list", "tone": "subtle"},
        },
        "when": "프로젝트 현황을 한눈에 보여줄 때 (KPI, 리스크, 진척, 일정)",
    },
    "deep_analysis": {
        "description": "전체 폭 데이터 분석 (차트 + 해설 세로 배치)",
        "layout": "full",
        "layout_params": {},
        "zones": {
            "main": {"component": "chart + narrative", "tone": "light"},
        },
        "when": "단일 데이터 포인트를 깊이 분석할 때",
    },
}


def suggest_recipe(content_type: str) -> dict | None:
    """콘텐츠 유형에 맞는 조합 레시피를 추천한다."""
    mapping = {
        "kpi_report": "kpi_summary_detail",
        "data_analysis": "analysis_insight",
        "performance": "stats_narrative",
        "option_comparison": "option_comparison",
        "status_report": "status_dashboard",
        "deep_dive": "deep_analysis",
    }
    recipe_key = mapping.get(content_type)
    if recipe_key:
        return COMPOSITION_RECIPES[recipe_key]
    return None
