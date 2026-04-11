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


# ============================================================
# 비직사각형 레이아웃 (Blueprint Phase 1)
# ============================================================


def _center_peripheral_4(r: Region, **kw) -> dict[str, Region]:
    """중앙 + 상하좌우 4개 존.

    CP1 대응: 다이아몬드/도넛 중앙 + 주변 텍스트 블록.
    Returns: {"center", "top", "right", "bottom", "left"}
    """
    cr = kw.get("center_ratio", 0.35)
    gap = kw.get("gap", 0.15)
    cw = r.w * cr
    ch = r.h * cr
    cx = r.x + (r.w - cw) / 2
    cy = r.y + (r.h - ch) / 2
    side_w = (r.w - cw) / 2 - gap
    top_h = cy - r.y - gap
    bot_h = (r.y + r.h) - (cy + ch) - gap
    return {
        "center": Region(cx, cy, cw, ch),
        "top": Region(r.x + side_w + gap, r.y, cw, top_h),
        "bottom": Region(r.x + side_w + gap, cy + ch + gap, cw, bot_h),
        "left": Region(r.x, r.y, side_w, r.h),
        "right": Region(cx + cw + gap, r.y, side_w, r.h),
    }


def _center_peripheral_6(r: Region, **kw) -> dict[str, Region]:
    """중앙 + 6개 주변 존 (좌3 + 우3).

    CP1 대응: 헥사곤 6단계 + 주변 설명.
    Returns: {"center", "tl", "ml", "bl", "tr", "mr", "br"}
    """
    cr = kw.get("center_ratio", 0.30)
    gap = kw.get("gap", 0.10)
    cw = r.w * cr
    ch = r.h * 0.70
    cx = r.x + (r.w - cw) / 2
    cy = r.y + (r.h - ch) / 2
    side_w = (r.w - cw) / 2 - gap
    row_h = (r.h - gap * 2) / 3
    return {
        "center": Region(cx, cy, cw, ch),
        "tl": Region(r.x, r.y, side_w, row_h),
        "ml": Region(r.x, r.y + row_h + gap, side_w, row_h),
        "bl": Region(r.x, r.y + (row_h + gap) * 2, side_w, row_h),
        "tr": Region(cx + cw + gap, r.y, side_w, row_h),
        "mr": Region(cx + cw + gap, r.y + row_h + gap, side_w, row_h),
        "br": Region(cx + cw + gap, r.y + (row_h + gap) * 2, side_w, row_h),
    }


def _grid_nxm(r: Region, **kw) -> dict[str, Region]:
    """N×M 균등 그리드.

    CP2 대응: 번호+색상 그리드.
    Returns: {"r0c0", "r0c1", ..., "r{n-1}c{m-1}"}
    """
    rows = kw.get("rows", 2)
    cols = kw.get("cols", 3)
    gap = kw.get("gap", 0.10)
    cell_w = (r.w - gap * (cols - 1)) / cols
    cell_h = (r.h - gap * (rows - 1)) / rows
    return {
        f"r{ri}c{ci}": Region(
            r.x + ci * (cell_w + gap),
            r.y + ri * (cell_h + gap),
            cell_w, cell_h,
        )
        for ri in range(rows) for ci in range(cols)
    }


def _timeline_band(r: Region, **kw) -> dict[str, Region]:
    """중앙 타임라인 밴드 + 상하 교차 콘텐츠 존.

    CP3 대응: 타임라인 + 상하 교차.
    Returns: {"band", "step_0", "step_1", ...}
    """
    steps = kw.get("steps", 5)
    band_ratio = kw.get("band_ratio", 0.08)
    gap = kw.get("gap", 0.08)
    band_h = r.h * band_ratio
    band_y = r.y + (r.h - band_h) / 2
    step_w = (r.w - gap * (steps - 1)) / steps
    above_h = band_y - r.y - gap
    below_h = r.y + r.h - (band_y + band_h) - gap

    zones: dict[str, Region] = {"band": Region(r.x, band_y, r.w, band_h)}
    for i in range(steps):
        sx = r.x + i * (step_w + gap)
        if i % 2 == 0:
            zones[f"step_{i}"] = Region(sx, band_y + band_h + gap, step_w, below_h)
        else:
            zones[f"step_{i}"] = Region(sx, r.y, step_w, above_h)
    return zones


def _asymmetric_lr(r: Region, **kw) -> dict[str, Region]:
    """비대칭 좌우 분할.

    CP5 대응: 도표 + 상세 해설.
    Returns: {"diagram", "annotation"}
    """
    lr = kw.get("left_ratio", 0.45)
    gap = kw.get("gap", 0.15)
    lw = (r.w - gap) * lr
    rw = r.w - lw - gap
    return {
        "diagram": Region(r.x, r.y, lw, r.h),
        "annotation": Region(r.x + lw + gap, r.y, rw, r.h),
    }


def _t_layout(r: Region, **kw) -> dict[str, Region]:
    """T자 레이아웃 — 상단 전폭 + 하단 좌우 분할.

    Returns: {"top", "bottom_left", "bottom_right"}
    """
    tr = kw.get("top_ratio", 0.35)
    rr = kw.get("right_ratio", 0.4)
    gap = kw.get("gap", 0.12)
    top_h = (r.h - gap) * tr
    bot_h = r.h - top_h - gap
    lw = (r.w - gap) * (1 - rr)
    rw = r.w - lw - gap
    return {
        "top": Region(r.x, r.y, r.w, top_h),
        "bottom_left": Region(r.x, r.y + top_h + gap, lw, bot_h),
        "bottom_right": Region(r.x + lw + gap, r.y + top_h + gap, rw, bot_h),
    }


def _l_layout(r: Region, **kw) -> dict[str, Region]:
    """L자 레이아웃 — 좌측 전높이 + 우측 상하 분할.

    CP4 대응: 사이드바 KPI/차트 + 우측 차트/해설.
    Returns: {"left_full", "right_top", "right_bottom"}
    """
    lr = kw.get("left_ratio", 0.35)
    tr = kw.get("top_ratio", 0.5)
    gap = kw.get("gap", 0.12)
    lw = (r.w - gap) * lr
    rw = r.w - lw - gap
    top_h = (r.h - gap) * tr
    bot_h = r.h - top_h - gap
    return {
        "left_full": Region(r.x, r.y, lw, r.h),
        "right_top": Region(r.x + lw + gap, r.y, rw, top_h),
        "right_bottom": Region(r.x + lw + gap, r.y + top_h + gap, rw, bot_h),
    }


# 레이아웃 레지스트리
LAYOUTS = {
    # Era 1 — 직사각형 레이아웃 (기존)
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
    # Era 2 — Blueprint 복합 레이아웃
    "center_peripheral_4": _center_peripheral_4,
    "center_peripheral_6": _center_peripheral_6,
    "grid_nxm": _grid_nxm,
    "timeline_band": _timeline_band,
    "asymmetric_lr": _asymmetric_lr,
    "t_layout": _t_layout,
    "l_layout": _l_layout,
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
            0.15, self._content_y,
            9.7, self._takeaway_y - self._content_y - 0.1,
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
        "archetype": "kpi",
        "zones": {
            "top": {"component": "comp_kpi_row", "tone": "light"},
            "bottom": {"component": "comp_native_chart OR comp_bullet_list", "tone": "light"},
        },
        "when": "KPI 성과 보고 + 세부 분석이 동시에 필요할 때",
    },
    "analysis_insight": {
        "description": "좌 분석 데이터 + 우 인사이트 해설",
        "layout": "two_column",
        "layout_params": {"split": 0.5},
        "archetype": "split",
        "zones": {
            "left": {"component": "comp_native_chart OR comp_heatmap_grid", "tone": "light"},
            "right": {"component": "comp_bullet_list", "tone": "light"},
        },
        "when": "데이터 분석 결과와 So What을 한 장에 담을 때",
    },
    "stats_narrative": {
        "description": "좌 핵심 지표 + 우 상세 해설",
        "layout": "l_layout",
        "layout_params": {"left_ratio": 0.35},
        "archetype": "split",
        "zones": {
            "left_full": {"component": "comp_kpi_card OR comp_hero_block", "tone": "subtle"},
            "right_top": {"component": "comp_native_chart", "tone": "light"},
            "right_bottom": {"component": "comp_bullet_list", "tone": "light"},
        },
        "when": "정량 지표와 정성 해설을 좌우로 대비할 때",
    },
    "option_comparison": {
        "description": "N개 옵션 비교 (추천안 강조)",
        "layout": "full",
        "layout_params": {},
        "archetype": "table",
        "zones": {
            "main": {"component": "comp_comparison_grid (highlight=추천안)", "tone": "light"},
        },
        "when": "전략적 옵션 A/B/C 비교 시",
    },
    "status_dashboard": {
        "description": "상단 KPI + 하단 리스크/진척",
        "layout": "t_layout",
        "layout_params": {"top_ratio": 0.30},
        "archetype": "kpi",
        "zones": {
            "top": {"component": "comp_kpi_row", "tone": "light"},
            "bottom_left": {"component": "comp_bullet_list", "tone": "light"},
            "bottom_right": {"component": "comp_heatmap_grid OR comp_bullet_list", "tone": "subtle"},
        },
        "when": "프로젝트 현황을 한눈에 보여줄 때 (KPI, 리스크, 진척, 일정)",
    },
    "deep_analysis": {
        "description": "전체 폭 데이터 분석 (워터폴/차트)",
        "layout": "full",
        "layout_params": {},
        "archetype": "single_focus",
        "zones": {
            "main": {"component": "comp_waterfall OR comp_native_chart", "tone": "light"},
        },
        "when": "단일 데이터 포인트를 깊이 분석할 때",
    },
    # Compound Component 활용 레시피
    "central_hub": {
        "description": "중앙 Hub-Spoke + 4방향 텍스트",
        "layout": "center_peripheral_4",
        "layout_params": {"center_ratio": 0.42},
        "archetype": "diagram",
        "zones": {
            "center": {"component": "comp_hub_spoke_diagram", "tone": "light"},
            "left/right/top/bottom": {"component": "comp_bullet_list", "tone": "light"},
        },
        "when": "시스템 구조, 생태계, 이해관계자 맵",
    },
    "process_chevron_detail": {
        "description": "상단 쉐브론 + 하단 상세/KPI",
        "layout": "t_layout",
        "layout_params": {"top_ratio": 0.22},
        "archetype": "flow",
        "zones": {
            "top": {"component": "comp_chevron_flow", "tone": "light"},
            "bottom_left": {"component": "comp_bullet_list", "tone": "light"},
            "bottom_right": {"component": "comp_kpi_row", "tone": "light"},
        },
        "when": "프로세스 로드맵 + 단계별 상세 또는 KPI",
    },
    "numbered_process_grid": {
        "description": "N×M 번호+색상 코딩 그리드",
        "layout": "grid_nxm",
        "layout_params": {"rows": 2, "cols": 3, "gap": 0.0},
        "archetype": "grid",
        "zones": {
            "r{i}c{j}": {"component": "comp_numbered_cell", "tone": "gradient"},
        },
        "when": "다단계 프로세스, 방법론 개요, 프레임워크 소개",
    },
    "cost_bridge": {
        "description": "상단 프로세스 + 하단 워터폴 분해",
        "layout": "top_bottom",
        "layout_params": {"split": 0.20},
        "archetype": "flow",
        "zones": {
            "top": {"component": "comp_chevron_flow", "tone": "light"},
            "bottom": {"component": "comp_waterfall", "tone": "light"},
        },
        "when": "비용/매출 분해를 프로세스 맥락에서 보여줄 때",
    },
    "strategy_hero": {
        "description": "좌 Hero 메시지 + 우 KPI + 불릿",
        "layout": "l_layout",
        "layout_params": {"left_ratio": 0.45},
        "archetype": "split",
        "zones": {
            "left_full": {"component": "comp_hero_block", "tone": "dark"},
            "right_top": {"component": "comp_kpi_row", "tone": "light"},
            "right_bottom": {"component": "comp_bullet_list", "tone": "light"},
        },
        "when": "전략 방향 선언 + 정량 근거",
    },
    "risk_dashboard": {
        "description": "좌 히트맵 + 우 대응 방안",
        "layout": "two_column",
        "layout_params": {"split": 0.55},
        "archetype": "table",
        "zones": {
            "left": {"component": "comp_heatmap_grid", "tone": "light"},
            "right": {"component": "comp_bullet_list", "tone": "light"},
        },
        "when": "리스크 평가 + 대응 전략",
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
        # Compound Component 레시피
        "architecture": "central_hub",
        "process_detail": "process_chevron_detail",
        "process_grid": "numbered_process_grid",
        "cost_analysis": "cost_bridge",
        "strategy": "strategy_hero",
        "risk": "risk_dashboard",
    }
    recipe_key = mapping.get(content_type)
    if recipe_key:
        return COMPOSITION_RECIPES[recipe_key]
    return None
