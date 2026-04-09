"""Phase C — 패턴 라이브러리 (5개).

각 패턴은 spec(dict)을 받아 한 슬라이드를 그린다. **출발점일 뿐** —
같은 패턴 안에서도 콘텐츠와 디테일은 매번 다를 수 있다.

Claude는 슬라이드를 만들 때:
1. design_check.decide_layout_archetype()으로 어떤 패턴이 적합한지 결정
2. 해당 pattern_*() 함수를 호출
3. 또는 패턴을 출발점으로 받아 Canvas로 자유롭게 추가/변형

5개 패턴:
- executive_summary: Hero + KPI grid + roadmap chevron + 산출물 boxes
- timeline_phases: 가로 chevron + 단계별 산출물 하단 stack
- comparison_matrix: N개 항목 columns + 강조 행/셀
- process_flow: arrow_chain 가로 + 각 단계 아래 callout 박스
- quadrant_story: 2×2 grid + 하단 인사이트 박스

각 함수는 PresentationOrSlide 객체를 받지 않고, 외부에서 만든 slide를
받아 그 위에 그린다 (재사용성).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from pptx.slide import Slide

from .primitives import Canvas


# ============================================================
# Spec dataclasses (입력)
# ============================================================


@dataclass
class SlideHeader:
    title: str
    breadcrumb: str = ""
    underline: bool = True


@dataclass
class SlideFooter:
    confidential: str = "Strictly Private and Confidential"
    source: str = ""
    left: str = "pwc"
    right: str = ""


# ============================================================
# 공통 헬퍼 — 헤더/푸터 + takeaway
# ============================================================


def _draw_header(c: Canvas, header: SlideHeader):
    c.title(
        header.title,
        x=0.3, y=0.2, w=9.4, h=0.45, size=15,
        underline=header.underline, underline_color="grey_700",
    )
    if header.breadcrumb:
        c.text(
            header.breadcrumb,
            x=0.3, y=0.75, w=9.4, h=0.25,
            size=9, color="grey_700", align="left",
        )


def _draw_footer(c: Canvas, footer: SlideFooter):
    c.divider_h(x=0.3, y=7.1, w=9.4, color="border", width=0.5)
    c.text(
        footer.confidential,
        x=0.3, y=7.18, w=3.5, h=0.18,
        size=7, color="grey_700",
    )
    if footer.source:
        c.text(
            footer.source,
            x=3.0, y=7.18, w=5.5, h=0.18,
            size=7, color="grey_700",
        )
    c.text(
        footer.left,
        x=0.3, y=7.32, w=0.8, h=0.15,
        size=7, bold=True, color="grey_900",
    )
    if footer.right:
        c.text(
            footer.right,
            x=8.8, y=7.32, w=1.0, h=0.15,
            size=7, bold=True, color="grey_900", align="right",
        )


def _draw_takeaway(c: Canvas, message: str, *, y: float = 6.55):
    if not message:
        return
    c.box(x=0.3, y=y, w=9.4, h=0.42, fill="grey_800", border=None)
    c.text(
        message,
        x=0.5, y=y, w=9.1, h=0.42,
        size=10, bold=True, color="white", anchor="middle",
    )


# ============================================================
# Pattern 1 — Executive Summary
# ============================================================


@dataclass
class ExecutiveSpec:
    header: SlideHeader
    hero_label: str  # "WHY NOW" 같은 칩 텍스트
    hero_headline: str  # 큰 메시지 (1~3줄)
    hero_subtitle: str  # 부연 한 줄
    bottlenecks: list[dict]
    # bottleneck = {"num": "01", "title": "...", "kpi": "...", "bullets": [...]}
    kpis: list[dict]
    # kpi = {"value": "14%", "label": "...", "detail": "..."}
    roadmap_phases: list[dict]
    # phase = {"tag": "L1", "name": "가시화", "duration": "2~3주",
    #          "deliverables": [...]}
    takeaway: str
    footer: SlideFooter


def executive_summary(slide: Slide, spec: ExecutiveSpec):
    """v3 형식 — Hero (좌) + KPI grid (우상) + Roadmap (우하) + Takeaway."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    # ----- Hero (좌) -----
    HX, HY = 0.3, 1.15
    HW, HH = 4.45, 5.15
    c.box(x=HX, y=HY, w=HW, h=HH, fill="grey_800", border=None)
    c.label_chip(
        spec.hero_label,
        x=HX + 0.3, y=HY + 0.3, w=1.3, h=0.28,
        fill="grey_400", text_color="white",
    )
    c.text(
        spec.hero_headline,
        x=HX + 0.3, y=HY + 0.7, w=HW - 0.55, h=1.0,
        size=18, bold=True, color="white", anchor="top",
    )
    c.text(
        spec.hero_subtitle,
        x=HX + 0.3, y=HY + 1.62, w=HW - 0.55, h=0.25,
        size=9, color="grey_200", anchor="top",
    )
    c.box(
        x=HX + 0.3, y=HY + 1.95, w=HW - 0.6, h=0.012,
        fill="grey_400", border=None,
    )
    # bottlenecks
    item_h = 1.05
    item_y = HY + 2.12
    for i, b in enumerate(spec.bottlenecks):
        iy = item_y + i * item_h
        c.circle(
            x=HX + 0.3, y=iy + 0.04, d=0.38,
            fill="grey_900", border=1.0, border_color="grey_400",
            text=b.get("num", f"{i+1:02d}"),
            text_color="white", text_size=10,
        )
        c.text(
            b["title"],
            x=HX + 0.78, y=iy, w=HW - 1.1, h=0.27,
            size=11, bold=True, color="white", anchor="top",
        )
        if b.get("kpi"):
            c.text(
                b["kpi"],
                x=HX + 0.78, y=iy + 0.27, w=HW - 1.1, h=0.2,
                size=8, bold=True, color="grey_200", anchor="top",
            )
        for bi, bul in enumerate(b.get("bullets", [])):
            c.text(
                f"▪  {bul}",
                x=HX + 0.78, y=iy + 0.5 + bi * 0.25,
                w=HW - 1.0, h=0.22,
                size=8, color="grey_200", anchor="top",
            )

    # ----- KPI grid (우상) -----
    R_X, R_W, R_Y = 4.95, 4.75, 1.15
    c.section_label("정량 효과", x=R_X, y=R_Y, w=R_W, size=10)
    KPI_AREA_Y = R_Y + 0.34
    n_kpis = len(spec.kpis)
    cols = 2 if n_kpis >= 2 else 1
    rows = (n_kpis + cols - 1) // cols
    KPI_W = (R_W - 0.15 * (cols - 1)) / cols
    KPI_H = 1.25
    for i, kpi in enumerate(spec.kpis):
        col = i % cols
        row = i // cols
        kx = R_X + col * (KPI_W + 0.15)
        ky = KPI_AREA_Y + row * (KPI_H + 0.15)
        c.kpi(
            value=kpi["value"], label=kpi["label"],
            detail=kpi.get("detail", ""),
            x=kx, y=ky, w=KPI_W, h=KPI_H,
        )

    # ----- Roadmap chevron + 산출물 (우하) -----
    if spec.roadmap_phases:
        RM_Y = KPI_AREA_Y + rows * KPI_H + (rows - 1) * 0.15 + 0.22
        c.section_label("도입 로드맵", x=R_X, y=RM_Y, w=R_W, size=10)
        n_phases = len(spec.roadmap_phases)
        chev_overlap = 0.08
        chev_w = (R_W + chev_overlap * (n_phases - 1)) / n_phases
        CHEV_Y = RM_Y + 0.3
        CHEV_H = 0.42
        fills = ["grey_800", "grey_700", "grey_400", "grey_200"]
        text_colors = ["white", "white", "white", "grey_900"]
        for i, p in enumerate(spec.roadmap_phases):
            cx = R_X + i * (chev_w - chev_overlap)
            c.chevron(
                x=cx, y=CHEV_Y, w=chev_w, h=CHEV_H,
                fill=fills[i % 4],
                text=f"{p['tag']}  {p['name']}",
                text_color=text_colors[i % 4], text_size=9,
            )
        # 산출물 박스
        DELIV_Y = CHEV_Y + CHEV_H + 0.12
        DELIV_H = 0.85
        deliv_w = (R_W - 0.15 * (n_phases - 1)) / n_phases
        for i, p in enumerate(spec.roadmap_phases):
            dx = R_X + i * (deliv_w + 0.15)
            c.box(
                x=dx, y=DELIV_Y, w=deliv_w, h=DELIV_H,
                fill="grey_100", border=0.5, border_color="grey_mid",
            )
            for di, item in enumerate(p.get("deliverables", [])):
                c.text(
                    f"▪ {item}",
                    x=dx + 0.07, y=DELIV_Y + 0.08 + di * 0.23,
                    w=deliv_w - 0.14, h=0.22,
                    size=7, color="grey_900", anchor="top",
                )

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 2 — Timeline Phases
# ============================================================


@dataclass
class TimelineSpec:
    header: SlideHeader
    intro: str
    phases: list[dict]
    # phase = {"tag": "L1", "name": "...", "duration": "...",
    #          "objective": "...", "deliverables": [...], "metrics": "..."}
    takeaway: str
    footer: SlideFooter


def timeline_phases(slide: Slide, spec: TimelineSpec):
    """가로 단계 chevron + 단계별 상세 카드 (4개 단계 균등 배치)."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    # 인트로
    if spec.intro:
        c.text(
            spec.intro,
            x=0.3, y=1.05, w=9.4, h=0.32,
            size=10, color="grey_900", anchor="top",
        )

    # ----- 가로 chevron -----
    n = len(spec.phases)
    CHEV_Y = 1.5
    CHEV_H = 0.5
    chev_overlap = 0.1
    chev_w = (9.4 + chev_overlap * (n - 1)) / n
    fills = ["grey_800", "grey_700", "grey_400", "grey_200"]
    text_colors = ["white", "white", "white", "grey_900"]
    for i, p in enumerate(spec.phases):
        cx = 0.3 + i * (chev_w - chev_overlap)
        c.chevron(
            x=cx, y=CHEV_Y, w=chev_w, h=CHEV_H,
            fill=fills[i % 4],
            text=f"{p.get('tag','')}  {p['name']}",
            text_color=text_colors[i % 4], text_size=10,
        )

    # ----- 각 단계 상세 카드 -----
    CARD_Y = 2.2
    CARD_H = 4.3
    card_gap = 0.15
    card_w = (9.4 - card_gap * (n - 1)) / n
    for i, p in enumerate(spec.phases):
        cx = 0.3 + i * (card_w + card_gap)
        # 카드 박스
        c.box(
            x=cx, y=CARD_Y, w=card_w, h=CARD_H,
            fill="white", border=0.75, border_color="grey_mid",
        )
        # 상단 색 stripe
        c.box(
            x=cx, y=CARD_Y, w=card_w, h=0.06,
            fill=fills[i % 4], border=None,
        )
        # 단계 헤더
        cy = CARD_Y + 0.18
        c.text(
            p["name"],
            x=cx + 0.15, y=cy, w=card_w - 0.3, h=0.3,
            size=12, bold=True, color="grey_900", anchor="top",
        )
        c.text(
            p.get("duration", ""),
            x=cx + 0.15, y=cy + 0.3, w=card_w - 0.3, h=0.22,
            size=8, color="grey_700", anchor="top",
        )
        # 가는 구분선
        c.box(
            x=cx + 0.15, y=cy + 0.55, w=card_w - 0.3, h=0.012,
            fill="grey_mid", border=None,
        )
        # objective
        if p.get("objective"):
            c.text(
                p["objective"],
                x=cx + 0.15, y=cy + 0.65, w=card_w - 0.3, h=0.5,
                size=8, bold=True, color="grey_900", anchor="top",
            )
        # deliverables
        DELIV_Y = cy + 1.15
        c.text(
            "산출물",
            x=cx + 0.15, y=DELIV_Y, w=card_w - 0.3, h=0.2,
            size=7, bold=True, color="grey_700", anchor="top",
        )
        for di, item in enumerate(p.get("deliverables", [])):
            c.text(
                f"▪ {item}",
                x=cx + 0.18, y=DELIV_Y + 0.22 + di * 0.22,
                w=card_w - 0.33, h=0.2,
                size=8, color="grey_900", anchor="top",
            )
        # metrics (하단)
        if p.get("metrics"):
            c.box(
                x=cx + 0.15, y=CARD_Y + CARD_H - 0.45, w=card_w - 0.3, h=0.32,
                fill="grey_100", border=None,
            )
            c.text(
                p["metrics"],
                x=cx + 0.2, y=CARD_Y + CARD_H - 0.45, w=card_w - 0.4, h=0.32,
                size=7, bold=True, color="grey_900", anchor="middle",
            )

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 3 — Comparison Matrix
# ============================================================


@dataclass
class ComparisonSpec:
    header: SlideHeader
    intro: str
    options: list[dict]
    # option = {"name": "...", "summary": "...", "criteria": [...],
    #           "highlight": False}
    # criteria 항목은 각 option마다 같은 순서/길이여야 함
    criteria_labels: list[str]
    takeaway: str
    footer: SlideFooter


def comparison_matrix(slide: Slide, spec: ComparisonSpec):
    """N개 옵션을 가로 columns로 비교 + 강조 옵션 표시."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(
            spec.intro,
            x=0.3, y=1.05, w=9.4, h=0.32,
            size=10, color="grey_900", anchor="top",
        )

    n = len(spec.options)
    n_crit = len(spec.criteria_labels)

    # 좌측 criteria 라벨 영역 + 우측 옵션 grid
    LABEL_W = 1.8
    GRID_X = 0.3 + LABEL_W + 0.1
    GRID_W = 9.4 - LABEL_W - 0.1
    COL_W = (GRID_W - 0.1 * (n - 1)) / n

    GRID_Y = 1.6
    HEADER_H = 0.6
    ROW_H = (4.5 - HEADER_H) / n_crit

    # ----- 헤더: 옵션명 -----
    for i, opt in enumerate(spec.options):
        ox = GRID_X + i * (COL_W + 0.1)
        is_highlight = opt.get("highlight", False)
        fill = "grey_900" if is_highlight else "grey_700"
        c.box(
            x=ox, y=GRID_Y, w=COL_W, h=HEADER_H,
            fill=fill, border=None,
        )
        c.text(
            opt["name"],
            x=ox + 0.1, y=GRID_Y + 0.06, w=COL_W - 0.2, h=0.28,
            size=11, bold=True, color="white",
            align="center", anchor="top",
        )
        c.text(
            opt.get("summary", ""),
            x=ox + 0.1, y=GRID_Y + 0.32, w=COL_W - 0.2, h=0.25,
            size=8, color="grey_200", align="center", anchor="top",
        )

    # ----- 좌측 criteria 라벨 -----
    for ci, label in enumerate(spec.criteria_labels):
        ly = GRID_Y + HEADER_H + ci * ROW_H
        c.box(
            x=0.3, y=ly, w=LABEL_W, h=ROW_H,
            fill="grey_100", border=0.5, border_color="grey_mid",
        )
        c.text(
            label,
            x=0.4, y=ly, w=LABEL_W - 0.2, h=ROW_H,
            size=9, bold=True, color="grey_900", anchor="middle",
        )

    # ----- 그리드 셀 -----
    for i, opt in enumerate(spec.options):
        ox = GRID_X + i * (COL_W + 0.1)
        is_highlight = opt.get("highlight", False)
        for ci, val in enumerate(opt.get("criteria", [])):
            cy = GRID_Y + HEADER_H + ci * ROW_H
            cell_fill = "grey_200" if is_highlight else "white"
            c.box(
                x=ox, y=cy, w=COL_W, h=ROW_H,
                fill=cell_fill, border=0.5, border_color="grey_mid",
            )
            text_color = "grey_900"
            text_bold = is_highlight
            c.text(
                str(val),
                x=ox + 0.1, y=cy, w=COL_W - 0.2, h=ROW_H,
                size=9, bold=text_bold, color=text_color,
                align="center", anchor="middle",
            )

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 4 — Process Flow
# ============================================================


@dataclass
class ProcessSpec:
    header: SlideHeader
    intro: str
    steps: list[dict]
    # step = {"name": "...", "actor": "...", "tools": "...",
    #         "output": "...", "duration": "..."}
    takeaway: str
    footer: SlideFooter


def process_flow(slide: Slide, spec: ProcessSpec):
    """가로 arrow_chain + 각 단계 아래 callout 박스."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(
            spec.intro,
            x=0.3, y=1.05, w=9.4, h=0.32,
            size=10, color="grey_900", anchor="top",
        )

    n = len(spec.steps)
    # ----- 상단 arrow_chain (단계명) -----
    chain_y = 1.6
    chain_h = 0.55
    c.arrow_chain(
        [s["name"] for s in spec.steps],
        x=0.3, y=chain_y, w=9.4, h=chain_h,
        gap=0.25, fill="grey_700",
        text_color="white", text_size=11,
        with_arrows=True, arrow_color="grey_700",
    )

    # ----- 각 단계 아래 상세 callout -----
    detail_y = chain_y + chain_h + 0.25
    detail_h = 4.05
    gap = 0.15
    box_w = (9.4 - gap * (n - 1)) / n

    for i, step in enumerate(spec.steps):
        bx = 0.3 + i * (box_w + gap)
        c.box(
            x=bx, y=detail_y, w=box_w, h=detail_h,
            fill="white", border=0.75, border_color="grey_mid",
        )
        # 상단 stripe
        c.box(
            x=bx, y=detail_y, w=box_w, h=0.06,
            fill="grey_700", border=None,
        )
        # actor 라벨
        cy = detail_y + 0.15
        if step.get("actor"):
            c.text(
                "수행 주체",
                x=bx + 0.15, y=cy, w=box_w - 0.3, h=0.18,
                size=7, color="grey_700", anchor="top",
            )
            c.text(
                step["actor"],
                x=bx + 0.15, y=cy + 0.18, w=box_w - 0.3, h=0.25,
                size=9, bold=True, color="grey_900", anchor="top",
            )
            cy += 0.5
        # tools
        if step.get("tools"):
            c.text(
                "도구 / 기술",
                x=bx + 0.15, y=cy, w=box_w - 0.3, h=0.18,
                size=7, color="grey_700", anchor="top",
            )
            tools_lines = step["tools"].split("\n")
            for ti, tline in enumerate(tools_lines):
                c.text(
                    f"▪ {tline}",
                    x=bx + 0.15, y=cy + 0.18 + ti * 0.2,
                    w=box_w - 0.3, h=0.2,
                    size=8, color="grey_900", anchor="top",
                )
            cy += 0.18 + len(tools_lines) * 0.2 + 0.1
        # output
        if step.get("output"):
            c.text(
                "산출물",
                x=bx + 0.15, y=cy, w=box_w - 0.3, h=0.18,
                size=7, color="grey_700", anchor="top",
            )
            c.text(
                step["output"],
                x=bx + 0.15, y=cy + 0.18, w=box_w - 0.3, h=0.5,
                size=8, color="grey_900", anchor="top",
            )
        # duration (하단)
        if step.get("duration"):
            c.box(
                x=bx, y=detail_y + detail_h - 0.3, w=box_w, h=0.3,
                fill="grey_100", border=None,
            )
            c.text(
                f"⏱  {step['duration']}",
                x=bx, y=detail_y + detail_h - 0.3, w=box_w, h=0.3,
                size=8, bold=True, color="grey_900",
                align="center", anchor="middle",
            )

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 5 — Quadrant Story
# ============================================================


@dataclass
class QuadrantSpec:
    header: SlideHeader
    intro: str
    x_axis_label: str  # 가로축 의미
    y_axis_label: str  # 세로축 의미
    x_low: str
    x_high: str
    y_low: str
    y_high: str
    quadrants: list[dict]
    # 4개 항목, 순서: TL, TR, BL, BR
    # quadrant = {"title": "...", "items": [...], "highlight": False}
    insight: str  # 하단 인사이트 박스
    footer: SlideFooter


def quadrant_story(slide: Slide, spec: QuadrantSpec):
    """2×2 grid + 양축 라벨 + 하단 인사이트 박스."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(
            spec.intro,
            x=0.3, y=1.05, w=9.4, h=0.32,
            size=10, color="grey_900", anchor="top",
        )

    # ----- 2×2 grid -----
    GRID_X = 1.2
    GRID_Y = 1.6
    GRID_W = 7.4
    GRID_H = 4.4
    cell_gap = 0.12
    cell_w = (GRID_W - cell_gap) / 2
    cell_h = (GRID_H - cell_gap) / 2

    # 세로축 라벨 (회전 대신 짧은 텍스트로 좌측 끝에)
    c.text(
        spec.y_high,
        x=0.3, y=GRID_Y, w=0.85, h=0.3,
        size=8, bold=True, color="grey_700", align="right", anchor="top",
    )
    c.text(
        spec.y_low,
        x=0.3, y=GRID_Y + GRID_H - 0.3, w=0.85, h=0.3,
        size=8, bold=True, color="grey_700", align="right", anchor="bottom",
    )
    c.text(
        spec.y_axis_label,
        x=0.3, y=GRID_Y + GRID_H / 2 - 0.15, w=0.85, h=0.3,
        size=7, color="grey_700", align="right", anchor="middle",
    )

    # 가로축 라벨 (하단)
    AXIS_Y = GRID_Y + GRID_H + 0.1
    c.text(
        spec.x_low,
        x=GRID_X, y=AXIS_Y, w=cell_w, h=0.25,
        size=8, bold=True, color="grey_700", align="left", anchor="top",
    )
    c.text(
        spec.x_axis_label,
        x=GRID_X + cell_w, y=AXIS_Y, w=cell_gap + 0.4, h=0.25,
        size=7, color="grey_700", align="center", anchor="top",
    )
    c.text(
        spec.x_high,
        x=GRID_X + cell_w + cell_gap, y=AXIS_Y, w=cell_w, h=0.25,
        size=8, bold=True, color="grey_700", align="right", anchor="top",
    )

    # 4개 셀 — TL, TR, BL, BR 순서
    positions = [
        (0, 0),  # TL
        (1, 0),  # TR
        (0, 1),  # BL
        (1, 1),  # BR
    ]
    for i, (col, row) in enumerate(positions):
        if i >= len(spec.quadrants):
            break
        q = spec.quadrants[i]
        cx = GRID_X + col * (cell_w + cell_gap)
        cy = GRID_Y + row * (cell_h + cell_gap)
        is_highlight = q.get("highlight", False)
        fill = "grey_200" if is_highlight else "white"
        border_w = 1.0 if is_highlight else 0.75
        border_c = "grey_800" if is_highlight else "grey_mid"
        c.box(
            x=cx, y=cy, w=cell_w, h=cell_h,
            fill=fill, border=border_w, border_color=border_c,
        )
        # 좌측 stripe
        c.box(
            x=cx, y=cy, w=0.08, h=cell_h,
            fill="grey_800" if is_highlight else "grey_400", border=None,
        )
        # 제목
        c.text(
            q["title"],
            x=cx + 0.2, y=cy + 0.15, w=cell_w - 0.3, h=0.32,
            size=11, bold=True, color="grey_900", anchor="top",
        )
        # 가는 구분선
        c.box(
            x=cx + 0.2, y=cy + 0.5, w=cell_w - 0.4, h=0.012,
            fill="grey_mid", border=None,
        )
        # items
        for ii, item in enumerate(q.get("items", [])):
            c.text(
                f"▪  {item}",
                x=cx + 0.2, y=cy + 0.6 + ii * 0.26, w=cell_w - 0.3, h=0.25,
                size=9, color="grey_900", anchor="top",
            )

    # 하단 인사이트 박스 (insight + footer 사이)
    if spec.insight:
        IN_Y = 6.4
        IN_H = 0.55
        c.box(
            x=0.3, y=IN_Y, w=9.4, h=IN_H,
            fill="grey_100", border=0.75, border_color="grey_700",
        )
        c.box(
            x=0.3, y=IN_Y, w=0.08, h=IN_H,
            fill="grey_900", border=None,
        )
        c.text(
            "INSIGHT",
            x=0.5, y=IN_Y + 0.06, w=1.0, h=0.2,
            size=8, bold=True, color="grey_700", anchor="top",
        )
        c.text(
            spec.insight,
            x=0.5, y=IN_Y + 0.22, w=9.0, h=0.32,
            size=10, bold=True, color="grey_900", anchor="top",
        )

    _draw_footer(c, spec.footer)
