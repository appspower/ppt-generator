"""Phase C — 패턴 라이브러리 (22개).

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

from .assembler.styles import estimate_text_height, estimate_block_height
from .primitives import Canvas


# ============================================================
# Spec dataclasses (입력)
# ============================================================


@dataclass
class SlideHeader:
    """컨설팅 슬라이드 헤더 — 레퍼런스 분석 기반 재구조화.

    레이아웃:
        ┌──────────────────────────────────────────────────┐
        │ category (라인 위 좌측, 11pt)   nav_path (우측) │
        │ ──────────────────────────────────────  (라인)   │
        │ head_message (라인 아래, 14pt 굵게, 한국어 어미) │
        └──────────────────────────────────────────────────┘

    Fields:
        title: head message — 라인 아래 큰 글씨, 슬라이드 결론 한 문장
            (예: "...구축", "...지원", "...임" 등 한국어 단호한 어미)
        category: 라인 위 좌측 — 슬라이드 소분류
            (예: "1. 일반 현황 - PwC Global 소개")
        nav_path: 라인 위 우측 — 다단계 nav, "/"로 join
            (예: ["1. 제안사 소개", "2. 협력사 소개"])
        breadcrumb: deprecated — 하위 호환을 위해 유지. 비어있으면 무시.
            기존 spec에서 breadcrumb만 채운 경우 nav_path 첫 항목으로 fallback
        underline: 라인 표시 여부
    """
    title: str
    category: str = ""
    nav_path: list[str] = field(default_factory=list)
    breadcrumb: str = ""  # deprecated, 하위 호환
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
    """레퍼런스 컨설팅 슬라이드 헤더 구조 (3단):
       1) 라인 위 좌측 category, 우측 nav_path
       2) 가는 회색 라인
       3) 라인 아래 head message (title 필드)

    하위 호환: category가 비어 있으면 breadcrumb를 fallback으로 사용
    """
    # 1) 라인 위 좌측 — category
    cat = header.category or header.breadcrumb
    if cat:
        c.text(
            cat,
            x=0.3, y=0.18, w=5.5, h=0.26,
            size=10, bold=True, color="grey_900", anchor="middle",
        )

    # 1) 라인 위 우측 — nav_path
    if header.nav_path:
        nav_str = "  /  ".join(header.nav_path)
        c.text(
            nav_str,
            x=5.8, y=0.18, w=3.9, h=0.26,
            size=9, color="grey_700", align="right", anchor="middle",
        )

    # 2) 가는 회색 라인 (cat 또는 nav 있을 때만)
    line_y = 0.5
    if header.underline and (cat or header.nav_path):
        c.box(
            x=0.3, y=line_y, w=9.4, h=0.012,
            fill="grey_700", border=None,
        )

    # 3) 라인 아래 — head message (title)
    head_y = 0.6 if (cat or header.nav_path) else 0.2
    c.text(
        header.title,
        x=0.3, y=head_y, w=9.4, h=0.5,
        size=14, bold=True, color="grey_900",
        font="Arial",  # 한글 가독성: Arial로 설정 (한글은 시스템 fallback)
        anchor="middle",
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


def _draw_aux_items(
    c: Canvas,
    aux_items: list[tuple[str, str]],
    *,
    x: float,
    y_start: float,
    w: float,
    y_limit: float,
    label_size: float = 7,
    value_size: float = 8,
):
    """aux 항목들을 동적 높이로 배치. 텍스트 길이에 따라 각 항목 높이가 다름.

    겹침 방지: estimate_text_height()로 value 텍스트 높이를 추정하고 누적.
    y_limit을 넘으면 남은 항목은 건너뜀.
    """
    text_w = w
    cy = y_start
    label_h = 0.16  # 라벨 고정 높이
    divider_h = 0.01
    pad_top = 0.03  # 구분선 후 패딩
    pad_bottom = 0.06  # 항목 간 하단 패딩

    for aux_label, aux_value in aux_items:
        # value 텍스트 높이 추정
        val_h = max(0.18, estimate_text_height(
            aux_value, font_pt=value_size, box_width_inches=text_w,
        ))
        item_total = divider_h + pad_top + label_h + val_h + pad_bottom

        # 남은 공간 부족하면 중단
        if cy + item_total > y_limit + 0.02:
            break

        # 구분선
        c.box(x=x, y=cy, w=w, h=divider_h, fill="grey_mid", border=None)
        # 라벨
        c.text(
            aux_label,
            x=x, y=cy + pad_top, w=w, h=label_h,
            size=label_size, bold=True, color="grey_700", anchor="top",
        )
        # value
        c.text(
            aux_value,
            x=x, y=cy + pad_top + label_h, w=w, h=val_h,
            size=value_size, color="grey_900", anchor="top",
        )
        cy += item_total


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

    # ----- Hero (좌) — 동적 높이 -----
    HX, HY = 0.3, 1.25
    HW = 4.45
    hero_content_w = HW - 0.55

    # bottlenecks 높이 추정
    bn_items = []
    for b in spec.bottlenecks:
        bn_items.append({"text": b["title"], "size": 11, "bold": True})
        if b.get("kpi"):
            bn_items.append({"text": b["kpi"], "size": 8, "bold": True})
        for bul in b.get("bullets", []):
            bn_items.append({"text": f"▪  {bul}", "size": 8})
        bn_items.append({"h": 0.1})  # 항목 간 gap
    bn_h = estimate_block_height(bn_items, hero_content_w)
    # overhead: chip(0.28+0.3pad) + headline(1.0) + subtitle(0.25) + line(0.012+0.1pad) + bottom pad
    hero_content_h = 0.58 + 1.0 + 0.25 + 0.112 + bn_h + 0.15
    default_avail = 6.55 - HY - 0.08  # takeaway 전까지
    abs_max = 7.05 - HY - 0.08
    # 가용 공간을 기본으로 채우되, 콘텐츠가 넘치면 확장
    HH = max(default_avail, min(hero_content_h, abs_max))
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
    # bottlenecks (Hero 박스 5.05" 안에 3개 균등 배치)
    item_h = 0.96
    item_y = HY + 2.06
    for i, b in enumerate(spec.bottlenecks):
        iy = item_y + i * item_h
        c.circle(
            x=HX + 0.3, y=iy + 0.02, d=0.36,
            fill="grey_900", border=1.0, border_color="grey_400",
            text=b.get("num", f"{i+1:02d}"),
            text_color="white", text_size=10,
        )
        c.text(
            b["title"],
            x=HX + 0.76, y=iy, w=HW - 1.1, h=0.25,
            size=11, bold=True, color="white", anchor="top",
        )
        if b.get("kpi"):
            c.text(
                b["kpi"],
                x=HX + 0.76, y=iy + 0.25, w=HW - 1.1, h=0.2,
                size=8, bold=True, color="grey_200", anchor="top",
            )
        for bi, bul in enumerate(b.get("bullets", [])):
            c.text(
                f"▪  {bul}",
                x=HX + 0.76, y=iy + 0.46 + bi * 0.23,
                w=HW - 1.0, h=0.22,
                size=8, color="grey_200", anchor="top",
            )

    # ----- KPI grid (우상) -----
    R_X, R_W, R_Y = 4.95, 4.75, 1.25
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
    #          "objective": "...", "deliverables": [...], "metrics": "...",
    #          -- aux (선택, 빈 공간 채움용) --
    #          "prerequisites": "...", "gate": "...", "team": "...", "risks": "..."}
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
            x=0.3, y=1.20, w=9.4, h=0.30,
            size=10, color="grey_900", anchor="top",
        )

    # ----- 가로 chevron -----
    n = len(spec.phases)
    CHEV_Y = 1.55
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
    CARD_Y = 2.25
    card_gap = 0.15
    card_w = (9.4 - card_gap * (n - 1)) / n
    phase_content_w = card_w - 0.3

    # 동적 높이: 각 phase 콘텐츠 높이 추정 → 최대값 기준 통일
    def _estimate_phase_h(p):
        items = [
            {"text": p["name"], "size": 12, "bold": True},
            {"text": p.get("duration", ""), "size": 8, "gap": 0.05},
            {"h": 0.012},  # 구분선
        ]
        if p.get("objective"):
            items.append({"text": p["objective"], "size": 8, "bold": True, "gap": 0.1})
        items.append({"text": "산출물", "size": 7, "bold": True})
        for d in p.get("deliverables", []):
            items.append({"text": f"▪ {d}", "size": 8})
        # overhead: stripe(0.06) + top pad(0.18) + metrics bar(0.50)
        return estimate_block_height(items, phase_content_w) + 0.74

    max_phase_h = max(_estimate_phase_h(p) for p in spec.phases)
    default_avail = 6.55 - CARD_Y - 0.08
    abs_max = 7.05 - CARD_Y - 0.08
    # 가용 공간을 기본으로 채우되, 콘텐츠가 넘치면 확장
    CARD_H = max(default_avail, min(max_phase_h, abs_max))
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
        n_deliverables = len(p.get("deliverables", []))
        for di, item in enumerate(p.get("deliverables", [])):
            c.text(
                f"▪ {item}",
                x=cx + 0.18, y=DELIV_Y + 0.22 + di * 0.22,
                w=card_w - 0.33, h=0.2,
                size=8, color="grey_900", anchor="top",
            )

        # --- aux 콘텐츠 (빈 공간 채움 — 동적 높이) ---
        aux_y = DELIV_Y + 0.22 + n_deliverables * 0.22 + 0.12
        metrics_top = CARD_Y + CARD_H - 0.50  # metrics 바 시작 y

        aux_items = []
        if p.get("prerequisites"):
            aux_items.append(("선행 조건", p["prerequisites"]))
        if p.get("gate"):
            aux_items.append(("Gate 기준", p["gate"]))
        if p.get("team"):
            aux_items.append(("투입 인력", p["team"]))
        if p.get("risks"):
            aux_items.append(("리스크", p["risks"]))

        _draw_aux_items(
            c, aux_items,
            x=cx + 0.15, y_start=aux_y, w=card_w - 0.3,
            y_limit=metrics_top,
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
            x=0.3, y=1.20, w=9.4, h=0.30,
            size=10, color="grey_900", anchor="top",
        )

    n = len(spec.options)
    n_crit = len(spec.criteria_labels)

    # 좌측 criteria 라벨 영역 + 우측 옵션 grid
    LABEL_W = 1.8
    GRID_X = 0.3 + LABEL_W + 0.1
    GRID_W = 9.4 - LABEL_W - 0.1
    COL_W = (GRID_W - 0.1 * (n - 1)) / n

    GRID_Y = 1.7
    HEADER_H = 0.6
    # 동적 행 높이: 콘텐츠 비례 배분으로 가용 공간 채움
    cell_text_w = COL_W - 0.2
    total_avail = 6.55 - GRID_Y - HEADER_H - 0.15  # takeaway 전까지

    # 각 행의 콘텐츠 기반 최소 높이 추정
    raw_heights = []
    for ci in range(n_crit):
        texts = [spec.criteria_labels[ci]]
        for opt in spec.options:
            crits = opt.get("criteria", [])
            if ci < len(crits):
                texts.append(str(crits[ci]))
        max_cell_h = max(
            estimate_text_height(t, font_pt=9, box_width_inches=cell_text_w)
            for t in texts
        )
        raw_heights.append(max(0.35, max_cell_h + 0.12))

    # 가용 공간을 콘텐츠 비율로 비례 배분 (축소 OR 확대)
    total_raw = sum(raw_heights)
    scale = total_avail / total_raw if total_raw > 0 else 1.0
    row_heights = [h * scale for h in raw_heights]

    # 행 y좌표 누적 계산
    row_y_offsets = []
    cum_y = 0.0
    for rh in row_heights:
        row_y_offsets.append(cum_y)
        cum_y += rh

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
        ly = GRID_Y + HEADER_H + row_y_offsets[ci]
        rh = row_heights[ci]
        c.box(
            x=0.3, y=ly, w=LABEL_W, h=rh,
            fill="grey_100", border=0.5, border_color="grey_mid",
        )
        c.text(
            label,
            x=0.4, y=ly, w=LABEL_W - 0.2, h=rh,
            size=9, bold=True, color="grey_900", anchor="middle",
        )

    # ----- 그리드 셀 -----
    for i, opt in enumerate(spec.options):
        ox = GRID_X + i * (COL_W + 0.1)
        is_highlight = opt.get("highlight", False)
        for ci, val in enumerate(opt.get("criteria", [])):
            cy = GRID_Y + HEADER_H + row_y_offsets[ci]
            rh = row_heights[ci]
            cell_fill = "grey_200" if is_highlight else "white"
            c.box(
                x=ox, y=cy, w=COL_W, h=rh,
                fill=cell_fill, border=0.5, border_color="grey_mid",
            )
            text_color = "grey_900"
            text_bold = is_highlight
            c.text(
                str(val),
                x=ox + 0.1, y=cy, w=COL_W - 0.2, h=rh,
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
    #         "output": "...", "duration": "...",
    #         -- aux (선택, 빈 공간 채움용) --
    #         "prerequisites": "...", "risks": "...", "metrics": "...", "example": "..."}
    takeaway: str
    footer: SlideFooter


def process_flow(slide: Slide, spec: ProcessSpec):
    """가로 arrow_chain + 각 단계 아래 callout 박스."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(
            spec.intro,
            x=0.3, y=1.20, w=9.4, h=0.30,
            size=10, color="grey_900", anchor="top",
        )

    n = len(spec.steps)
    # ----- 상단 arrow_chain (단계명) -----
    chain_y = 1.65
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
    gap = 0.15
    box_w = (9.4 - gap * (n - 1)) / n
    content_w = box_w - 0.3  # 내부 텍스트 폭

    # 동적 높이: 가용 공간을 채우되, 콘텐츠 오버플로 시 확장
    def _estimate_step_h(step):
        items = []
        if step.get("actor"):
            items += [
                {"text": "수행 주체", "size": 7, "bold": True},
                {"text": step["actor"], "size": 9, "gap": 0.07},
            ]
        if step.get("tools"):
            items.append({"text": "도구 / 기술", "size": 7, "bold": True})
            for tl in step["tools"].split("\n"):
                items.append({"text": f"▪ {tl}", "size": 8})
            items[-1]["gap"] = 0.1
        if step.get("output"):
            items += [
                {"text": "산출물", "size": 7, "bold": True},
                {"text": step["output"], "size": 8, "gap": 0.07},
            ]
        # overhead: top padding(0.15) + stripe(0.06) + duration bar(0.35)
        return estimate_block_height(items, content_w) + 0.56

    max_content_h = max(_estimate_step_h(s) for s in spec.steps)
    # 가용 공간 (takeaway y=6.55 전까지)
    default_avail = 6.55 - detail_y - 0.08
    # 절대 상한 (footer 전까지)
    abs_max = 7.05 - detail_y - 0.08
    # 가용 공간을 기본으로 채우되, 콘텐츠가 넘치면 확장
    detail_h = max(default_avail, min(max_content_h, abs_max))

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
                x=bx + 0.15, y=cy + 0.18, w=box_w - 0.3, h=0.35,
                size=8, color="grey_900", anchor="top",
            )
            cy += 0.55

        # --- aux 콘텐츠 (빈 공간 채움 — 동적 높이) ---
        duration_top = detail_y + detail_h - 0.35  # duration 바 시작 y

        aux_items = []
        if step.get("prerequisites"):
            aux_items.append(("전제 조건", step["prerequisites"]))
        if step.get("risks"):
            aux_items.append(("리스크", step["risks"]))
        if step.get("metrics"):
            aux_items.append(("기대 효과", step["metrics"]))
        if step.get("example"):
            aux_items.append(("실증 사례", step["example"]))

        _draw_aux_items(
            c, aux_items,
            x=bx + 0.15, y_start=cy, w=box_w - 0.3,
            y_limit=duration_top,
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
    # quadrant = {"title": "...", "items": [...], "highlight": False,
    #             -- aux (선택) --
    #             "description": "...", "action": "...", "metrics": "..."}
    insight: str  # 하단 인사이트 박스
    footer: SlideFooter


def quadrant_story(slide: Slide, spec: QuadrantSpec):
    """2×2 grid + 양축 라벨 + 하단 인사이트 박스."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(
            spec.intro,
            x=0.3, y=1.20, w=9.4, h=0.30,
            size=10, color="grey_900", anchor="top",
        )

    # ----- 2×2 grid -----
    GRID_X = 1.2
    GRID_Y = 1.7
    GRID_W = 7.4
    cell_gap = 0.12
    cell_w = (GRID_W - cell_gap) / 2
    q_text_w = cell_w - 0.3

    # 동적 셀 높이: 각 사분면 콘텐츠 추정 → 같은 행 최대값 기준
    def _estimate_quad_h(q):
        items = [
            {"text": q["title"], "size": 11, "bold": True, "gap": 0.03},
            {"h": 0.012},  # 구분선
        ]
        for it in q.get("items", []):
            items.append({"text": f"▪  {it}", "size": 9})
        # overhead: stripe(0.08) + top pad(0.15) + bottom pad(0.1)
        return estimate_block_height(items, q_text_w) + 0.33

    quad_heights = [_estimate_quad_h(q) for q in spec.quadrants[:4]]
    # 부족한 사분면은 기본값
    while len(quad_heights) < 4:
        quad_heights.append(1.5)

    # 행별 콘텐츠 비율: 같은 행 2개 중 큰 값 (TL/TR = row0, BL/BR = row1)
    raw_row0 = max(quad_heights[0], quad_heights[1])
    raw_row1 = max(quad_heights[2], quad_heights[3])

    # 가용 공간을 비례 배분 (기본: insight 전까지 채움)
    default_grid_h = 6.4 - GRID_Y - 0.35  # insight y=6.4, axis label ~0.35
    total_raw = raw_row0 + raw_row1
    if total_raw > 0:
        ratio0 = raw_row0 / total_raw
        ratio1 = raw_row1 / total_raw
    else:
        ratio0 = ratio1 = 0.5
    usable = default_grid_h - cell_gap
    row0_h = usable * ratio0
    row1_h = usable * ratio1

    GRID_H = row0_h + cell_gap + row1_h
    row_heights_q = [row0_h, row1_h]

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
        cell_h = row_heights_q[row]
        row_y_offset = sum(row_heights_q[:row]) + row * cell_gap
        cx = GRID_X + col * (cell_w + cell_gap)
        cy = GRID_Y + row_y_offset
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
        n_items = len(q.get("items", []))
        for ii, item in enumerate(q.get("items", [])):
            c.text(
                f"▪  {item}",
                x=cx + 0.2, y=cy + 0.6 + ii * 0.26, w=cell_w - 0.3, h=0.25,
                size=9, color="grey_900", anchor="top",
            )

        # --- aux 콘텐츠 (빈 사분면 공간 채움 — 동적 높이) ---
        aux_start_y = cy + 0.6 + n_items * 0.26 + 0.15
        cell_bottom = cy + cell_h - 0.1

        aux_items = []
        if q.get("description"):
            aux_items.append(("설명", q["description"]))
        if q.get("action"):
            aux_items.append(("실행 과제", q["action"]))
        if q.get("metrics"):
            aux_items.append(("정량 효과", q["metrics"]))

        _draw_aux_items(
            c, aux_items,
            x=cx + 0.2, y_start=aux_start_y, w=cell_w - 0.4,
            y_limit=cell_bottom,
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


# ============================================================
# Pattern 6 — Hub & Spoke
# ============================================================


@dataclass
class HubSpokeSpec:
    header: SlideHeader
    intro: str
    hub: dict  # {"title": "...", "subtitle": "..."}
    spokes: list[dict]
    # spoke = {"title": "...", "detail": "...", "badge": "..."}
    takeaway: str
    footer: SlideFooter


def hub_spoke(slide: Slide, spec: HubSpokeSpec):
    """중심 원 + 방사형 연결 — 아키텍처, 생태계, 역할 구조."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(
            spec.intro,
            x=0.3, y=1.20, w=9.4, h=0.30,
            size=10, color="grey_900", anchor="top",
        )

    import math

    # ----- 중심 원 -----
    HUB_CX, HUB_CY = 5.0, 4.0  # 헤더와 겹침 방지: 3.7→4.0
    HUB_D = 1.6  # 약간 축소: 1.8→1.6
    c.circle(
        x=HUB_CX - HUB_D / 2, y=HUB_CY - HUB_D / 2, d=HUB_D,
        fill="grey_800", border=None,
        text=spec.hub["title"], text_color="white", text_size=13, text_bold=True,
    )
    if spec.hub.get("subtitle"):
        c.text(
            spec.hub["subtitle"],
            x=HUB_CX - 0.8, y=HUB_CY + 0.25, w=1.6, h=0.3,
            size=8, color="grey_200", align="center", anchor="top",
        )

    # ----- Spoke 배치 (원형) -----
    n = len(spec.spokes)
    # spoke 수에 따라 자동 크기 조정
    if n <= 4:
        SPOKE_R, SPOKE_W, SPOKE_H = 2.2, 2.0, 1.1
    elif n <= 5:
        SPOKE_R, SPOKE_W, SPOKE_H = 2.2, 1.9, 1.0
    else:  # 6+
        SPOKE_R, SPOKE_W, SPOKE_H = 2.0, 1.7, 0.85
    angle_offset = -math.pi / 2

    for i, sp in enumerate(spec.spokes):
        angle = angle_offset + (2 * math.pi * i / n)
        sx = HUB_CX + SPOKE_R * math.cos(angle) - SPOKE_W / 2
        sy = HUB_CY + SPOKE_R * math.sin(angle) - SPOKE_H / 2

        line_x1 = HUB_CX + (HUB_D / 2 + 0.05) * math.cos(angle)
        line_y1 = HUB_CY + (HUB_D / 2 + 0.05) * math.sin(angle)
        line_x2 = HUB_CX + (SPOKE_R - SPOKE_W / 2 - 0.05) * math.cos(angle)
        line_y2 = HUB_CY + (SPOKE_R - SPOKE_W / 2 - 0.05) * math.sin(angle)
        c.line(x1=line_x1, y1=line_y1, x2=line_x2, y2=line_y2,
               color="grey_400", width=1.0)

        c.box(x=sx, y=sy, w=SPOKE_W, h=SPOKE_H,
              fill="white", border=0.75, border_color="grey_mid")
        c.box(x=sx, y=sy, w=SPOKE_W, h=0.05,
              fill="grey_700", border=None)

        if sp.get("badge"):
            c.badge(sp["badge"], x=sx + 0.08, y=sy + 0.12,
                    fill="grey_800", text_color="white", size=7)

        c.text(sp["title"],
               x=sx + 0.1, y=sy + 0.35, w=SPOKE_W - 0.2, h=0.28,
               size=10, bold=True, color="grey_900", anchor="top")
        if sp.get("detail"):
            c.text(sp["detail"],
                   x=sx + 0.1, y=sy + 0.62, w=SPOKE_W - 0.2, h=0.4,
                   size=8, color="grey_700", anchor="top")

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 7 — Before / After
# ============================================================


@dataclass
class BeforeAfterSpec:
    header: SlideHeader
    intro: str
    before_title: str
    after_title: str
    before_items: list[dict]
    after_items: list[dict]
    arrow_label: str
    takeaway: str
    footer: SlideFooter


def before_after(slide: Slide, spec: BeforeAfterSpec):
    """좌우 대비 (AS-IS / TO-BE) — 변화, 전환, 개선 표현."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    LEFT_X, RIGHT_X = 0.3, 5.35
    PANEL_W = 4.35
    PANEL_Y = 1.7
    PANEL_H = 4.7

    c.box(x=LEFT_X, y=PANEL_Y, w=PANEL_W, h=PANEL_H,
          fill="grey_200", border=0.75, border_color="grey_mid")
    c.box(x=LEFT_X, y=PANEL_Y, w=PANEL_W, h=0.5,
          fill="grey_700", border=None)
    c.text(spec.before_title,
           x=LEFT_X + 0.2, y=PANEL_Y, w=PANEL_W - 0.4, h=0.5,
           size=14, bold=True, color="white", anchor="middle")

    c.box(x=RIGHT_X, y=PANEL_Y, w=PANEL_W, h=PANEL_H,
          fill="white", border=0.75, border_color="grey_mid")
    c.box(x=RIGHT_X, y=PANEL_Y, w=PANEL_W, h=0.5,
          fill="grey_900", border=None)
    c.text(spec.after_title,
           x=RIGHT_X + 0.2, y=PANEL_Y, w=PANEL_W - 0.4, h=0.5,
           size=14, bold=True, color="white", anchor="middle")

    arrow_y = PANEL_Y + PANEL_H / 2
    c.arrow(x1=LEFT_X + PANEL_W + 0.08, y1=arrow_y,
            x2=RIGHT_X - 0.08, y2=arrow_y,
            color="grey_900", width=2.0)
    if spec.arrow_label:
        c.text(spec.arrow_label,
               x=LEFT_X + PANEL_W + 0.05, y=arrow_y - 0.3, w=0.6, h=0.25,
               size=7, bold=True, color="grey_900", align="center", anchor="bottom")

    item_w = PANEL_W - 0.4
    n_items = max(len(spec.before_items), len(spec.after_items))
    item_h = min(0.9, (PANEL_H - 0.7) / max(n_items, 1))

    for i, item in enumerate(spec.before_items):
        iy = PANEL_Y + 0.6 + i * item_h
        c.text(item["label"],
               x=LEFT_X + 0.2, y=iy, w=item_w, h=0.22,
               size=10, bold=True, color="grey_900", anchor="top")
        if item.get("detail"):
            c.text(item["detail"],
                   x=LEFT_X + 0.2, y=iy + 0.22, w=item_w, h=0.3,
                   size=8, color="grey_700", anchor="top")
        if item.get("kpi"):
            c.text(item["kpi"],
                   x=LEFT_X + 0.2, y=iy + item_h - 0.22, w=item_w, h=0.2,
                   size=9, bold=True, color="negative", anchor="top")
        if i < len(spec.before_items) - 1:
            c.box(x=LEFT_X + 0.2, y=iy + item_h - 0.02, w=item_w, h=0.01,
                  fill="grey_mid", border=None)

    for i, item in enumerate(spec.after_items):
        iy = PANEL_Y + 0.6 + i * item_h
        c.text(item["label"],
               x=RIGHT_X + 0.2, y=iy, w=item_w, h=0.22,
               size=10, bold=True, color="grey_900", anchor="top")
        if item.get("detail"):
            c.text(item["detail"],
                   x=RIGHT_X + 0.2, y=iy + 0.22, w=item_w, h=0.3,
                   size=8, color="grey_700", anchor="top")
        if item.get("kpi"):
            c.text(item["kpi"],
                   x=RIGHT_X + 0.2, y=iy + item_h - 0.22, w=item_w, h=0.2,
                   size=9, bold=True, color="positive", anchor="top")
        if i < len(spec.after_items) - 1:
            c.box(x=RIGHT_X + 0.2, y=iy + item_h - 0.02, w=item_w, h=0.01,
                  fill="grey_mid", border=None)

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 8 — KPI Dashboard
# ============================================================


@dataclass
class KpiDashboardSpec:
    header: SlideHeader
    intro: str
    kpis: list[dict]
    bottom_note: str
    takeaway: str
    footer: SlideFooter


def kpi_dashboard(slide: Slide, spec: KpiDashboardSpec):
    """대형 KPI 카드 배열 — 성과, 지표, 대시보드."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n = len(spec.kpis)
    if n <= 3:
        rows, cols = 1, n
    elif n <= 6:
        cols = min(3, (n + 1) // 2)
        rows = (n + cols - 1) // cols
    else:
        cols, rows = 4, 2

    KPI_Y = 1.7
    gap_val = 0.18
    total_w = 9.4
    total_h = 4.6
    kpi_w = (total_w - gap_val * (cols - 1)) / cols
    kpi_h = (total_h - gap_val * (rows - 1)) / rows

    for i, kpi in enumerate(spec.kpis):
        col = i % cols
        row = i // cols
        kx = 0.3 + col * (kpi_w + gap_val)
        ky = KPI_Y + row * (kpi_h + gap_val)

        c.box(x=kx, y=ky, w=kpi_w, h=kpi_h,
              fill="white", border=0.75, border_color="grey_mid")

        trend = kpi.get("trend", "flat")
        stripe_color = ("positive" if trend == "up"
                        else "negative" if trend == "down"
                        else "grey_700")
        c.box(x=kx, y=ky, w=0.08, h=kpi_h,
              fill=stripe_color, border=None)

        trend_char = "▲" if trend == "up" else ("▼" if trend == "down" else "●")
        c.text(trend_char,
               x=kx + kpi_w - 0.4, y=ky + 0.1, w=0.3, h=0.25,
               size=14, bold=True, color=stripe_color, align="right", anchor="top")

        v_size = 28 if len(kpi["value"]) <= 4 else 22
        c.text(kpi["value"],
               x=kx + 0.25, y=ky + 0.15, w=kpi_w - 0.7, h=kpi_h * 0.35,
               size=v_size, bold=True, color="grey_900",
               font="Georgia", anchor="top")

        c.text(kpi["label"],
               x=kx + 0.25, y=ky + kpi_h * 0.42, w=kpi_w - 0.4, h=0.28,
               size=10, bold=True, color="grey_900", anchor="top")

        if kpi.get("subtitle"):
            c.text(kpi["subtitle"],
                   x=kx + 0.25, y=ky + kpi_h * 0.42 + 0.28, w=kpi_w - 0.4, h=0.2,
                   size=8, color="grey_700", anchor="top")

        if kpi.get("detail"):
            c.text(kpi["detail"],
                   x=kx + 0.25, y=ky + kpi_h - 0.35, w=kpi_w - 0.4, h=0.25,
                   size=8, color="grey_700", anchor="top")

    if spec.bottom_note:
        c.text(spec.bottom_note,
               x=0.3, y=6.35, w=9.4, h=0.18,
               size=7, color="grey_700", anchor="top")

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 9 — Waterfall Bridge
# ============================================================


@dataclass
class WaterfallSpec:
    header: SlideHeader
    intro: str
    start: dict
    steps: list[dict]
    end: dict
    unit: str
    takeaway: str
    footer: SlideFooter


def waterfall_bridge(slide: Slide, spec: WaterfallSpec):
    """증감 바 차트 (Bridge) — 비용/매출/시간 분해."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    all_items = [spec.start] + spec.steps + [spec.end]
    n = len(all_items)
    bar_gap = 0.1
    bar_area_x = 0.8
    bar_area_w = 8.6
    bar_w = (bar_area_w - bar_gap * (n - 1)) / n

    cumulative = [spec.start["value"]]
    for s in spec.steps:
        cumulative.append(cumulative[-1] + s["value"])
    all_values = cumulative + [spec.end["value"]]
    max_val = max(abs(v) for v in all_values) if all_values else 1

    CHART_Y = 1.7
    CHART_H = 4.4  # 3.8→4.4 (하단 빈 공간 해소)
    BASELINE_Y = CHART_Y + CHART_H * 0.72  # 0.75→0.72

    def val_to_y(v):
        if max_val == 0:
            return BASELINE_Y
        scale = (CHART_H * 0.65) / max_val
        return BASELINE_Y - v * scale

    c.line(x1=bar_area_x - 0.1, y1=BASELINE_Y,
           x2=bar_area_x + bar_area_w + 0.1, y2=BASELINE_Y,
           color="grey_mid", width=0.5)

    running = spec.start["value"]
    for i, item in enumerate(all_items):
        bx = bar_area_x + i * (bar_w + bar_gap)
        is_start = (i == 0)
        is_end = (i == n - 1)

        if is_start or is_end:
            val = item["value"]
            bar_top = val_to_y(val)
            bar_bottom = BASELINE_Y
            bh = abs(bar_bottom - bar_top)
            c.box(x=bx, y=min(bar_top, bar_bottom), w=bar_w, h=max(bh, 0.05),
                  fill="grey_800", border=None)
        else:
            step_val = item["value"]
            prev_cum = running
            new_cum = running + step_val
            running = new_cum
            bar_top = val_to_y(max(prev_cum, new_cum))
            bar_bottom = val_to_y(min(prev_cum, new_cum))
            bh = abs(bar_bottom - bar_top)
            fill = "positive" if step_val > 0 else "negative"
            c.box(x=bx, y=bar_top, w=bar_w, h=max(bh, 0.05),
                  fill=fill, border=None)

            if i > 0:
                prev_x = bar_area_x + (i - 1) * (bar_w + bar_gap) + bar_w
                conn_y = val_to_y(prev_cum)
                c.line(x1=prev_x, y1=conn_y, x2=bx, y2=conn_y,
                       color="grey_400", width=0.5)

        c.text(item["label"],
               x=bx, y=CHART_Y + CHART_H + 0.05, w=bar_w, h=0.35,
               size=8, bold=True, color="grey_900", align="center", anchor="top")

        val = item["value"]
        is_step = not is_start and not is_end
        prefix = "+" if is_step and val > 0 else ""
        # 소수점이 있으면 1자리, 정수면 0자리
        fmt = ",.1f" if isinstance(val, float) and val != int(val) else ",.0f"
        val_str = f"{prefix}{val:{fmt}}{spec.unit}" if isinstance(val, (int, float)) else str(val)
        if is_start or is_end:
            vy = val_to_y(val) - 0.25
        elif val >= 0:
            vy = val_to_y(max(running, running - val)) - 0.25
        else:
            vy = val_to_y(min(running, running - val)) + 0.02
        c.text(val_str,
               x=bx, y=vy, w=bar_w, h=0.22,
               size=9, bold=True, color="grey_900", align="center", anchor="top")

        if item.get("detail"):
            c.text(item["detail"],
                   x=bx, y=CHART_Y + CHART_H + 0.38, w=bar_w, h=0.25,
                   size=7, color="grey_700", align="center", anchor="top")

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 10 — Swimlane
# ============================================================


@dataclass
class SwimlaneSpec:
    header: SlideHeader
    intro: str
    lanes: list[str]
    phases: list[str]
    activities: list[dict]
    takeaway: str
    footer: SlideFooter


def swimlane(slide: Slide, spec: SwimlaneSpec):
    """부서x단계 그리드 — Cross-functional 프로세스."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n_lanes = len(spec.lanes)
    n_phases = len(spec.phases)

    LANE_LABEL_W = 1.4
    GRID_X = 0.3 + LANE_LABEL_W
    GRID_Y = 1.7
    GRID_W = 9.4 - LANE_LABEL_W
    GRID_H = 4.7
    PHASE_HEADER_H = 0.45

    col_gap = 0.08
    row_gap = 0.08
    col_w = (GRID_W - col_gap * (n_phases - 1)) / n_phases
    row_h = (GRID_H - PHASE_HEADER_H - row_gap * (n_lanes - 1)) / n_lanes

    fills = ["grey_800", "grey_700", "grey_400", "grey_200", "grey_400"]
    text_colors_list = ["white", "white", "white", "grey_900", "white"]
    for j, phase in enumerate(spec.phases):
        px = GRID_X + j * (col_w + col_gap)
        c.box(x=px, y=GRID_Y, w=col_w, h=PHASE_HEADER_H,
              fill=fills[j % len(fills)], border=None)
        c.text(phase,
               x=px, y=GRID_Y, w=col_w, h=PHASE_HEADER_H,
               size=10, bold=True, color=text_colors_list[j % len(text_colors_list)],
               align="center", anchor="middle")

    body_y = GRID_Y + PHASE_HEADER_H + 0.08
    for i, lane in enumerate(spec.lanes):
        ly = body_y + i * (row_h + row_gap)
        c.box(x=0.3, y=ly, w=LANE_LABEL_W - 0.08, h=row_h,
              fill="grey_100", border=0.5, border_color="grey_mid")
        c.text(lane,
               x=0.35, y=ly, w=LANE_LABEL_W - 0.18, h=row_h,
               size=9, bold=True, color="grey_900", anchor="middle")
        for j in range(n_phases):
            px = GRID_X + j * (col_w + col_gap)
            c.box(x=px, y=ly, w=col_w, h=row_h,
                  fill="white", border=0.5, border_color="grey_mid")

    for act in spec.activities:
        lane_idx = act["lane"]
        phase_idx = act["phase"]
        if lane_idx >= n_lanes or phase_idx >= n_phases:
            continue
        ly = body_y + lane_idx * (row_h + row_gap)
        px = GRID_X + phase_idx * (col_w + col_gap)
        is_hl = act.get("highlight", False)
        pad = 0.06
        c.box(x=px + pad, y=ly + pad, w=col_w - 2 * pad, h=row_h - 2 * pad,
              fill="grey_200" if is_hl else "grey_100",
              border=0.75 if is_hl else 0.5,
              border_color="grey_800" if is_hl else "grey_mid")
        c.text(act["text"],
               x=px + pad + 0.05, y=ly + pad, w=col_w - 2 * pad - 0.1,
               h=row_h - 2 * pad,
               size=8, bold=is_hl, color="grey_900", anchor="middle",
               align="center")

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 11 — Pyramid Layers
# ============================================================


@dataclass
class PyramidSpec:
    header: SlideHeader
    intro: str
    layers: list[dict]
    side_notes: list[dict]
    takeaway: str
    footer: SlideFooter


def pyramid_layers(slide: Slide, spec: PyramidSpec):
    """3~5단 피라미드 — 전략 위계, 성숙도 모델."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n = len(spec.layers)
    PYR_X = 0.8
    PYR_W_MAX = 6.0
    PYR_W_MIN = 2.0
    PYR_Y = 1.7
    PYR_H = 4.5
    layer_gap = 0.06
    layer_h = (PYR_H - layer_gap * (n - 1)) / n

    layer_fills = ["grey_900", "grey_800", "grey_700", "grey_400", "grey_200"]
    layer_text_colors = ["white", "white", "white", "white", "grey_900"]

    for i, layer in enumerate(spec.layers):
        ratio = i / max(n - 1, 1)
        lw = PYR_W_MIN + (PYR_W_MAX - PYR_W_MIN) * ratio
        lx = PYR_X + (PYR_W_MAX - lw) / 2
        ly = PYR_Y + i * (layer_h + layer_gap)

        c.box(x=lx, y=ly, w=lw, h=layer_h,
              fill=layer_fills[i % len(layer_fills)], border=None)

        if layer.get("badge"):
            c.text(layer["badge"],
                   x=lx + 0.15, y=ly + 0.08, w=0.5, h=0.2,
                   size=8, bold=True,
                   color=layer_text_colors[i % len(layer_text_colors)],
                   anchor="top")

        c.text(layer["title"],
               x=lx + 0.15, y=ly + layer_h * 0.15, w=lw - 0.3, h=layer_h * 0.4,
               size=12, bold=True,
               color=layer_text_colors[i % len(layer_text_colors)],
               align="center", anchor="middle")

        if layer.get("detail"):
            c.text(layer["detail"],
                   x=lx + 0.15, y=ly + layer_h * 0.55, w=lw - 0.3, h=layer_h * 0.4,
                   size=8,
                   color=layer_text_colors[i % len(layer_text_colors)],
                   align="center", anchor="top")

    if spec.side_notes:
        NOTE_X = PYR_X + PYR_W_MAX + 0.3
        NOTE_W = 9.7 - NOTE_X
        c.section_label("핵심 지표", x=NOTE_X, y=PYR_Y, w=NOTE_W, size=10)
        for si, note in enumerate(spec.side_notes):
            ny = PYR_Y + 0.35 + si * 0.7
            c.stat_block(
                value=note["value"], label=note["label"],
                x=NOTE_X, y=ny, w=NOTE_W, h=0.6, align="left",
            )

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 12 — Data Narrative
# ============================================================


@dataclass
class DataNarrativeSpec:
    header: SlideHeader
    intro: str
    chart_title: str
    chart_data: list[dict]
    chart_unit: str
    narratives: list[dict]
    takeaway: str
    footer: SlideFooter


def data_narrative(slide: Slide, spec: DataNarrativeSpec):
    """좌 차트(수평 바) + 우 해설 — 분석+인사이트."""
    c = Canvas(slide)
    _draw_header(c, spec.header)

    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    CHART_X = 0.3
    CHART_Y = 1.7
    CHART_W = 5.0
    CHART_H = 4.8

    c.box(x=CHART_X, y=CHART_Y, w=CHART_W, h=CHART_H,
          fill="white", border=0.75, border_color="grey_mid")
    c.text(spec.chart_title,
           x=CHART_X + 0.2, y=CHART_Y + 0.12, w=CHART_W - 0.4, h=0.28,
           size=11, bold=True, color="grey_900", anchor="top")

    n_bars = len(spec.chart_data)
    bar_area_y = CHART_Y + 0.5
    bar_area_h = CHART_H - 0.7
    bar_h = min(0.4, (bar_area_h - 0.1 * (n_bars - 1)) / max(n_bars, 1))
    bar_gap_val = ((bar_area_h - bar_h * n_bars) / max(n_bars - 1, 1)
                   if n_bars > 1 else 0)

    max_val = max((d["value"] for d in spec.chart_data), default=1)
    BAR_MAX_W = CHART_W - 1.8

    for i, d in enumerate(spec.chart_data):
        by = bar_area_y + i * (bar_h + bar_gap_val)
        is_hl = d.get("highlight", False)

        c.text(d["label"],
               x=CHART_X + 0.15, y=by, w=1.2, h=bar_h,
               size=8, bold=True, color="grey_900", anchor="middle")

        bw = BAR_MAX_W * (d["value"] / max_val) if max_val > 0 else 0
        fill = "grey_900" if is_hl else "grey_400"
        c.box(x=CHART_X + 1.4, y=by + 0.04, w=max(bw, 0.05), h=bar_h - 0.08,
              fill=fill, border=None)

        val_str = (f"{d['value']:,.0f}{spec.chart_unit}"
                   if isinstance(d["value"], (int, float)) else str(d["value"]))
        c.text(val_str,
               x=CHART_X + 1.4 + bw + 0.08, y=by, w=0.8, h=bar_h,
               size=9, bold=is_hl, color="grey_900", anchor="middle")

    NARR_X = 5.5
    NARR_W = 4.2
    c.section_label("분석 인사이트", x=NARR_X, y=CHART_Y, w=NARR_W, size=10)

    narr_y = CHART_Y + 0.4
    narr_item_h = (CHART_H - 0.5) / max(len(spec.narratives), 1)

    for i, narr in enumerate(spec.narratives):
        ny = narr_y + i * narr_item_h
        c.callout_box(
            x=NARR_X, y=ny, w=NARR_W, h=narr_item_h - 0.1,
            title=narr["title"],
            body=narr.get("detail", ""),
            bar_color="grey_800" if i == 0 else "grey_400",
            title_size=10, body_size=8,
        )

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 13 — Maturity Model (수평 성숙도)
# ============================================================


@dataclass
class MaturityModelSpec:
    header: SlideHeader
    intro: str
    stages: list[dict]
    current: int
    target: int
    takeaway: str
    footer: SlideFooter


def maturity_model(slide: Slide, spec: MaturityModelSpec):
    """수평 성숙도 모델 — 현재 위치 + 목표 마커."""
    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n = len(spec.stages)
    STAGE_Y = 1.7
    # 마커(0.3) + 화살표(0.3) + takeaway 전 여백
    STAGE_H = 6.55 - STAGE_Y - 0.75  # = 4.3" (빈 공간 해소)
    gap = 0.1
    stage_w = (9.4 - gap * (n - 1)) / n
    fills = ["grey_200", "grey_400", "grey_700", "grey_800", "grey_900"]
    tcols = ["grey_900", "grey_900", "white", "white", "white"]

    for i, st in enumerate(spec.stages):
        sx = 0.3 + i * (stage_w + gap)
        c.box(x=sx, y=STAGE_Y, w=stage_w, h=STAGE_H,
              fill=fills[i % len(fills)], border=None)
        if st.get("level"):
            c.text(st["level"], x=sx + 0.1, y=STAGE_Y + 0.12, w=0.5, h=0.2,
                   size=8, bold=True, color=tcols[i % len(tcols)], anchor="top")
        c.text(st["title"], x=sx + 0.1, y=STAGE_Y + 0.4, w=stage_w - 0.2, h=0.35,
               size=12, bold=True, color=tcols[i % len(tcols)], anchor="top")
        if st.get("detail"):
            c.text(st["detail"], x=sx + 0.1, y=STAGE_Y + 0.8,
                   w=stage_w - 0.2, h=STAGE_H - 1.0,
                   size=8, color=tcols[i % len(tcols)], anchor="top")
        if i == spec.current:
            c.text("▲ 현재", x=sx + stage_w / 2 - 0.4, y=STAGE_Y + STAGE_H + 0.08,
                   w=0.8, h=0.25, size=9, bold=True, color="negative",
                   align="center", anchor="top")
        if i == spec.target:
            c.text("★ 목표", x=sx + stage_w / 2 - 0.4, y=STAGE_Y + STAGE_H + 0.08,
                   w=0.8, h=0.25, size=9, bold=True, color="positive",
                   align="center", anchor="top")

    if spec.current < spec.target:
        arr_y = STAGE_Y + STAGE_H + 0.4
        c.arrow(x1=0.3 + spec.current * (stage_w + gap) + stage_w / 2, y1=arr_y,
                x2=0.3 + spec.target * (stage_w + gap) + stage_w / 2, y2=arr_y,
                color="grey_900", width=1.5)

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 14 — Milestone Timeline
# ============================================================


@dataclass
class MilestoneTimelineSpec:
    header: SlideHeader
    intro: str
    milestones: list[dict]
    takeaway: str
    footer: SlideFooter


def milestone_timeline(slide: Slide, spec: MilestoneTimelineSpec):
    """수평선 + 위아래 교대 마일스톤 — 히스토리, 일정."""
    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n = len(spec.milestones)
    LINE_Y, LINE_X, LINE_W = 3.8, 0.8, 8.4
    ms_gap = LINE_W / max(n - 1, 1) if n > 1 else LINE_W
    c.line(x1=LINE_X, y1=LINE_Y, x2=LINE_X + LINE_W, y2=LINE_Y,
           color="grey_700", width=2.0)

    CARD_W = min(1.8, ms_gap - 0.15)
    CARD_H = 1.5

    for i, ms in enumerate(spec.milestones):
        mx = LINE_X + i * ms_gap
        is_above = (i % 2 == 0)
        is_hl = ms.get("highlight", False)
        if is_above:
            c.line(x1=mx, y1=LINE_Y, x2=mx, y2=LINE_Y - 0.6,
                   color="grey_700" if is_hl else "grey_400", width=1.0)
            card_y = LINE_Y - 0.6 - CARD_H
        else:
            c.line(x1=mx, y1=LINE_Y, x2=mx, y2=LINE_Y + 0.6,
                   color="grey_700" if is_hl else "grey_400", width=1.0)
            card_y = LINE_Y + 0.6
        c.circle(x=mx - 0.1, y=LINE_Y - 0.1, d=0.2,
                 fill="grey_900" if is_hl else "grey_400",
                 border=None, text="", text_size=1)
        card_x = mx - CARD_W / 2
        c.box(x=card_x, y=card_y, w=CARD_W, h=CARD_H,
              fill="grey_200" if is_hl else "white",
              border=0.75, border_color="grey_800" if is_hl else "grey_mid")
        c.text(ms["date"], x=card_x + 0.08, y=card_y + 0.08,
               w=CARD_W - 0.16, h=0.22, size=8, bold=True, color="grey_700", anchor="top")
        c.text(ms["title"], x=card_x + 0.08, y=card_y + 0.32,
               w=CARD_W - 0.16, h=0.35, size=10, bold=True, color="grey_900", anchor="top")
        if ms.get("detail"):
            c.text(ms["detail"], x=card_x + 0.08, y=card_y + 0.7,
                   w=CARD_W - 0.16, h=0.7, size=8, color="grey_700", anchor="top")

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 15 — RAG Status Table
# ============================================================


@dataclass
class RagStatusSpec:
    header: SlideHeader
    intro: str
    columns: list[str]
    rows: list[dict]
    takeaway: str
    footer: SlideFooter


def rag_status_table(slide: Slide, spec: RagStatusSpec):
    """RAG 색상 상태표 — PMO 대시보드."""
    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n_cols = len(spec.columns)
    n_rows = len(spec.rows)
    TABLE_X, TABLE_Y, TABLE_W = 0.3, 1.7, 9.4
    HEADER_H = 0.45
    NAME_COL_W = 2.5
    RAG_COL_W = (TABLE_W - NAME_COL_W) / max(n_cols - 1, 1)
    avail_h = 6.55 - TABLE_Y - HEADER_H - 0.15
    ROW_H = min(0.6, avail_h / max(n_rows, 1))
    rag_colors = {"G": "positive", "A": "#F5A623", "R": "negative", "-": "grey_400"}

    c.box(x=TABLE_X, y=TABLE_Y, w=NAME_COL_W, h=HEADER_H, fill="grey_800", border=None)
    c.text(spec.columns[0], x=TABLE_X + 0.1, y=TABLE_Y, w=NAME_COL_W - 0.2,
           h=HEADER_H, size=10, bold=True, color="white", anchor="middle")
    for j in range(1, n_cols):
        hx = TABLE_X + NAME_COL_W + (j - 1) * RAG_COL_W
        c.box(x=hx, y=TABLE_Y, w=RAG_COL_W, h=HEADER_H, fill="grey_800", border=None)
        c.text(spec.columns[j], x=hx, y=TABLE_Y, w=RAG_COL_W, h=HEADER_H,
               size=9, bold=True, color="white", align="center", anchor="middle")

    for i, row in enumerate(spec.rows):
        ry = TABLE_Y + HEADER_H + i * ROW_H
        rf = "white" if i % 2 == 0 else "grey_100"
        c.box(x=TABLE_X, y=ry, w=NAME_COL_W, h=ROW_H, fill=rf,
              border=0.5, border_color="grey_mid")
        c.text(row["name"], x=TABLE_X + 0.1, y=ry, w=NAME_COL_W - 0.2, h=ROW_H,
               size=9, bold=True, color="grey_900", anchor="middle")
        for j, val in enumerate(row.get("values", [])):
            cx = TABLE_X + NAME_COL_W + j * RAG_COL_W
            c.box(x=cx, y=ry, w=RAG_COL_W, h=ROW_H, fill=rf,
                  border=0.5, border_color="grey_mid")
            rag_d = 0.28
            c.circle(x=cx + RAG_COL_W / 2 - rag_d / 2, y=ry + ROW_H / 2 - rag_d / 2,
                     d=rag_d, fill=rag_colors.get(val, "grey_400"),
                     border=None, text="", text_size=1)

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 16 — Architecture Stack
# ============================================================


@dataclass
class ArchStackSpec:
    header: SlideHeader
    intro: str
    layers: list[dict]
    side_label: str
    takeaway: str
    footer: SlideFooter


def architecture_stack(slide: Slide, spec: ArchStackSpec):
    """수평 레이어 스택 — 기술 아키텍처."""
    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n = len(spec.layers)
    STACK_X, STACK_W, STACK_Y, STACK_H = 0.5, 8.0, 1.7, 4.7
    layer_gap = 0.06
    layer_h = (STACK_H - layer_gap * (n - 1)) / n
    lfills = ["grey_900", "grey_800", "grey_700", "grey_400", "grey_200", "grey_100"]
    ltxt = ["white", "white", "white", "white", "grey_900", "grey_900"]

    for i, layer in enumerate(spec.layers):
        ly = STACK_Y + i * (layer_h + layer_gap)
        c.box(x=STACK_X, y=ly, w=STACK_W, h=layer_h,
              fill=lfills[i % len(lfills)], border=None)
        if layer.get("badge"):
            c.text(layer["badge"], x=STACK_X + 0.15, y=ly + 0.05, w=0.8, h=0.2,
                   size=7, bold=True, color=ltxt[i % len(ltxt)], anchor="top")
        c.text(layer["title"], x=STACK_X + 0.15, y=ly + layer_h * 0.2,
               w=2.0, h=layer_h * 0.5, size=11, bold=True,
               color=ltxt[i % len(ltxt)], anchor="middle")
        for ji, item in enumerate(layer.get("items", [])):
            iw = (STACK_W - 2.5) / max(len(layer.get("items", [])), 1)
            c.text(item, x=STACK_X + 2.3 + ji * iw, y=ly, w=iw, h=layer_h,
                   size=8, color=ltxt[i % len(ltxt)], align="center", anchor="middle")

    if spec.side_label:
        c.text(spec.side_label, x=STACK_X + STACK_W + 0.2,
               y=STACK_Y + STACK_H * 0.3, w=1.0, h=STACK_H * 0.4,
               size=9, bold=True, color="grey_700", align="center", anchor="middle")
        c.arrow(x1=STACK_X + STACK_W + 0.7, y1=STACK_Y + 0.3,
                x2=STACK_X + STACK_W + 0.7, y2=STACK_Y + STACK_H - 0.3,
                color="grey_400", width=1.0)

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 17 — Gantt Roadmap
# ============================================================


@dataclass
class GanttSpec:
    header: SlideHeader
    intro: str
    phases: list[str]
    streams: list[dict]
    milestones: list[dict]
    takeaway: str
    footer: SlideFooter


def gantt_roadmap(slide: Slide, spec: GanttSpec):
    """멀티 스트림 Gantt — 프로그램 로드맵."""
    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n_phases = len(spec.phases)
    n_streams = len(spec.streams)
    LABEL_W = 1.8
    GRID_X = 0.3 + LABEL_W
    GRID_Y, GRID_W, GRID_H = 1.7, 9.4 - LABEL_W, 4.7
    PHASE_H = 0.4
    col_w = GRID_W / n_phases
    body_y = GRID_Y + PHASE_H + 0.08
    body_h = GRID_H - PHASE_H - 0.08
    row_h = body_h / max(n_streams, 1)

    pfills = ["grey_800", "grey_700", "grey_400", "grey_200"]
    ptxt = ["white", "white", "white", "grey_900"]
    for j, phase in enumerate(spec.phases):
        px = GRID_X + j * col_w
        c.box(x=px, y=GRID_Y, w=col_w, h=PHASE_H,
              fill=pfills[j % len(pfills)], border=None)
        c.text(phase, x=px, y=GRID_Y, w=col_w, h=PHASE_H,
               size=9, bold=True, color=ptxt[j % len(ptxt)], align="center", anchor="middle")
    for j in range(1, n_phases):
        c.line(x1=GRID_X + j * col_w, y1=body_y, x2=GRID_X + j * col_w, y2=body_y + body_h,
               color="grey_mid", width=0.3)

    for i, stream in enumerate(spec.streams):
        sy = body_y + i * row_h
        sf = "grey_100" if i % 2 == 0 else "white"
        c.box(x=0.3, y=sy, w=LABEL_W - 0.08, h=row_h, fill=sf,
              border=0.5, border_color="grey_mid")
        c.text(stream["name"], x=0.35, y=sy, w=LABEL_W - 0.18, h=row_h,
               size=9, bold=True, color="grey_900", anchor="middle")
        c.box(x=GRID_X, y=sy, w=GRID_W, h=row_h, fill=sf,
              border=0.5, border_color="grey_mid")
        bp = row_h * 0.2
        for bar in stream.get("bars", []):
            bx = GRID_X + bar["start"] * col_w + 0.04
            bw = (bar["end"] - bar["start"]) * col_w - 0.08
            is_hl = bar.get("highlight", False)
            c.box(x=bx, y=sy + bp, w=bw, h=row_h - 2 * bp,
                  fill="grey_900" if is_hl else "grey_400", border=None)
            if bar.get("label"):
                c.text(bar["label"], x=bx + 0.05, y=sy + bp,
                       w=bw - 0.1, h=row_h - 2 * bp,
                       size=8, bold=is_hl, color="white", anchor="middle")

    for ms in spec.milestones:
        mx = GRID_X + ms["phase"] * col_w
        c.line(x1=mx, y1=body_y - 0.05, x2=mx, y2=body_y + body_h + 0.05,
               color="negative", width=1.5)
        c.text(ms.get("label", ""), x=mx - 0.5, y=body_y + body_h + 0.08,
               w=1.0, h=0.2, size=7, bold=True, color="negative",
               align="center", anchor="top")

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 18 — Cycle Diagram (원형 순환)
# ============================================================


@dataclass
class CycleSpec:
    header: SlideHeader
    intro: str
    center: dict
    stages: list[dict]
    takeaway: str
    footer: SlideFooter


def cycle_diagram(slide: Slide, spec: CycleSpec):
    """원형 순환 다이어그램 — PDCA, 플라이휠."""
    import math
    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    CX, CY = 5.0, 4.0
    n = len(spec.stages)
    CENTER_D = 1.4
    c.circle(x=CX - CENTER_D / 2, y=CY - CENTER_D / 2, d=CENTER_D,
             fill="grey_900", border=None,
             text=spec.center["title"], text_color="white", text_size=12, text_bold=True)
    if spec.center.get("subtitle"):
        c.text(spec.center["subtitle"], x=CX - 0.6, y=CY + 0.2, w=1.2, h=0.25,
               size=7, color="grey_200", align="center", anchor="top")

    RING_R = 1.3
    cfills = ["grey_800", "grey_700", "grey_400", "grey_200", "grey_400", "grey_700"]
    ctxt = ["white", "white", "white", "grey_900", "white", "white"]
    CARD_R, CARD_W, CARD_H = 2.5, 1.8, 0.9
    angle_offset = -math.pi / 2

    for i, st in enumerate(spec.stages):
        angle = angle_offset + (2 * math.pi * i / n)
        next_angle = angle_offset + (2 * math.pi * (i + 1) / n)
        mid_angle = (angle + next_angle) / 2

        seg_x = CX + RING_R * math.cos(mid_angle) - 0.35
        seg_y = CY + RING_R * math.sin(mid_angle) - 0.18
        c.box(x=seg_x, y=seg_y, w=0.7, h=0.36,
              fill=cfills[i % len(cfills)], border=None)
        if st.get("badge"):
            c.text(st["badge"], x=seg_x, y=seg_y, w=0.7, h=0.36,
                   size=8, bold=True, color=ctxt[i % len(ctxt)],
                   align="center", anchor="middle")

        arr_a = angle_offset + (2 * math.pi * (i + 0.75) / n)
        a2_a = angle_offset + (2 * math.pi * (i + 1.05) / n)
        c.arrow(x1=CX + (RING_R + 0.25) * math.cos(arr_a),
                y1=CY + (RING_R + 0.25) * math.sin(arr_a),
                x2=CX + (RING_R + 0.25) * math.cos(a2_a),
                y2=CY + (RING_R + 0.25) * math.sin(a2_a),
                color="grey_400", width=0.75)

        card_cx = CX + CARD_R * math.cos(mid_angle)
        card_cy = CY + CARD_R * math.sin(mid_angle)
        card_x = card_cx - CARD_W / 2
        card_y = card_cy - CARD_H / 2
        lx1 = CX + (RING_R + 0.4) * math.cos(mid_angle)
        ly1 = CY + (RING_R + 0.4) * math.sin(mid_angle)
        c.line(x1=lx1, y1=ly1, x2=card_cx, y2=card_cy,
               color="grey_mid", width=0.5)
        c.box(x=card_x, y=card_y, w=CARD_W, h=CARD_H,
              fill="white", border=0.75, border_color="grey_mid")
        c.text(st["title"], x=card_x + 0.08, y=card_y + 0.08,
               w=CARD_W - 0.16, h=0.28, size=10, bold=True, color="grey_900", anchor="top")
        if st.get("detail"):
            c.text(st["detail"], x=card_x + 0.08, y=card_y + 0.38,
                   w=CARD_W - 0.16, h=0.45, size=8, color="grey_700", anchor="top")

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 19 — Value Chain (Porter)
# ============================================================


@dataclass
class ValueChainSpec:
    header: SlideHeader
    intro: str
    primary: list[dict]
    support: list[dict]
    margin_label: str
    takeaway: str
    footer: SlideFooter


def value_chain(slide: Slide, spec: ValueChainSpec):
    """Porter 가치사슬 — 주요활동 체브론 + 지원활동."""
    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n_primary = len(spec.primary)
    n_support = len(spec.support)
    SUPPORT_X, SUPPORT_Y, SUPPORT_W = 0.3, 1.7, 7.8
    support_h, support_gap = 0.55, 0.06

    for i, sup in enumerate(spec.support):
        sy = SUPPORT_Y + i * (support_h + support_gap)
        c.box(x=SUPPORT_X, y=sy, w=SUPPORT_W, h=support_h,
              fill="grey_100", border=0.5, border_color="grey_mid")
        c.text(sup["title"], x=SUPPORT_X + 0.15, y=sy, w=2.0, h=support_h,
               size=9, bold=True, color="grey_900", anchor="middle")
        if sup.get("detail"):
            c.text(sup["detail"], x=SUPPORT_X + 2.2, y=sy,
                   w=SUPPORT_W - 2.4, h=support_h,
                   size=8, color="grey_700", anchor="middle")

    primary_y = SUPPORT_Y + n_support * (support_h + support_gap) + 0.15
    primary_h = 6.4 - primary_y - 0.1
    chev_overlap = 0.1
    chev_w = (SUPPORT_W + chev_overlap * (n_primary - 1)) / n_primary
    chev_fills = ["grey_800", "grey_700", "grey_400", "grey_200", "grey_400"]
    chev_txt = ["white", "white", "white", "grey_900", "white"]

    for i, prim in enumerate(spec.primary):
        px = SUPPORT_X + i * (chev_w - chev_overlap)
        is_hl = prim.get("highlight", False)
        fill = "grey_900" if is_hl else chev_fills[i % len(chev_fills)]
        tc = "white" if is_hl else chev_txt[i % len(chev_txt)]
        c.chevron(x=px, y=primary_y, w=chev_w, h=0.5,
                  fill=fill, text=prim["title"], text_color=tc, text_size=9)
        if prim.get("detail"):
            dx = SUPPORT_X + i * (SUPPORT_W / n_primary)
            dw = SUPPORT_W / n_primary - 0.08
            c.box(x=dx, y=primary_y + 0.55, w=dw, h=primary_h - 0.6,
                  fill="white", border=0.5, border_color="grey_mid")
            c.text(prim["detail"], x=dx + 0.08, y=primary_y + 0.6,
                   w=dw - 0.16, h=primary_h - 0.7,
                   size=8, color="grey_900", anchor="top")

    margin_x = SUPPORT_X + SUPPORT_W + 0.1
    margin_h = primary_y + primary_h - SUPPORT_Y
    c.box(x=margin_x, y=SUPPORT_Y, w=1.2, h=margin_h, fill="grey_800", border=None)
    c.text(spec.margin_label, x=margin_x, y=SUPPORT_Y + margin_h * 0.3,
           w=1.2, h=margin_h * 0.4, size=12, bold=True, color="white",
           align="center", anchor="middle")

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 20 — Bubble Chart (matplotlib)
# ============================================================


@dataclass
class BubbleChartSpec:
    header: SlideHeader
    intro: str
    x_label: str
    y_label: str
    bubbles: list[dict]
    narratives: list[dict]
    takeaway: str
    footer: SlideFooter


def bubble_chart(slide: Slide, spec: BubbleChartSpec):
    """버블 차트 (matplotlib PNG) + 우측 해설."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
    import tempfile, os
    # 한글 폰트 설정 (Windows: 맑은 고딕)
    for fname in ["Malgun Gothic", "맑은 고딕", "Arial"]:
        if any(fname in f.name for f in fm.fontManager.ttflist):
            plt.rcParams["font.family"] = fname
            break
    plt.rcParams["axes.unicode_minus"] = False

    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    fig, ax = plt.subplots(figsize=(5.5, 4.5))
    for b in spec.bubbles:
        is_hl = b.get("highlight", False)
        ax.scatter(b["x"], b["y"], s=b["size"] * 50,
                   c="#2E333A" if is_hl else "#9AA0A8",
                   alpha=0.85 if is_hl else 0.5,
                   edgecolors="#4A4F58", linewidth=1)
        ax.annotate(b["label"], (b["x"], b["y"]),
                    textcoords="offset points", xytext=(0, 8), ha="center",
                    fontsize=8, fontweight="bold" if is_hl else "normal")
    ax.set_xlabel(spec.x_label, fontsize=10)
    ax.set_ylabel(spec.y_label, fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()

    tmp_path = tempfile.mktemp(suffix=".png")
    fig.savefig(tmp_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    from pptx.util import Inches as _In
    slide.shapes.add_picture(tmp_path, _In(0.3), _In(1.7), _In(5.5), _In(4.5))
    try:
        os.unlink(tmp_path)
    except OSError:
        pass  # Windows lock — 임시 파일은 OS가 정리

    NARR_X, NARR_W = 6.0, 3.7
    c.section_label("분석 포인트", x=NARR_X, y=1.7, w=NARR_W, size=10)
    narr_y = 2.1
    nh = 4.3 / max(len(spec.narratives), 1)
    for i, narr in enumerate(spec.narratives):
        c.callout_box(x=NARR_X, y=narr_y + i * nh, w=NARR_W, h=nh - 0.1,
                      title=narr["title"], body=narr.get("detail", ""),
                      bar_color="grey_800" if i == 0 else "grey_400",
                      title_size=10, body_size=8)

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 21 — Tree Diagram (MECE Issue Tree)
# ============================================================


@dataclass
class TreeSpec:
    header: SlideHeader
    intro: str
    root: dict
    branches: list[dict]
    takeaway: str
    footer: SlideFooter


def tree_diagram(slide: Slide, spec: TreeSpec):
    """계층 분기 트리 — Issue Tree, 조직도."""
    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n_br = len(spec.branches)
    ROOT_X, ROOT_Y, ROOT_W, ROOT_H = 0.4, 2.8, 2.0, 0.8
    c.box(x=ROOT_X, y=ROOT_Y, w=ROOT_W, h=ROOT_H, fill="grey_900", border=None)
    c.text(spec.root["title"], x=ROOT_X + 0.1, y=ROOT_Y, w=ROOT_W - 0.2, h=ROOT_H,
           size=12, bold=True, color="white", anchor="middle")

    BRANCH_X, BRANCH_W = ROOT_X + ROOT_W + 0.6, 2.2
    avail_h, branch_gap = 4.5, 0.12
    branch_h = (avail_h - branch_gap * (n_br - 1)) / max(n_br, 1)
    branch_start_y = 1.7

    for i, br in enumerate(spec.branches):
        by = branch_start_y + i * (branch_h + branch_gap)
        is_hl = br.get("highlight", False)
        c.line(x1=ROOT_X + ROOT_W, y1=ROOT_Y + ROOT_H / 2,
               x2=BRANCH_X, y2=by + branch_h / 2, color="grey_700", width=0.75)
        c.box(x=BRANCH_X, y=by, w=BRANCH_W, h=branch_h,
              fill="grey_800" if is_hl else "grey_200",
              border=0.75 if is_hl else 0.5,
              border_color="grey_900" if is_hl else "grey_mid")
        c.text(br["title"], x=BRANCH_X + 0.1, y=by + 0.05, w=BRANCH_W - 0.2, h=0.3,
               size=10, bold=True, color="white" if is_hl else "grey_900", anchor="top")

        children = br.get("children", [])
        if children:
            LEAF_X = BRANCH_X + BRANCH_W + 0.5
            LEAF_W = 9.4 - LEAF_X + 0.3
            leaf_h = min(0.4, (branch_h - 0.05) / max(len(children), 1))
            lg = (branch_h - leaf_h * len(children)) / max(len(children), 1)
            for j, child in enumerate(children):
                ly = by + j * (leaf_h + lg)
                c.line(x1=BRANCH_X + BRANCH_W, y1=by + branch_h / 2,
                       x2=LEAF_X, y2=ly + leaf_h / 2, color="grey_400", width=0.5)
                c.box(x=LEAF_X, y=ly, w=LEAF_W, h=leaf_h,
                      fill="grey_100", border=0.5, border_color="grey_mid")
                c.text(child, x=LEAF_X + 0.1, y=ly, w=LEAF_W - 0.2, h=leaf_h,
                       size=8, color="grey_900", anchor="middle")

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)


# ============================================================
# Pattern 22 — Harvey Ball Matrix
# ============================================================


@dataclass
class HarveyBallSpec:
    header: SlideHeader
    intro: str
    row_labels: list[str]
    col_labels: list[str]
    scores: list[list[int]]   # 0~4
    highlight_row: int
    takeaway: str
    footer: SlideFooter


def harvey_ball_matrix(slide: Slide, spec: HarveyBallSpec):
    """Harvey Ball 점수 매트릭스 — 옵션/벤더 비교."""
    c = Canvas(slide)
    _draw_header(c, spec.header)
    if spec.intro:
        c.text(spec.intro, x=0.3, y=1.20, w=9.4, h=0.30,
               size=10, color="grey_900", anchor="top")

    n_rows = len(spec.row_labels)
    n_cols = len(spec.col_labels)
    LABEL_W, TABLE_X, TABLE_Y, TABLE_W = 2.2, 0.3, 1.7, 9.4
    HEADER_H = 0.55
    COL_W = (TABLE_W - LABEL_W) / max(n_cols, 1)
    avail_h = 6.55 - TABLE_Y - HEADER_H - 0.15
    ROW_H = avail_h / max(n_rows, 1)  # 가용 공간 균등 배분 (빈 공간 해소)

    hb_fills = {0: "white", 1: "grey_200", 2: "grey_400", 3: "grey_700", 4: "grey_900"}

    c.box(x=TABLE_X, y=TABLE_Y, w=LABEL_W, h=HEADER_H, fill="grey_800", border=None)
    for j, col in enumerate(spec.col_labels):
        cx = TABLE_X + LABEL_W + j * COL_W
        c.box(x=cx, y=TABLE_Y, w=COL_W, h=HEADER_H, fill="grey_800", border=None)
        c.text(col, x=cx, y=TABLE_Y, w=COL_W, h=HEADER_H,
               size=9, bold=True, color="white", align="center", anchor="middle")

    for i, label in enumerate(spec.row_labels):
        ry = TABLE_Y + HEADER_H + i * ROW_H
        is_hl = (i == spec.highlight_row)
        rf = "grey_200" if is_hl else ("white" if i % 2 == 0 else "grey_100")
        c.box(x=TABLE_X, y=ry, w=LABEL_W, h=ROW_H, fill=rf,
              border=0.5, border_color="grey_mid")
        c.text(label, x=TABLE_X + 0.1, y=ry, w=LABEL_W - 0.2, h=ROW_H,
               size=9, bold=is_hl, color="grey_900", anchor="middle")

        scores_row = spec.scores[i] if i < len(spec.scores) else []
        for j in range(n_cols):
            cx = TABLE_X + LABEL_W + j * COL_W
            c.box(x=cx, y=ry, w=COL_W, h=ROW_H, fill=rf,
                  border=0.5, border_color="grey_mid")
            score = scores_row[j] if j < len(scores_row) else 0
            bd = 0.3
            c.circle(x=cx + COL_W / 2 - bd / 2, y=ry + ROW_H / 2 - bd / 2,
                     d=bd, fill=hb_fills.get(score, "grey_400"),
                     border=1.0 if is_hl else 0.75,
                     border_color="grey_900" if is_hl else "grey_700",
                     text="", text_size=1)

    _draw_takeaway(c, spec.takeaway)
    _draw_footer(c, spec.footer)
