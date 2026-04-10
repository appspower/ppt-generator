"""Region-aware 컴포넌트 라이브러리 (20개).

각 컴포넌트는 (canvas, spec, region) → height_used 시그니처.
Region 안에서 상대좌표로 렌더하고, 사용한 높이를 반환한다.
스마트 컴포넌트는 데이터에 따라 색상/강조를 자동 결정한다.

사용 예:
    from ppt_builder.primitives import Canvas, Region
    from ppt_builder.components import comp_kpi_row

    region = Region(0.3, 1.7, 9.4, 2.0)
    h = comp_kpi_row(canvas, kpis=[...], region=region)
"""

from __future__ import annotations

from ppt_builder.primitives import Canvas, Region
from ppt_builder.assembler.styles import estimate_text_height, DesignToken

token = DesignToken()


# ============================================================
# 1. KPI Card (단일)
# ============================================================

def comp_kpi_card(
    c: Canvas,
    *,
    value: str,
    label: str,
    detail: str = "",
    trend: str = "flat",  # "up"/"down"/"flat"
    region: Region,
) -> float:
    """단일 KPI 카드 — region 전체를 사용."""
    r = region
    c.box(x=0, y=0, w=r.w, h=r.h,
          fill="white", border=0.75, border_color="grey_mid", region=r)

    # 좌측 stripe
    stripe_color = ("positive" if trend == "up"
                    else "negative" if trend == "down"
                    else "grey_700")
    c.box(x=0, y=0, w=0.08, h=r.h, fill=stripe_color, border=None, region=r)

    # 큰 숫자
    v_size = 26 if len(value) <= 4 else 20
    c.text(value, x=0.2, y=0.12, w=r.w - 0.3, h=r.h * 0.4,
           size=v_size, bold=True, color="grey_900", font="Georgia",
           anchor="top", region=r)

    # 라벨
    c.text(label, x=0.2, y=r.h * 0.45, w=r.w - 0.3, h=0.25,
           size=9, bold=True, color="grey_900", anchor="top", region=r)

    # 디테일
    if detail:
        c.text(detail, x=0.2, y=r.h * 0.45 + 0.27, w=r.w - 0.3, h=0.2,
               size=7, color="grey_700", anchor="top", region=r)

    return r.h


# ============================================================
# 2. KPI Row (N개 카드 가로 배열)
# ============================================================

def comp_kpi_row(
    c: Canvas,
    *,
    kpis: list[dict],
    region: Region,
    gap: float = 0.15,
) -> float:
    """KPI 카드 N개를 가로 배열."""
    n = len(kpis)
    if n == 0:
        return 0
    card_w = (region.w - gap * (n - 1)) / n
    for i, kpi in enumerate(kpis):
        card_r = region.sub(i * (card_w + gap), 0, card_w, region.h)
        comp_kpi_card(c, value=kpi["value"], label=kpi["label"],
                      detail=kpi.get("detail", ""), trend=kpi.get("trend", "flat"),
                      region=card_r)
    return region.h


# ============================================================
# 3. Mini Table
# ============================================================

def comp_mini_table(
    c: Canvas,
    *,
    headers: list[str],
    rows: list[list[str]],
    region: Region,
    col_ratios: list[float] | None = None,
) -> float:
    """컴팩트 테이블 — region 내 렌더 (push_region 활용)."""
    c.push_region(region)
    try:
        c.mini_table(
            x=0, y=0, w=region.w, h=region.h,
            headers=headers, rows=rows, col_ratios=col_ratios,
        )
    finally:
        c.pop_region()
    return region.h


# ============================================================
# 4. Bullet List
# ============================================================

def comp_bullet_list(
    c: Canvas,
    *,
    title: str = "",
    items: list[str],
    region: Region,
    title_size: float = 11,
    item_size: float = 9,
) -> float:
    """불릿 목록 — 제목(선택) + 항목들."""
    r = region
    cy = 0.0

    if title:
        c.text(title, x=0.1, y=cy, w=r.w - 0.2, h=0.3,
               size=title_size, bold=True, color="grey_900", anchor="top", region=r)
        cy += 0.35

    for item in items:
        item_h = max(0.22, estimate_text_height(
            f"▪  {item}", font_pt=item_size, box_width_inches=r.w - 0.3))
        c.text(f"▪  {item}", x=0.1, y=cy, w=r.w - 0.2, h=item_h,
               size=item_size, color="grey_900", anchor="top", region=r)
        cy += item_h + 0.04

    return cy


# ============================================================
# 5. Bar Chart (수평)
# ============================================================

def comp_bar_chart_h(
    c: Canvas,
    *,
    title: str = "",
    data: list[dict],  # [{"label": "...", "value": float, "highlight": False}]
    unit: str = "",
    region: Region,
) -> float:
    """수평 바 차트 — region 내 렌더."""
    r = region
    cy = 0.0

    if title:
        c.text(title, x=0.1, y=cy, w=r.w - 0.2, h=0.3,
               size=11, bold=True, color="grey_900", anchor="top", region=r)
        cy += 0.35

    n = len(data)
    max_val = max((d["value"] for d in data), default=1)
    bar_h = min(0.35, (r.h - cy - 0.1) / max(n, 1))
    bar_gap = 0.06
    label_w = 1.2
    bar_max_w = r.w - label_w - 1.0

    for i, d in enumerate(data):
        by = cy + i * (bar_h + bar_gap)
        is_hl = d.get("highlight", False)

        c.text(d["label"], x=0.05, y=by, w=label_w, h=bar_h,
               size=8, bold=True, color="grey_900", anchor="middle", region=r)

        bw = bar_max_w * (d["value"] / max_val) if max_val > 0 else 0
        c.box(x=label_w + 0.1, y=by + 0.04, w=max(bw, 0.05), h=bar_h - 0.08,
              fill="grey_900" if is_hl else "grey_400", border=None, region=r)

        val_str = f"{d['value']:,.0f}{unit}" if isinstance(d["value"], (int, float)) else str(d["value"])
        c.text(val_str, x=label_w + 0.15 + bw, y=by, w=0.8, h=bar_h,
               size=9, bold=is_hl, color="grey_900", anchor="middle", region=r)

    return cy + n * (bar_h + bar_gap)


# ============================================================
# 6. Stat Row (수평 통계 블록 배열)
# ============================================================

def comp_stat_row(
    c: Canvas,
    *,
    stats: list[dict],  # [{"value": "14%", "label": "일정 단축"}]
    region: Region,
    gap: float = 0.15,
) -> float:
    """stat_block N개를 가로 배열."""
    n = len(stats)
    if n == 0:
        return 0
    sw = (region.w - gap * (n - 1)) / n
    for i, st in enumerate(stats):
        sr = region.sub(i * (sw + gap), 0, sw, region.h)
        c.stat_block(value=st["value"], label=st["label"],
                     x=0, y=0, w=sr.w, h=sr.h, region=sr)
    return region.h


# ============================================================
# 7. Callout Block (강조 박스)
# ============================================================

def comp_callout(
    c: Canvas,
    *,
    title: str = "",
    body: str = "",
    bullets: list[str] | None = None,
    bar_color: str = "grey_700",
    region: Region,
) -> float:
    """callout_box — region 내 렌더 (push_region 활용)."""
    c.push_region(region)
    try:
        c.callout_box(
            x=0, y=0, w=region.w, h=region.h,
            title=title, body=body, bullets=bullets,
            bar_color=bar_color,
        )
    finally:
        c.pop_region()
    return region.h


# ============================================================
# 8. RAG Row (상태 행 하나)
# ============================================================

def comp_rag_row(
    c: Canvas,
    *,
    label: str,
    values: list[str],  # ["G", "A", "R", "-"]
    region: Region,
    label_w: float = 2.0,
) -> float:
    """RAG 상태 행 하나 — 라벨 + 색상 원들."""
    r = region
    rag_colors = {"G": "positive", "A": "#F5A623", "R": "negative", "-": "grey_400"}

    c.box(x=0, y=0, w=label_w, h=r.h,
          fill="grey_100", border=0.5, border_color="grey_mid", region=r)
    c.text(label, x=0.1, y=0, w=label_w - 0.2, h=r.h,
           size=9, bold=True, color="grey_900", anchor="middle", region=r)

    n = len(values)
    col_w = (r.w - label_w) / max(n, 1)
    for j, val in enumerate(values):
        cx = label_w + j * col_w
        c.box(x=cx, y=0, w=col_w, h=r.h,
              fill="white", border=0.5, border_color="grey_mid", region=r)
        d = 0.28
        c.circle(x=cx + col_w / 2 - d / 2, y=r.h / 2 - d / 2, d=d,
                 fill=rag_colors.get(val, "grey_400"),
                 border=None, text="", text_size=1, region=r)

    return r.h


# ============================================================
# 9. Numbered Items (번호 목록)
# ============================================================

def comp_numbered_items(
    c: Canvas,
    *,
    items: list[tuple[str, str]],  # [(title, detail), ...]
    region: Region,
    item_h: float = 0.6,
) -> float:
    """번호 원 + 제목 + 디테일 목록 (push_region 활용)."""
    c.push_region(region)
    try:
        c.numbered_list(
            items=items,
            x=0, y=0, w=region.w, item_h=item_h,
        )
    finally:
        c.pop_region()
    return len(items) * (item_h + 0.12)


# ============================================================
# 10. Section Header (섹션 구분)
# ============================================================

def comp_section_header(
    c: Canvas,
    *,
    title: str,
    region: Region,
) -> float:
    """섹션 헤더 — 좌측 바 + 굵은 제목."""
    c.section_label(title, x=0, y=0, w=region.w, region=region)
    return 0.3


# ============================================================
# 11. Progress Bar (스마트 — 값에 따라 색상 자동)
# ============================================================

def comp_progress_bar(
    c: Canvas,
    *,
    label: str,
    value: float,       # 0~100
    target: float = 80,  # 목표값
    region: Region,
) -> float:
    """진행률 바 — 값에 따라 색상 자동 (≥target 녹, 50~target 주황, <50 빨강)."""
    r = region
    bar_color = token.auto_trend_color(value, target, higher_is_better=True)
    bar_h = min(0.3, r.h * 0.4)
    label_h = 0.22

    c.text(label, x=0.05, y=0, w=r.w * 0.4, h=label_h,
           size=9, bold=True, color=token.TEXT_PRIMARY, anchor="middle", region=r)

    bar_x = r.w * 0.42
    bar_w = r.w * 0.45
    bar_y = (r.h - bar_h) / 2

    # 배경 트랙
    c.box(x=bar_x, y=bar_y, w=bar_w, h=bar_h,
          fill=token.MUTED, border=None, region=r)
    # 채워진 바
    fill_w = bar_w * min(value / 100, 1.0)
    c.box(x=bar_x, y=bar_y, w=fill_w, h=bar_h,
          fill=bar_color, border=None, region=r)
    # 값 텍스트
    c.text(f"{value:.0f}%", x=bar_x + bar_w + 0.08, y=0,
           w=0.6, h=r.h, size=10, bold=True, color=bar_color,
           anchor="middle", region=r)

    return r.h


# ============================================================
# 12. Vertical Bar Chart (스마트 — 최대값 자동 강조)
# ============================================================

def comp_vertical_bars(
    c: Canvas,
    *,
    data: list[dict],   # [{"label": "...", "value": float}]
    unit: str = "",
    region: Region,
) -> float:
    """세로 바 차트 — 최대값 자동 강조(accent), 나머지 neutral."""
    r = region
    n = len(data)
    if n == 0:
        return 0

    max_val = max(d["value"] for d in data)
    gap = 0.08
    bar_w = (r.w - gap * (n + 1)) / n
    chart_h = r.h * 0.65
    label_area = r.h * 0.2

    for i, d in enumerate(data):
        bx = gap + i * (bar_w + gap)
        ratio = d["value"] / max_val if max_val > 0 else 0
        bh = chart_h * ratio
        by = r.h - label_area - bh

        is_max = (d["value"] == max_val)
        fill = token.DATA_PRIMARY if is_max else token.DATA_SECONDARY
        c.box(x=bx, y=by, w=bar_w, h=bh, fill=fill, border=None, region=r)

        # 값 (바 위)
        fmt = ",.1f" if isinstance(d["value"], float) and d["value"] != int(d["value"]) else ",.0f"
        val_str = f"{d['value']:{fmt}}{unit}"
        c.text(val_str, x=bx, y=by - 0.22, w=bar_w, h=0.2,
               size=8, bold=is_max, color=token.TEXT_PRIMARY,
               align="center", anchor="bottom", region=r)

        # 라벨 (하단)
        c.text(d["label"], x=bx, y=r.h - label_area, w=bar_w, h=label_area,
               size=7, color=token.TEXT_SECONDARY, align="center", anchor="top", region=r)

    return r.h


# ============================================================
# 13. Heat Row (스마트 — 값에 따라 셀 배경 농도 자동)
# ============================================================

def comp_heat_row(
    c: Canvas,
    *,
    label: str,
    values: list[float],
    col_labels: list[str] | None = None,
    max_val: float = 100,
    region: Region,
) -> float:
    """히트맵 행 — 값이 클수록 진한 배경."""
    r = region
    label_w = r.w * 0.25
    n = len(values)
    cell_w = (r.w - label_w) / max(n, 1)

    c.box(x=0, y=0, w=label_w, h=r.h,
          fill=token.ZONE_SUBTLE, border=0.5, border_color=token.BORDER, region=r)
    c.text(label, x=0.08, y=0, w=label_w - 0.16, h=r.h,
           size=9, bold=True, color=token.TEXT_PRIMARY, anchor="middle", region=r)

    heat_colors = ["grey_100", "grey_200", "grey_400", "grey_700", "grey_900"]
    heat_txt = [token.TEXT_PRIMARY, token.TEXT_PRIMARY, token.TEXT_PRIMARY, "white", "white"]

    for j, val in enumerate(values):
        cx = label_w + j * cell_w
        # 0~4 단계로 매핑
        level = min(4, int((val / max_val) * 5)) if max_val > 0 else 0
        c.box(x=cx, y=0, w=cell_w, h=r.h,
              fill=heat_colors[level], border=0.5, border_color=token.BORDER, region=r)
        fmt = ",.0f" if val == int(val) else ",.1f"
        c.text(f"{val:{fmt}}", x=cx, y=0, w=cell_w, h=r.h,
               size=9, bold=(level >= 3), color=heat_txt[level],
               align="center", anchor="middle", region=r)

    return r.h


# ============================================================
# 14. Gauge (원형 달성률 — 간소화 버전)
# ============================================================

def comp_gauge(
    c: Canvas,
    *,
    value: float,       # 0~100
    label: str,
    target: float = 80,
    region: Region,
) -> float:
    """원형 게이지 (간소화) — 큰 숫자 + 배경 원 + 전경 원."""
    r = region
    d = min(r.w, r.h) * 0.6
    cx = (r.w - d) / 2
    cy = 0.05

    gauge_color = token.auto_trend_color(value, target)

    # 배경 원
    c.circle(x=cx, y=cy, d=d, fill=token.MUTED,
             border=None, text="", text_size=1, region=r)
    # 전경 (작은 원으로 오버레이 — 달성률 표현)
    inner_d = d * 0.7
    c.circle(x=cx + (d - inner_d) / 2, y=cy + (d - inner_d) / 2,
             d=inner_d, fill="white",
             border=2.0, border_color=gauge_color,
             text=f"{value:.0f}%", text_color=gauge_color,
             text_size=16, text_bold=True, region=r)

    # 라벨
    c.text(label, x=0, y=cy + d + 0.08, w=r.w, h=0.25,
           size=9, bold=True, color=token.TEXT_PRIMARY,
           align="center", anchor="top", region=r)

    return r.h


# ============================================================
# 15. Tag Group (태그/뱃지 묶음 — 자동 줄바꿈)
# ============================================================

def comp_tag_group(
    c: Canvas,
    *,
    tags: list[str],
    region: Region,
    fill: str = "grey_200",
    text_color: str = "grey_900",
) -> float:
    """태그/뱃지 묶음 — 자동 가로 배열, 줄바꿈."""
    r = region
    cx, cy = 0.0, 0.0
    tag_h = 0.26
    gap_x, gap_y = 0.08, 0.06

    for tag in tags:
        tw = c.badge(tag, x=cx, y=cy, fill=fill, text_color=text_color,
                     size=8, region=r)
        cx += tw + gap_x
        if cx + 0.5 > r.w:  # 줄바꿈
            cx = 0.0
            cy += tag_h + gap_y

    return cy + tag_h


# ============================================================
# 16. Comparison Row (A vs B — 더 큰 쪽 자동 강조)
# ============================================================

def comp_comparison_row(
    c: Canvas,
    *,
    label: str,
    value_a: str,
    value_b: str,
    num_a: float = 0,
    num_b: float = 0,
    region: Region,
) -> float:
    """비교 행 — 값이 큰 쪽 자동 강조."""
    r = region
    third = r.w / 3

    c.text(label, x=0, y=0, w=third, h=r.h,
           size=9, bold=True, color=token.TEXT_PRIMARY, anchor="middle", region=r)

    a_bold = num_a >= num_b
    b_bold = num_b > num_a
    a_color = token.DATA_HIGHLIGHT if a_bold else token.TEXT_SECONDARY
    b_color = token.DATA_HIGHLIGHT if b_bold else token.TEXT_SECONDARY

    c.text(value_a, x=third, y=0, w=third, h=r.h,
           size=10, bold=a_bold, color=a_color,
           align="center", anchor="middle", region=r)
    c.text(value_b, x=third * 2, y=0, w=third, h=r.h,
           size=10, bold=b_bold, color=b_color,
           align="center", anchor="middle", region=r)

    return r.h


# ============================================================
# 17. Metric Delta (변화량 — 자동 색상/화살표)
# ============================================================

def comp_metric_delta(
    c: Canvas,
    *,
    label: str,
    current: float,
    previous: float,
    unit: str = "",
    higher_is_better: bool = True,
    region: Region,
) -> float:
    """변화량 표시 — 트렌드 자동 판단, 색상 자동."""
    r = region
    delta = current - previous
    symbol = token.auto_trend_symbol(current, previous)
    delta_str = token.auto_delta(current, previous, unit)

    is_good = (delta > 0 and higher_is_better) or (delta < 0 and not higher_is_better)
    color = "positive" if is_good else ("negative" if delta != 0 else token.NEUTRAL)

    c.text(label, x=0.05, y=0, w=r.w * 0.35, h=r.h,
           size=9, bold=True, color=token.TEXT_PRIMARY, anchor="middle", region=r)

    fmt = ",.1f" if isinstance(current, float) and current != int(current) else ",.0f"
    c.text(f"{current:{fmt}}{unit}", x=r.w * 0.37, y=0, w=r.w * 0.25, h=r.h,
           size=12, bold=True, color=token.TEXT_PRIMARY, anchor="middle", region=r)

    c.text(f"{symbol} {delta_str}", x=r.w * 0.65, y=0, w=r.w * 0.33, h=r.h,
           size=10, bold=True, color=color, anchor="middle", region=r)

    return r.h


# ============================================================
# 18. Timeline Mini (미니 타임라인 — 현재 위치 마커)
# ============================================================

def comp_timeline_mini(
    c: Canvas,
    *,
    phases: list[str],
    current: int,   # 0-based
    region: Region,
) -> float:
    """미니 수평 타임라인 — 현재 단계 강조."""
    r = region
    n = len(phases)
    line_y = r.h * 0.4
    dot_d = 0.16
    gap = r.w / max(n - 1, 1) if n > 1 else r.w

    # 수평선
    c.line(x1=0.2, y1=line_y, x2=r.w - 0.2, y2=line_y,
           color=token.NEUTRAL, width=1.5, region=r)

    for i, phase in enumerate(phases):
        px = 0.2 + i * gap if n > 1 else r.w / 2
        is_current = (i == current)
        is_past = (i < current)

        fill = token.DATA_HIGHLIGHT if is_current else (
            token.NEUTRAL if is_past else token.MUTED)
        dd = dot_d * 1.3 if is_current else dot_d
        c.circle(x=px - dd / 2, y=line_y - dd / 2, d=dd,
                 fill=fill, border=None, text="", text_size=1, region=r)

        c.text(phase, x=px - 0.5, y=line_y + 0.15, w=1.0, h=0.25,
               size=7, bold=is_current, color=token.TEXT_PRIMARY if is_current else token.TEXT_SECONDARY,
               align="center", anchor="top", region=r)

    return r.h


# ============================================================
# 19. Icon List (아이콘+텍스트 — 번호/체크/경고 자동)
# ============================================================

def comp_icon_list(
    c: Canvas,
    *,
    items: list[dict],  # [{"text": "...", "icon": "check"/"warn"/"num"/"bullet"}]
    region: Region,
) -> float:
    """아이콘+텍스트 목록. icon 타입에 따라 자동 스타일."""
    r = region
    n = len(items)
    item_h = min(0.4, (r.h - 0.05) / max(n, 1))
    icons = {
        "check": ("✓", "positive"),
        "warn": ("!", "warning"),
        "error": ("✗", "negative"),
        "bullet": ("▪", token.TEXT_SECONDARY),
    }

    for i, item in enumerate(items):
        iy = i * item_h
        icon_type = item.get("icon", "bullet")

        if icon_type == "num":
            symbol = f"{i + 1:02d}"
            sym_color = "white"
            c.circle(x=0.02, y=iy + 0.04, d=0.28,
                     fill=token.DATA_HIGHLIGHT, border=None,
                     text=symbol, text_color=sym_color, text_size=9, region=r)
        else:
            symbol, sym_color = icons.get(icon_type, ("▪", token.TEXT_SECONDARY))
            c.text(symbol, x=0.05, y=iy, w=0.25, h=item_h,
                   size=12, bold=True, color=sym_color, anchor="middle", region=r)

        c.text(item["text"], x=0.38, y=iy, w=r.w - 0.45, h=item_h,
               size=9, color=token.TEXT_PRIMARY, anchor="middle", region=r)

    return n * item_h


# ============================================================
# 20. Data Card (스마트 KPI — 자동 트렌드/색상/델타)
# ============================================================

def comp_data_card(
    c: Canvas,
    *,
    value: float,
    label: str,
    previous: float | None = None,
    target: float | None = None,
    unit: str = "",
    higher_is_better: bool = True,
    detail: str = "",
    region: Region,
) -> float:
    """스마트 데이터 카드 — 모든 것 자동 판단."""
    r = region

    # 자동 색상 결정
    if target is not None:
        stripe_color = token.auto_trend_color(value, target, higher_is_better)
    elif previous is not None:
        delta = value - previous
        is_good = (delta > 0 and higher_is_better) or (delta < 0 and not higher_is_better)
        stripe_color = "positive" if is_good else "negative"
    else:
        stripe_color = token.STRIPE

    # 카드 배경
    c.box(x=0, y=0, w=r.w, h=r.h,
          fill=token.ZONE_LIGHT, border=0.75, border_color=token.BORDER, region=r)
    # 좌측 stripe
    c.box(x=0, y=0, w=0.08, h=r.h, fill=stripe_color, border=None, region=r)

    # 큰 숫자
    fmt = ",.1f" if isinstance(value, float) and value != int(value) else ",.0f"
    val_str = f"{value:{fmt}}{unit}"
    v_size = 22 if len(val_str) <= 5 else 16
    c.text(val_str, x=0.2, y=0.1, w=r.w - 0.3, h=r.h * 0.35,
           size=v_size, bold=True, color=token.TEXT_PRIMARY,
           font="Georgia", anchor="top", region=r)

    # 라벨
    c.text(label, x=0.2, y=r.h * 0.4, w=r.w - 0.3, h=0.22,
           size=9, bold=True, color=token.TEXT_PRIMARY, anchor="top", region=r)

    # 트렌드 (previous 있으면)
    if previous is not None:
        symbol = token.auto_trend_symbol(value, previous)
        delta_str = token.auto_delta(value, previous, unit)
        c.text(f"{symbol} {delta_str}", x=0.2, y=r.h * 0.4 + 0.24,
               w=r.w - 0.3, h=0.2, size=8, bold=True,
               color=stripe_color, anchor="top", region=r)

    # 디테일
    if detail:
        c.text(detail, x=0.2, y=r.h - 0.28, w=r.w - 0.3, h=0.2,
               size=7, color=token.TEXT_SECONDARY, anchor="top", region=r)

    return r.h
