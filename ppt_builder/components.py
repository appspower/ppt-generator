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
    """단일 KPI 카드 — 콤팩트 배치, 최대 1.5" 높이."""
    r = region
    # 카드 높이를 콘텐츠에 맞게 제한 (최대 1.5")
    card_h = min(r.h, 1.5)
    # 세로 중앙 정렬
    y_offset = (r.h - card_h) / 2

    c.box(x=0, y=y_offset, w=r.w, h=card_h,
          fill="white", border=0.75, border_color="grey_mid", region=r)

    # 좌측 stripe
    stripe_color = ("positive" if trend == "up"
                    else "negative" if trend == "down"
                    else "grey_700")
    c.box(x=0, y=y_offset, w=0.08, h=card_h,
          fill=stripe_color, border=None, region=r)

    # 고정 위치 배치 (비율 대신 절대값)
    v_size = 26 if len(value) <= 4 else 20
    c.text(value, x=0.2, y=y_offset + 0.12, w=r.w - 0.3, h=0.55,
           size=v_size, bold=True, color="grey_900", font="Georgia",
           anchor="top", region=r)

    c.text(label, x=0.2, y=y_offset + 0.70, w=r.w - 0.3, h=0.30,
           size=9, bold=True, color="grey_900", anchor="top", region=r)

    if detail:
        c.text(detail, x=0.2, y=y_offset + 1.02, w=r.w - 0.3, h=0.25,
               size=7, color="grey_700", anchor="top", region=r)

    return card_h


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
    """불릿 목록 — 제목(선택) + 항목들. 밀도 우선 배치."""
    r = region
    cy = 0.0

    if title:
        c.text(title, x=0.1, y=cy, w=r.w - 0.2, h=0.22,
               size=title_size, bold=True, color="grey_900", anchor="top", region=r)
        cy += 0.22

    for item in items:
        item_h = max(0.18, estimate_text_height(
            f"▪  {item}", font_pt=item_size, box_width_inches=r.w - 0.3))
        c.text(f"▪  {item}", x=0.1, y=cy, w=r.w - 0.2, h=item_h,
               size=item_size, color="grey_900", anchor="top", region=r)
        cy += item_h + 0.02

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


# ============================================================
# Track B: 아이콘 시스템
# ============================================================

# PwC 스타일 아이콘 (유니코드 기반 — 별도 이미지 불필요)
ICONS = {
    # 비즈니스 개념
    "database": "🗄",  "cloud": "☁",  "globe": "🌐",  "lock": "🔒",
    "chart": "📊",  "target": "🎯",  "rocket": "🚀",  "gear": "⚙",
    "people": "👥",  "money": "💰",  "clock": "⏱",  "check_circle": "✅",
    "warning": "⚠",  "lightning": "⚡",  "star": "★",  "flag": "🚩",
    # 화살표/방향
    "arrow_right": "→",  "arrow_up": "↑",  "arrow_down": "↓",
    "chevron_right": "›",  "bullet": "▪",  "circle": "●",
    # 상태
    "check": "✓",  "cross": "✗",  "dash": "—",
    "triangle_up": "▲",  "triangle_down": "▼",
}


def comp_icon_card(
    c: Canvas,
    *,
    icon: str,
    title: str,
    body: str = "",
    style: str = "light",  # "light"/"dark"/"accent"/"outline"
    region: Region,
) -> float:
    """아이콘 + 제목 + 본문 카드. 스타일 5종 지원."""
    r = region
    styles = {
        "light":   {"bg": "white", "border": 0.75, "bc": "grey_mid",
                    "tc": token.TEXT_PRIMARY, "ic": token.TEXT_PRIMARY},
        "dark":    {"bg": "grey_800", "border": None, "bc": None,
                    "tc": "white", "ic": "white"},
        "accent":  {"bg": "accent", "border": None, "bc": None,
                    "tc": "white", "ic": "white"},
        "subtle":  {"bg": "grey_100", "border": 0.5, "bc": "grey_mid",
                    "tc": token.TEXT_PRIMARY, "ic": token.TEXT_SECONDARY},
        "outline": {"bg": "white", "border": 1.5, "bc": "grey_800",
                    "tc": token.TEXT_PRIMARY, "ic": "accent"},
    }
    s = styles.get(style, styles["light"])

    c.box(x=0, y=0, w=r.w, h=r.h,
          fill=s["bg"], border=s["border"], border_color=s["bc"] or "grey_mid", region=r)

    # 아이콘
    icon_char = ICONS.get(icon, icon)
    c.text(icon_char, x=0.15, y=0.12, w=0.5, h=0.4,
           size=20, color=s["ic"], anchor="top", region=r)

    # 제목
    c.text(title, x=0.15, y=0.55, w=r.w - 0.3, h=0.3,
           size=10, bold=True, color=s["tc"], anchor="top", region=r)

    # 본문
    if body:
        c.text(body, x=0.15, y=0.85, w=r.w - 0.3, h=r.h - 0.95,
               size=8, color=s["tc"], anchor="top", region=r)

    return r.h


def comp_icon_row(
    c: Canvas,
    *,
    items: list[dict],  # [{"icon": "...", "title": "...", "body": "...", "style": "..."}]
    region: Region,
    gap: float = 0.12,
) -> float:
    """아이콘 카드 N개 가로 배열."""
    r = region
    n = len(items)
    if n == 0:
        return 0
    card_w = (r.w - gap * (n - 1)) / n
    for i, item in enumerate(items):
        cr = r.sub(i * (card_w + gap), 0, card_w, r.h)
        comp_icon_card(c, icon=item.get("icon", "circle"),
                       title=item["title"], body=item.get("body", ""),
                       style=item.get("style", "light"), region=cr)
    return r.h


# ============================================================
# Track C: 스타일 카드 변형 (5종)
# ============================================================

def comp_styled_card(
    c: Canvas,
    *,
    title: str,
    body: str = "",
    bullets: list[str] | None = None,
    number: str = "",       # "01", "02" 등 큰 번호
    kpi_value: str = "",    # 큰 숫자
    style: str = "light",   # light/dark/accent/subtle/numbered
    region: Region,
) -> float:
    """다양한 스타일의 콘텐츠 카드.

    light: 흰 배경 + 테두리
    dark: 진회색 배경 + 흰 글씨
    accent: 오렌지 배경 + 흰 글씨
    subtle: 연회색 배경 + 회색 글씨
    numbered: 좌상단 큰 번호 + 내용
    """
    r = region
    card_styles = {
        "light":    {"bg": "white", "border": 0.75, "bc": "grey_mid",
                     "tc": token.TEXT_PRIMARY, "sc": token.TEXT_SECONDARY},
        "dark":     {"bg": "grey_800", "border": None, "bc": None,
                     "tc": "white", "sc": "grey_200"},
        "accent":   {"bg": "accent", "border": None, "bc": None,
                     "tc": "white", "sc": "white"},
        "subtle":   {"bg": "grey_100", "border": 0.5, "bc": "grey_mid",
                     "tc": token.TEXT_PRIMARY, "sc": token.TEXT_SECONDARY},
        "numbered": {"bg": "white", "border": 0.75, "bc": "grey_mid",
                     "tc": token.TEXT_PRIMARY, "sc": token.TEXT_SECONDARY},
    }
    s = card_styles.get(style, card_styles["light"])

    c.box(x=0, y=0, w=r.w, h=r.h,
          fill=s["bg"], border=s["border"], border_color=s["bc"] or "grey_mid", region=r)

    cy = 0.08

    # 큰 번호 (numbered 스타일)
    if number and style == "numbered":
        c.text(number, x=0.12, y=cy, w=0.6, h=0.38,
               size=28, bold=True, color="accent", anchor="top", region=r)
        cy += 0.35

    # KPI 큰 숫자
    if kpi_value:
        c.text(kpi_value, x=0.12, y=cy, w=r.w - 0.24, h=0.35,
               size=24, bold=True, color=s["tc"], font="Georgia", anchor="top", region=r)
        cy += 0.35

    # 제목
    c.text(title, x=0.12, y=cy, w=r.w - 0.24, h=0.22,
           size=10, bold=True, color=s["tc"], anchor="top", region=r)
    cy += 0.24

    # 본문
    if body:
        c.text(body, x=0.12, y=cy, w=r.w - 0.24, h=r.h - cy - 0.08,
               size=8, color=s["sc"], anchor="top", region=r)

    # 불릿 — 밀도 우선 배치
    if bullets:
        for bul in bullets:
            bul_h = max(0.18, estimate_text_height(
                f"▪  {bul}", font_pt=8, box_width_inches=r.w - 0.36))
            c.text(f"▪  {bul}", x=0.12, y=cy, w=r.w - 0.24, h=bul_h,
                   size=8, color=s["sc"], anchor="top", region=r)
            cy += bul_h + 0.02

    return r.h


def comp_styled_card_row(
    c: Canvas,
    *,
    cards: list[dict],
    region: Region,
    gap: float = 0.12,
) -> float:
    """스타일 카드 N개 가로 배열."""
    r = region
    n = len(cards)
    if n == 0:
        return 0
    card_w = (r.w - gap * (n - 1)) / n
    for i, card in enumerate(cards):
        cr = r.sub(i * (card_w + gap), 0, card_w, r.h)
        comp_styled_card(c, region=cr, **card)
    return r.h


# ============================================================
# Track A 연동: 네이티브 차트 컴포넌트 래퍼
# ============================================================

def comp_native_chart(
    c: Canvas,
    *,
    chart_type: str,  # "vertical_bar"/"line"/"donut"/"stacked_bar"/"scatter"
    chart_kwargs: dict,
    region: Region,
) -> float:
    """네이티브 PPT 차트를 Region 안에 배치.

    chart_kwargs는 charts.native의 각 함수에 전달되는 인자.
    slide는 canvas에서 가져옴.
    """
    from ppt_builder.charts.native import (
        chart_vertical_bar, chart_line, chart_donut,
        chart_stacked_bar, chart_scatter,
    )
    funcs = {
        "vertical_bar": chart_vertical_bar,
        "line": chart_line,
        "donut": chart_donut,
        "stacked_bar": chart_stacked_bar,
        "scatter": chart_scatter,
    }
    func = funcs.get(chart_type)
    if func is None:
        raise ValueError(f"Unknown chart type: {chart_type}")

    func(c.slide, region=region, **chart_kwargs)
    return region.h


# ============================================================
# Blueprint Phase 1 — 앵커/복합 컴포넌트
# ============================================================


def comp_numbered_cell(
    c: Canvas,
    *,
    number: str,
    header: str,
    body: str = "",
    bg_color: str = "white",
    number_size: int = 36,
    region: Region,
) -> float:
    """번호+색상 코딩된 그리드 셀.

    CP2 대응 (PwC B00/B01 재현).
    전체 region을 bg_color로 채우고, 좌상단에 큰 번호, 아래에 헤더+본문.
    """
    from ppt_builder.primitives import color as resolve_color
    r = region

    # 배경
    c.box(x=0, y=0, w=r.w, h=r.h,
          fill=bg_color, border=None, region=r)

    # 텍스트 색상: 어두운 배경이면 흰색
    dark_bgs = {"accent", "accent_mid", "dark", "grey_800", "grey_900", "grey_700"}
    txt_color = "white" if bg_color in dark_bgs else "grey_900"
    num_color = "white" if bg_color in dark_bgs else "grey_800"

    # 번호 (좌상단, 크게)
    c.text(number, x=0.15, y=0.12, w=r.w - 0.3, h=r.h * 0.40,
           size=number_size, bold=True, color=num_color,
           font="Georgia", anchor="top", region=r)

    # 헤더
    hdr_y = r.h * 0.48
    c.text(header, x=0.15, y=hdr_y, w=r.w - 0.3, h=0.28,
           size=11, bold=True, color=txt_color, anchor="top", region=r)

    # 본문
    if body:
        c.text(body, x=0.15, y=hdr_y + 0.30, w=r.w - 0.3, h=r.h - hdr_y - 0.40,
               size=8, color=txt_color, anchor="top", region=r)

    return r.h


def comp_timeline_marker(
    c: Canvas,
    *,
    labels: list[str],
    style: str = "arrow",  # "arrow" | "dots" | "bar"
    highlight_idx: int = -1,
    region: Region,
) -> float:
    """타임라인 밴드 마커.

    CP3 대응 (PwC B08/B10 재현).
    Region 내에 수평 밴드를 그리고, 라벨을 등간격으로 배치.
    """
    r = region
    n = len(labels)
    if n == 0:
        return r.h

    # 그라데이션 밴드 (왼쪽 연한색 → 오른쪽 진한색)
    seg_w = r.w / n
    grad_colors = ["grey_200", "grey_400", "grey_700", "accent_mid", "accent"]

    for i in range(n):
        ci = min(i, len(grad_colors) - 1)
        gc = grad_colors[ci] if n <= len(grad_colors) else (
            grad_colors[int(i / n * (len(grad_colors) - 1))]
        )
        if i == highlight_idx:
            gc = "accent"
        c.box(x=i * seg_w, y=0, w=seg_w + 0.01, h=r.h,
              fill=gc, border=None, region=r)

    # 라벨
    for i, lbl in enumerate(labels):
        txt_clr = "white" if i >= n // 2 else "grey_900"
        c.text(lbl, x=i * seg_w, y=0, w=seg_w, h=r.h,
               size=10, bold=True, color=txt_clr, anchor="middle", region=r)

    if style == "dots":
        # 각 세그먼트 경계에 마커 도트
        for i in range(1, n):
            dot_x = i * seg_w - 0.06
            c.box(x=dot_x, y=r.h / 2 - 0.06, w=0.12, h=0.12,
                  fill="accent", border=None, region=r)

    return r.h


def comp_icon_header_card(
    c: Canvas,
    *,
    icon: str,
    header: str,
    body: str,
    icon_size: float = 0.45,
    region: Region,
) -> float:
    """아이콘 + 헤더 + 본문 카드.

    CP6 대응 (PwC B07/B09 재현).
    상단에 아이콘, 그 아래에 헤더+본문.
    """
    r = region

    # 아이콘 (센터 정렬)
    icon_x = (r.w - icon_size) / 2
    c.icon(name=icon, x=icon_x, y=0.05, size=icon_size,
           color="accent", region=r)

    # 헤더
    hdr_y = icon_size + 0.15
    c.text(header, x=0.05, y=hdr_y, w=r.w - 0.1, h=0.25,
           size=10, bold=True, color="grey_900", anchor="top", region=r)

    # 본문
    body_y = hdr_y + 0.28
    c.text(body, x=0.05, y=body_y, w=r.w - 0.1, h=r.h - body_y - 0.05,
           size=8, color="grey_700", anchor="top", region=r)

    return r.h


# ============================================================
# Compound Component 1: Chevron Flow
# ============================================================

def comp_chevron_flow(
    c: Canvas,
    *,
    phases: list[dict],       # [{"label": "분석", "tag": "01"}, ...]
    style: str = "gradient",  # "gradient" | "uniform" | "accent_last"
    show_details: bool = False,
    region: Region,
) -> float:
    """수평 쉐브론 화살표 체인 (3~6단계).

    timeline_phases, executive_summary, value_chain, chevron_timeline에서 추출.
    4개 패턴에서 중복 사용된 최고 빈도 시각 요소.

    사용 맥락: 프로세스 로드맵, 프로젝트 페이즈, 의사결정 흐름
    """
    r = region
    n = len(phases)
    if n == 0:
        return 0.0

    # 쉐브론 영역 (show_details면 상단 40%, 아니면 전체)
    chev_h = min(r.h * 0.40, 0.55) if show_details else min(r.h, 0.55)
    overlap = 0.08
    chev_w = (r.w + overlap * (n - 1)) / n

    # 색상 팔레트
    if style == "gradient":
        fills = ["grey_800", "grey_700", "grey_400", "grey_200", "grey_100", "white"]
        txt_c = ["white", "white", "white", "grey_900", "grey_900", "grey_900"]
    elif style == "accent_last":
        fills = ["grey_200"] * (n - 1) + ["accent"]
        txt_c = ["grey_900"] * (n - 1) + ["white"]
    else:  # uniform
        fills = ["grey_700"] * n
        txt_c = ["white"] * n

    for i, p in enumerate(phases):
        cx = i * (chev_w - overlap)
        fi = fills[min(i, len(fills) - 1)]
        tc = txt_c[min(i, len(txt_c) - 1)]
        tag = p.get("tag", f"{i + 1:02d}")
        label = p.get("label", "")
        text = f"{tag}  {label}" if tag else label
        c.chevron(x=cx, y=0, w=chev_w, h=chev_h,
                  fill=fi, text=text, text_color=tc, text_size=9,
                  region=r)

    # 상세 카드 (선택) — 높이를 콘텐츠에 맞게 제한
    if show_details and n > 0:
        detail_y = chev_h + 0.08
        # 카드 높이: 항목 수 기반 동적 계산 (최소 1.0", 최대 2.5")
        max_items = max(len(p.get("details", [])) for p in phases)
        content_h = 0.08 + max_items * 0.20 + 0.08  # 패딩 + 항목들 + 패딩
        detail_h = min(max(content_h, 1.0), min(2.5, r.h - detail_y))
        card_gap = 0.06
        card_w = (r.w - card_gap * (n - 1)) / n
        for i, p in enumerate(phases):
            dx = i * (card_w + card_gap)
            c.box(x=dx, y=detail_y, w=card_w, h=detail_h,
                  fill="grey_100", border=0.5, border_color="grey_mid",
                  region=r)
            details = p.get("details", [])
            for di, item in enumerate(details):
                c.text(f"▪ {item}",
                       x=dx + 0.06, y=detail_y + 0.06 + di * 0.20,
                       w=card_w - 0.12, h=0.18,
                       size=8, color="grey_900", anchor="top", region=r)

    return chev_h if not show_details else detail_y + detail_h


# ============================================================
# Compound Component 2: Hero Block
# ============================================================

def comp_hero_block(
    c: Canvas,
    *,
    headline: str,
    sub_points: list[str] = None,
    label: str = "",
    bg_color: str = "grey_800",
    text_color: str = "white",
    region: Region,
) -> float:
    """대형 색상 박스 — 핵심 메시지 강조.

    executive_summary의 좌측 Hero 영역에서 추출.

    사용 맥락: 전략 방향 선언, 핵심 발견 강조, 섹션 도입부
    """
    r = region
    sub_points = sub_points or []

    # 배경
    c.box(x=0, y=0, w=r.w, h=r.h,
          fill=bg_color, border=None, region=r)

    cy = 0.20
    pad = 0.25

    # 레이블 칩 (선택) — label_chip은 region 미지원이므로 절대좌표 계산
    if label:
        c.label_chip(label, x=r.x + pad, y=r.y + cy,
                     w=min(1.5, r.w * 0.4), h=0.26,
                     fill="grey_400", text_color="white")
        cy += 0.40

    # 헤드라인
    hl_h = min(r.h * 0.35, 1.2)
    c.text(headline, x=pad, y=cy, w=r.w - pad * 2, h=hl_h,
           size=18, bold=True, color=text_color, anchor="top",
           region=r)
    cy += hl_h + 0.08

    # 구분선
    c.box(x=pad, y=cy, w=r.w - pad * 2, h=0.012,
          fill="grey_400", border=None, region=r)
    cy += 0.15

    # 하위 포인트
    for i, pt in enumerate(sub_points):
        if cy + 0.22 > r.h - 0.1:
            break
        c.text(f"▪  {pt}", x=pad, y=cy, w=r.w - pad * 2, h=0.22,
               size=9, color=text_color, anchor="top", region=r)
        cy += 0.24

    return r.h


# ============================================================
# Compound Component 3: Hub-Spoke Diagram
# ============================================================

def comp_hub_spoke_diagram(
    c: Canvas,
    *,
    center: str,
    center_sub: str = "",
    spokes: list[dict],       # [{"title": "...", "detail": "...", "badge": ""}, ...]
    center_color: str = "grey_800",
    spoke_color: str = "white",
    region: Region,
) -> float:
    """허브-스포크 방사형 다이어그램.

    hub_spoke 패턴에서 추출. 삼각함수 기반 자동 배치.

    사용 맥락: 시스템 통합 구조, 핵심 역량 + 영향 영역, 이해관계자 맵
    """
    import math
    r = region
    n = len(spokes)
    if n == 0:
        return r.h

    # 중심 원
    hub_cx = r.w / 2
    hub_cy = r.h / 2
    hub_d = min(r.w, r.h) * 0.25
    c.circle(x=hub_cx - hub_d / 2, y=hub_cy - hub_d / 2, d=hub_d,
             fill=center_color, border=None,
             text=center, text_color="white", text_size=12, text_bold=True,
             region=r)
    if center_sub:
        c.text(center_sub,
               x=hub_cx - hub_d / 2, y=hub_cy + hub_d * 0.15,
               w=hub_d, h=0.25,
               size=7, color="grey_200", align="center", anchor="top",
               region=r)

    # 스포크 크기 자동 조정
    spoke_r = min(r.w, r.h) * 0.38
    if n <= 4:
        sp_w, sp_h = min(r.w * 0.28, 2.0), min(r.h * 0.28, 1.0)
    elif n <= 6:
        sp_w, sp_h = min(r.w * 0.24, 1.7), min(r.h * 0.24, 0.85)
    else:
        sp_w, sp_h = min(r.w * 0.20, 1.5), min(r.h * 0.20, 0.75)

    angle_offset = -math.pi / 2

    for i, sp in enumerate(spokes):
        angle = angle_offset + (2 * math.pi * i / n)
        sx = hub_cx + spoke_r * math.cos(angle) - sp_w / 2
        sy = hub_cy + spoke_r * math.sin(angle) - sp_h / 2

        # 연결선
        lx1 = hub_cx + (hub_d / 2 + 0.03) * math.cos(angle)
        ly1 = hub_cy + (hub_d / 2 + 0.03) * math.sin(angle)
        lx2 = hub_cx + (spoke_r - sp_w / 2 - 0.03) * math.cos(angle)
        ly2 = hub_cy + (spoke_r - sp_h / 2 - 0.03) * math.sin(angle)
        c.line(x1=lx1, y1=ly1, x2=lx2, y2=ly2,
               color="grey_400", width=1.0, region=r)

        # 스포크 박스
        c.box(x=sx, y=sy, w=sp_w, h=sp_h,
              fill=spoke_color, border=0.75, border_color="grey_mid",
              region=r)
        c.box(x=sx, y=sy, w=sp_w, h=0.04,
              fill="grey_700", border=None, region=r)

        # 텍스트
        c.text(sp["title"],
               x=sx + 0.08, y=sy + 0.12, w=sp_w - 0.16, h=0.25,
               size=9, bold=True, color="grey_900", anchor="top",
               region=r)
        if sp.get("detail"):
            c.text(sp["detail"],
                   x=sx + 0.08, y=sy + 0.36, w=sp_w - 0.16, h=sp_h - 0.44,
                   size=7, color="grey_700", anchor="top",
                   region=r)

    return r.h


# ============================================================
# Compound Component 4: Comparison Grid
# ============================================================

def comp_comparison_grid(
    c: Canvas,
    *,
    columns: list[dict],      # [{"name": "Option A", "summary": "", "highlight": False, "criteria": [...]}, ...]
    row_labels: list[str],
    region: Region,
) -> float:
    """N열 비교 표 — 컬럼별 헤더(색상 구분) + 행별 비교 데이터.

    comparison_matrix에서 추출. highlight=True인 컬럼은 강조 배경.

    사용 맥락: 옵션 A/B/C 비교, 솔루션 벤더 비교, AS-IS vs TO-BE
    """
    r = region
    n_cols = len(columns)
    n_rows = len(row_labels)
    if n_cols == 0 or n_rows == 0:
        return 0.0

    label_w = min(r.w * 0.18, 1.4)
    col_gap = 0.04
    grid_x = label_w + col_gap
    grid_w = r.w - grid_x
    col_w = (grid_w - col_gap * (n_cols - 1)) / n_cols

    header_h = 0.45
    row_h = (r.h - header_h - 0.02) / n_rows

    # 헤더
    for i, col in enumerate(columns):
        ox = grid_x + i * (col_w + col_gap)
        is_hl = col.get("highlight", False)
        fill = "grey_900" if is_hl else "grey_700"
        c.box(x=ox, y=0, w=col_w, h=header_h,
              fill=fill, border=None, region=r)
        c.text(col["name"],
               x=ox + 0.06, y=0.06, w=col_w - 0.12, h=0.22,
               size=10, bold=True, color="white", align="center", anchor="top",
               region=r)
        if col.get("summary"):
            c.text(col["summary"],
                   x=ox + 0.06, y=0.28, w=col_w - 0.12, h=0.18,
                   size=7, color="grey_200", align="center", anchor="top",
                   region=r)

    # 행 라벨 + 셀 — PwC 수준 밀도: 진한 배경 + top-anchor 멀티라인
    row_fills_normal = ["grey_200", "grey_100"]   # 눈에 보이는 교대 행
    row_fills_hl = ["grey_400", "grey_200"]       # 하이라이트 교대 행

    for ri, label in enumerate(row_labels):
        ry = header_h + ri * row_h
        alt = ri % 2
        # 라벨
        c.box(x=0, y=ry, w=label_w, h=row_h,
              fill="grey_200", border=0.5, border_color="grey_mid", region=r)
        c.text(label, x=0.06, y=ry, w=label_w - 0.12, h=row_h,
               size=8, bold=True, color="grey_900", anchor="middle", region=r)

        # 셀
        for ci, col in enumerate(columns):
            ox = grid_x + ci * (col_w + col_gap)
            is_hl = col.get("highlight", False)
            cell_fill = row_fills_hl[alt] if is_hl else row_fills_normal[alt]
            crits = col.get("criteria", [])
            val = crits[ri] if ri < len(crits) else ""
            c.box(x=ox, y=ry, w=col_w, h=row_h,
                  fill=cell_fill, border=0.5, border_color="grey_mid", region=r)
            # top-anchor + 패딩: 긴 텍스트가 셀 안에서 자연스럽게 줄바꿈
            txt_color = "white" if is_hl and alt == 0 else "grey_900"
            c.text(str(val),
                   x=ox + 0.08, y=ry + 0.06, w=col_w - 0.16, h=row_h - 0.12,
                   size=8, color=txt_color, anchor="top", region=r)

    return r.h


# ============================================================
# Compound Component 5: Architecture Stack
# ============================================================

def comp_architecture_stack(
    c: Canvas,
    *,
    layers: list[dict],       # [{"title": "Presentation", "items": ["React", "Next.js"]}, ...] (top→bottom)
    style: str = "gradient",  # "gradient" | "alternating" | "uniform"
    region: Region,
) -> float:
    """수직 레이어 스택 (기술 아키텍처).

    architecture_stack에서 추출. 가장 단순한 compound component.

    사용 맥락: 기술 스택, 시스템 레이어, 조직 계층
    """
    r = region
    n = len(layers)
    if n == 0:
        return 0.0

    gap = 0.05
    layer_h = (r.h - gap * (n - 1)) / n

    if style == "gradient":
        fills = ["grey_900", "grey_800", "grey_700", "grey_400", "grey_200", "grey_100"]
        txts = ["white", "white", "white", "white", "grey_900", "grey_900"]
    elif style == "alternating":
        fills = ["grey_800", "grey_200"] * (n // 2 + 1)
        txts = ["white", "grey_900"] * (n // 2 + 1)
    else:
        fills = ["grey_700"] * n
        txts = ["white"] * n

    title_w = min(r.w * 0.28, 2.2)

    for i, layer in enumerate(layers):
        ly = i * (layer_h + gap)
        fi = fills[min(i, len(fills) - 1)]
        tc = txts[min(i, len(txts) - 1)]

        c.box(x=0, y=ly, w=r.w, h=layer_h,
              fill=fi, border=None, region=r)

        # 레이어 타이틀 (좌측)
        c.text(layer["title"],
               x=0.12, y=ly, w=title_w, h=layer_h,
               size=10, bold=True, color=tc, anchor="middle",
               region=r)

        # 아이템 (우측 균등 배치)
        items = layer.get("items", [])
        if items:
            item_area_w = r.w - title_w - 0.15
            item_w = item_area_w / max(len(items), 1)
            for ji, item in enumerate(items):
                c.text(item,
                       x=title_w + 0.10 + ji * item_w, y=ly,
                       w=item_w, h=layer_h,
                       size=8, color=tc, align="center", anchor="middle",
                       region=r)

    return r.h


# ============================================================
# Compound Component 6: Pyramid
# ============================================================

def comp_pyramid(
    c: Canvas,
    *,
    layers: list[dict],
    style: str = "gradient",
    region: Region,
) -> float:
    """피라미드 레이어 — 상위일수록 좁은 사다리꼴 스택.

    pyramid_layers에서 추출. 전략 위계, 성숙도 모델, 조직 계층.
    """
    r = region
    n = len(layers)
    if n == 0:
        return 0.0

    gap = 0.05
    layer_h = (r.h - gap * (n - 1)) / n
    w_min = r.w * 0.30
    w_max = r.w * 0.95

    # 색상: 상위(좁은) 레이어에 accent 컬러로 시각적 무게감 부여
    if style == "warm":
        fills = ["accent", "accent_mid", "grey_700", "grey_400", "grey_200"]
        txts = ["white", "white", "white", "white", "grey_900"]
    else:  # gradient (default)
        fills = ["grey_900", "grey_700", "accent_mid", "grey_400", "grey_200"]
        txts = ["white", "white", "white", "white", "grey_900"]

    for i, layer in enumerate(layers):
        ratio = i / max(n - 1, 1)
        lw = w_min + (w_max - w_min) * ratio
        lx = (r.w - lw) / 2
        ly = i * (layer_h + gap)
        fi = fills[min(i, len(fills) - 1)]
        tc = txts[min(i, len(txts) - 1)]
        c.box(x=lx, y=ly, w=lw, h=layer_h, fill=fi, border=None, region=r)
        c.text(layer["title"], x=lx + 0.1, y=ly, w=lw - 0.2, h=layer_h * 0.55,
               size=11, bold=True, color=tc, align="center", anchor="middle", region=r)
        if layer.get("detail"):
            c.text(layer["detail"], x=lx + 0.1, y=ly + layer_h * 0.50,
                   w=lw - 0.2, h=layer_h * 0.45, size=8, color=tc,
                   align="center", anchor="top", region=r)

    return r.h


# ============================================================
# Compound Component 7: Cycle Arrows
# ============================================================

def comp_cycle_arrows(
    c: Canvas,
    *,
    center: str,
    center_sub: str = "",
    stages: list[dict],
    center_color: str = "grey_900",
    region: Region,
) -> float:
    """원형 순환 다이어그램 — PDCA, 플라이휠.

    cycle_diagram에서 추출. 순환 프로세스, 반복 개선 구조.
    """
    import math
    r = region
    n = len(stages)
    if n == 0:
        return r.h

    cx, cy = r.w / 2, r.h / 2
    center_d = min(r.w, r.h) * 0.22
    c.circle(x=cx - center_d / 2, y=cy - center_d / 2, d=center_d,
             fill=center_color, border=None,
             text=center, text_color="white", text_size=11, text_bold=True, region=r)
    if center_sub:
        c.text(center_sub, x=cx - center_d / 2, y=cy + center_d * 0.15,
               w=center_d, h=0.22, size=7, color="grey_200",
               align="center", anchor="top", region=r)

    ring_r = min(r.w, r.h) * 0.20
    card_r = min(r.w, r.h) * 0.40
    card_w = min(r.w * 0.24, 1.7)
    card_h = min(r.h * 0.22, 0.85)
    angle_offset = -math.pi / 2

    for i, st in enumerate(stages):
        angle = angle_offset + (2 * math.pi * i / n)
        next_angle = angle_offset + (2 * math.pi * (i + 1) / n)
        mid_angle = (angle + next_angle) / 2

        arr_a = angle_offset + (2 * math.pi * (i + 0.75) / n)
        a2_a = angle_offset + (2 * math.pi * (i + 1.05) / n)
        c.arrow(x1=cx + (ring_r + 0.15) * math.cos(arr_a),
                y1=cy + (ring_r + 0.15) * math.sin(arr_a),
                x2=cx + (ring_r + 0.15) * math.cos(a2_a),
                y2=cy + (ring_r + 0.15) * math.sin(a2_a),
                color="grey_400", width=0.75, region=r)

        card_cx = cx + card_r * math.cos(mid_angle)
        card_cy = cy + card_r * math.sin(mid_angle)
        card_x, card_y = card_cx - card_w / 2, card_cy - card_h / 2

        c.line(x1=cx + (ring_r + 0.25) * math.cos(mid_angle),
               y1=cy + (ring_r + 0.25) * math.sin(mid_angle),
               x2=card_cx, y2=card_cy, color="grey_mid", width=0.5, region=r)

        c.box(x=card_x, y=card_y, w=card_w, h=card_h,
              fill="white", border=0.75, border_color="grey_mid", region=r)
        c.text(st["title"], x=card_x + 0.06, y=card_y + 0.06,
               w=card_w - 0.12, h=0.22, size=9, bold=True,
               color="grey_900", anchor="top", region=r)
        if st.get("detail"):
            c.text(st["detail"], x=card_x + 0.06, y=card_y + 0.28,
                   w=card_w - 0.12, h=card_h - 0.34, size=7,
                   color="grey_700", anchor="top", region=r)

    return r.h


# ============================================================
# Compound Component 8: Waterfall
# ============================================================

def comp_waterfall(
    c: Canvas,
    *,
    start: dict,
    steps: list[dict],
    end: dict,
    unit: str = "",
    region: Region,
) -> float:
    """증감 바 차트 (Waterfall/Bridge).

    waterfall_bridge에서 추출. 비용 분해, 매출 Bridge, 시간 분석.
    """
    r = region
    all_items = [start] + steps + [end]
    n = len(all_items)
    if n == 0:
        return 0.0

    bar_gap = 0.06
    bar_w = (r.w - bar_gap * (n - 1)) / n
    cumulative = [start["value"]]
    for s in steps:
        cumulative.append(cumulative[-1] + s["value"])
    all_values = cumulative + [end["value"]]
    max_val = max(abs(v) for v in all_values) if all_values else 1

    label_h = 0.30
    chart_h = r.h - label_h - 0.05
    baseline_y = chart_h * 0.72

    def val_to_y(v):
        if max_val == 0:
            return baseline_y
        return baseline_y - v * (chart_h * 0.60) / max_val

    c.line(x1=0, y1=baseline_y, x2=r.w, y2=baseline_y,
           color="grey_mid", width=0.5, region=r)

    running = start["value"]
    for i, item in enumerate(all_items):
        bx = i * (bar_w + bar_gap)
        is_start, is_end = (i == 0), (i == n - 1)

        if is_start or is_end:
            val = item["value"]
            bar_top = val_to_y(val)
            bh = abs(baseline_y - bar_top)
            c.box(x=bx, y=min(bar_top, baseline_y), w=bar_w, h=max(bh, 0.04),
                  fill="grey_800", border=None, region=r)
        else:
            step_val = item["value"]
            prev_cum = running
            running += step_val
            bar_top = val_to_y(max(prev_cum, running))
            bar_bot = val_to_y(min(prev_cum, running))
            c.box(x=bx, y=bar_top, w=bar_w, h=max(abs(bar_bot - bar_top), 0.04),
                  fill="positive" if step_val > 0 else "negative", border=None, region=r)
            if i > 0:
                prev_x = (i - 1) * (bar_w + bar_gap) + bar_w
                c.line(x1=prev_x, y1=val_to_y(prev_cum), x2=bx, y2=val_to_y(prev_cum),
                       color="grey_400", width=0.5, region=r)

        c.text(item["label"], x=bx, y=chart_h + 0.03, w=bar_w, h=label_h,
               size=7, bold=True, color="grey_900", align="center", anchor="top", region=r)

        val = item["value"]
        prefix = "+" if (not is_start and not is_end) and val > 0 else ""
        val_str = f"{prefix}{val:,.0f}{unit}"
        vy = val_to_y(abs(val) if (is_start or is_end) else max(running, running - val)) - 0.22
        c.text(val_str, x=bx, y=vy, w=bar_w, h=0.20,
               size=8, bold=True, color="grey_900", align="center", anchor="top", region=r)

    return r.h


# ============================================================
# Compound Component 9: Before-After
# ============================================================

def comp_before_after(
    c: Canvas,
    *,
    before_title: str = "AS-IS",
    after_title: str = "TO-BE",
    before_items: list[dict],
    after_items: list[dict],
    arrow_label: str = "",
    region: Region,
) -> float:
    """좌우 대비 패널 — AS-IS / TO-BE 변화 표현.

    before_after에서 추출. 전환, 개선, 혁신 시나리오.
    """
    r = region
    panel_w = (r.w - 0.65) / 2
    header_h = 0.42

    c.box(x=0, y=0, w=panel_w, h=r.h,
          fill="grey_200", border=0.75, border_color="grey_mid", region=r)
    c.box(x=0, y=0, w=panel_w, h=header_h,
          fill="grey_700", border=None, region=r)
    c.text(before_title, x=0.1, y=0, w=panel_w - 0.2, h=header_h,
           size=12, bold=True, color="white", anchor="middle", region=r)

    ax = panel_w + 0.65
    c.box(x=ax, y=0, w=panel_w, h=r.h,
          fill="white", border=0.75, border_color="grey_mid", region=r)
    c.box(x=ax, y=0, w=panel_w, h=header_h,
          fill="grey_900", border=None, region=r)
    c.text(after_title, x=ax + 0.1, y=0, w=panel_w - 0.2, h=header_h,
           size=12, bold=True, color="white", anchor="middle", region=r)

    arrow_y = r.h / 2
    c.arrow(x1=panel_w + 0.05, y1=arrow_y, x2=ax - 0.05, y2=arrow_y,
            color="grey_900", width=2.0, region=r)
    if arrow_label:
        c.text(arrow_label, x=panel_w + 0.05, y=arrow_y - 0.28,
               w=0.55, h=0.22, size=7, bold=True, color="grey_900",
               align="center", anchor="bottom", region=r)

    n_items = max(len(before_items), len(after_items))
    item_h = min(0.80, (r.h - header_h - 0.15) / max(n_items, 1))

    for i, item in enumerate(before_items):
        iy = header_h + 0.10 + i * item_h
        c.text(item["label"], x=0.12, y=iy, w=panel_w - 0.24, h=0.20,
               size=9, bold=True, color="grey_900", anchor="top", region=r)
        if item.get("detail"):
            c.text(item["detail"], x=0.12, y=iy + 0.20, w=panel_w - 0.24, h=0.25,
                   size=8, color="grey_700", anchor="top", region=r)
        if item.get("kpi"):
            c.text(item["kpi"], x=0.12, y=iy + item_h - 0.22, w=panel_w - 0.24, h=0.18,
                   size=8, bold=True, color="negative", anchor="top", region=r)

    for i, item in enumerate(after_items):
        iy = header_h + 0.10 + i * item_h
        c.text(item["label"], x=ax + 0.12, y=iy, w=panel_w - 0.24, h=0.20,
               size=9, bold=True, color="grey_900", anchor="top", region=r)
        if item.get("detail"):
            c.text(item["detail"], x=ax + 0.12, y=iy + 0.20, w=panel_w - 0.24, h=0.25,
                   size=8, color="grey_700", anchor="top", region=r)
        if item.get("kpi"):
            c.text(item["kpi"], x=ax + 0.12, y=iy + item_h - 0.22, w=panel_w - 0.24, h=0.18,
                   size=8, bold=True, color="positive", anchor="top", region=r)

    return r.h


# ============================================================
# Compound Component 10: Gantt Bars
# ============================================================

def comp_gantt_bars(
    c: Canvas,
    *,
    phases: list[str],
    streams: list[dict],
    milestones: list[dict] = None,
    region: Region,
) -> float:
    """멀티 스트림 Gantt 타임라인.

    gantt_roadmap에서 추출. 프로젝트 일정, 프로그램 로드맵.
    """
    r = region
    milestones = milestones or []
    n_phases = len(phases)
    n_streams = len(streams)
    if n_phases == 0 or n_streams == 0:
        return 0.0

    label_w = min(r.w * 0.20, 1.6)
    grid_x = label_w + 0.05
    grid_w = r.w - grid_x
    phase_h = 0.35
    col_w = grid_w / n_phases
    body_y = phase_h + 0.06
    body_h = r.h - body_y
    row_h = body_h / n_streams

    pfills = ["grey_800", "grey_700", "grey_400", "grey_200"]
    ptxt = ["white", "white", "white", "grey_900"]
    for j, phase in enumerate(phases):
        px = grid_x + j * col_w
        c.box(x=px, y=0, w=col_w, h=phase_h,
              fill=pfills[j % len(pfills)], border=None, region=r)
        c.text(phase, x=px, y=0, w=col_w, h=phase_h,
               size=8, bold=True, color=ptxt[j % len(ptxt)],
               align="center", anchor="middle", region=r)

    for j in range(1, n_phases):
        c.line(x1=grid_x + j * col_w, y1=body_y,
               x2=grid_x + j * col_w, y2=body_y + body_h,
               color="grey_mid", width=0.3, region=r)

    for i, stream in enumerate(streams):
        sy = body_y + i * row_h
        sf = "grey_100" if i % 2 == 0 else "white"
        c.box(x=0, y=sy, w=label_w, h=row_h, fill=sf,
              border=0.5, border_color="grey_mid", region=r)
        c.text(stream["name"], x=0.05, y=sy, w=label_w - 0.10, h=row_h,
               size=8, bold=True, color="grey_900", anchor="middle", region=r)
        c.box(x=grid_x, y=sy, w=grid_w, h=row_h, fill=sf,
              border=0.5, border_color="grey_mid", region=r)
        bp = row_h * 0.18
        for bar in stream.get("bars", []):
            bx = grid_x + bar["start"] * col_w + 0.03
            bw = (bar["end"] - bar["start"]) * col_w - 0.06
            is_hl = bar.get("highlight", False)
            c.box(x=bx, y=sy + bp, w=bw, h=row_h - 2 * bp,
                  fill="grey_900" if is_hl else "grey_400", border=None, region=r)
            if bar.get("label"):
                c.text(bar["label"], x=bx + 0.04, y=sy + bp,
                       w=bw - 0.08, h=row_h - 2 * bp,
                       size=7, bold=is_hl, color="white", anchor="middle", region=r)

    for ms in milestones:
        mx = grid_x + ms["phase"] * col_w
        c.line(x1=mx, y1=body_y - 0.03, x2=mx, y2=body_y + body_h + 0.03,
               color="negative", width=1.5, region=r)
        c.text(ms.get("label", ""), x=mx - 0.4, y=body_y + body_h + 0.04,
               w=0.8, h=0.18, size=6, bold=True, color="negative",
               align="center", anchor="top", region=r)

    return r.h


# ============================================================
# Compound Component 11: Value Chain (Porter)
# ============================================================

def comp_value_chain(
    c: Canvas,
    *,
    support: list[dict],      # [{"title": "HR", "detail": "인재 관리"}, ...]
    primary: list[dict],      # [{"title": "Inbound", "detail": "...", "highlight": False}, ...]
    margin_label: str = "Margin",
    region: Region,
) -> float:
    """Porter 가치사슬 — 지원활동(수평 행) + 주요활동(쉐브론) + 마진.

    value_chain에서 추출. 전략 분석, 비즈니스 모델, 경쟁우위 분석.
    """
    r = region
    n_support = len(support)
    n_primary = len(primary)
    if n_primary == 0:
        return 0.0

    margin_w = min(r.w * 0.12, 1.1)
    body_w = r.w - margin_w - 0.08
    sup_h = 0.42
    sup_gap = 0.04

    # 지원활동 (수평 행)
    for i, sup in enumerate(support):
        sy = i * (sup_h + sup_gap)
        c.box(x=0, y=sy, w=body_w, h=sup_h,
              fill="grey_100", border=0.5, border_color="grey_mid", region=r)
        c.text(sup["title"], x=0.10, y=sy, w=min(1.8, body_w * 0.25), h=sup_h,
               size=8, bold=True, color="grey_900", anchor="middle", region=r)
        if sup.get("detail"):
            c.text(sup["detail"], x=min(1.8, body_w * 0.25) + 0.15, y=sy,
                   w=body_w - min(1.8, body_w * 0.25) - 0.25, h=sup_h,
                   size=7, color="grey_700", anchor="middle", region=r)

    # 주요활동 (쉐브론)
    prim_y = n_support * (sup_h + sup_gap) + 0.10
    chev_h = min(0.45, (r.h - prim_y) * 0.35)
    overlap = 0.08
    chev_w = (body_w + overlap * (n_primary - 1)) / n_primary
    cfills = ["grey_800", "grey_700", "grey_400", "grey_200", "grey_400"]
    ctxt = ["white", "white", "white", "grey_900", "white"]

    for i, prim in enumerate(primary):
        px = i * (chev_w - overlap)
        is_hl = prim.get("highlight", False)
        fill = "grey_900" if is_hl else cfills[i % len(cfills)]
        tc = "white" if is_hl else ctxt[i % len(ctxt)]
        c.chevron(x=px, y=prim_y, w=chev_w, h=chev_h,
                  fill=fill, text=prim["title"], text_color=tc, text_size=8, region=r)

    # 상세 카드
    detail_y = prim_y + chev_h + 0.08
    detail_h = r.h - detail_y
    if detail_h > 0.3:
        card_w = body_w / n_primary - 0.06
        for i, prim in enumerate(primary):
            if prim.get("detail"):
                dx = i * (body_w / n_primary) + 0.03
                c.box(x=dx, y=detail_y, w=card_w, h=detail_h,
                      fill="white", border=0.5, border_color="grey_mid", region=r)
                c.text(prim["detail"], x=dx + 0.06, y=detail_y + 0.04,
                       w=card_w - 0.12, h=detail_h - 0.08,
                       size=7, color="grey_900", anchor="top", region=r)

    # 마진 (우측 세로 블록)
    total_h = r.h
    c.box(x=body_w + 0.08, y=0, w=margin_w, h=total_h,
          fill="grey_800", border=None, region=r)
    c.text(margin_label, x=body_w + 0.08, y=total_h * 0.3,
           w=margin_w, h=total_h * 0.4,
           size=11, bold=True, color="white", align="center", anchor="middle",
           region=r)

    return r.h


# ============================================================
# Compound Component 12: Logic Tree
# ============================================================

def comp_logic_tree(
    c: Canvas,
    *,
    root: str,
    branches: list[dict],     # [{"title": "...", "children": ["...", "..."], "highlight": False}, ...]
    region: Region,
) -> float:
    """계층 분기 트리 — MECE 이슈 트리, 수익 분해, 조직도.

    tree_diagram에서 추출. 논리 분해, 의사결정 구조에 사용.
    """
    r = region
    n_br = len(branches)
    if n_br == 0:
        return r.h

    root_w = min(r.w * 0.22, 2.0)
    root_h = min(r.h * 0.18, 0.75)
    root_x = 0
    root_y = (r.h - root_h) / 2

    c.box(x=root_x, y=root_y, w=root_w, h=root_h,
          fill="grey_900", border=None, region=r)
    c.text(root, x=root_x + 0.08, y=root_y, w=root_w - 0.16, h=root_h,
           size=11, bold=True, color="white", anchor="middle", region=r)

    br_x = root_w + r.w * 0.06
    br_w = min(r.w * 0.22, 2.0)
    br_gap = 0.08
    br_h = (r.h - br_gap * (n_br - 1)) / n_br

    for i, br in enumerate(branches):
        by = i * (br_h + br_gap)
        is_hl = br.get("highlight", False)
        fill = "grey_800" if is_hl else "grey_200"
        tc = "white" if is_hl else "grey_900"

        # 루트 → 브랜치 연결선
        c.line(x1=root_w, y1=root_y + root_h / 2,
               x2=br_x, y2=by + br_h / 2,
               color="grey_700", width=0.75, region=r)

        c.box(x=br_x, y=by, w=br_w, h=br_h,
              fill=fill, border=0.75 if is_hl else 0.5,
              border_color="grey_900" if is_hl else "grey_mid", region=r)
        c.text(br["title"], x=br_x + 0.08, y=by + 0.04, w=br_w - 0.16, h=0.25,
               size=9, bold=True, color=tc, anchor="top", region=r)

        # 자식 노드
        children = br.get("children", [])
        if children:
            leaf_x = br_x + br_w + r.w * 0.04
            leaf_w = r.w - leaf_x
            leaf_h = min(0.35, (br_h - 0.04) / max(len(children), 1))
            lg = (br_h - leaf_h * len(children)) / max(len(children) + 1, 1)
            for j, child in enumerate(children):
                ly = by + lg * (j + 1) + leaf_h * j
                c.line(x1=br_x + br_w, y1=by + br_h / 2,
                       x2=leaf_x, y2=ly + leaf_h / 2,
                       color="grey_400", width=0.5, region=r)
                c.box(x=leaf_x, y=ly, w=leaf_w, h=leaf_h,
                      fill="grey_100", border=0.5, border_color="grey_mid", region=r)
                c.text(child, x=leaf_x + 0.06, y=ly, w=leaf_w - 0.12, h=leaf_h,
                       size=7, color="grey_900", anchor="middle", region=r)

    return r.h


# ============================================================
# Compound Component 13: Quadrant Matrix (2x2)
# ============================================================

def comp_quadrant_matrix(
    c: Canvas,
    *,
    quadrants: list[dict],    # [{"title": "Stars", "items": ["...", "..."]}, ...] (TL, TR, BL, BR)
    x_axis: str = "",
    y_axis: str = "",
    x_low: str = "",
    x_high: str = "",
    y_low: str = "",
    y_high: str = "",
    region: Region,
) -> float:
    """2x2 포지셔닝 매트릭스 — BCG, SWOT, 우선순위.

    quadrant_story에서 코어 추출. 축 라벨 + 4 사분면.
    """
    r = region
    quads = quadrants[:4]
    while len(quads) < 4:
        quads.append({"title": "", "items": []})

    axis_pad = 0.45
    grid_x = axis_pad
    grid_y = 0
    grid_w = r.w - axis_pad
    grid_h = r.h - 0.30  # 하단 x축 라벨 공간
    gap = 0.08
    cell_w = (grid_w - gap) / 2
    cell_h = (grid_h - gap) / 2

    fills = ["grey_200", "grey_100", "grey_100", "grey_200"]
    stripe_fills = ["grey_800", "grey_700", "grey_700", "grey_400"]

    positions = [
        (grid_x, grid_y),                        # TL
        (grid_x + cell_w + gap, grid_y),          # TR
        (grid_x, grid_y + cell_h + gap),          # BL
        (grid_x + cell_w + gap, grid_y + cell_h + gap),  # BR
    ]

    for qi, q in enumerate(quads):
        qx, qy = positions[qi]
        c.box(x=qx, y=qy, w=cell_w, h=cell_h,
              fill=fills[qi], border=0.5, border_color="grey_mid", region=r)
        # 상단 스트라이프
        c.box(x=qx, y=qy, w=cell_w, h=0.06,
              fill=stripe_fills[qi], border=None, region=r)
        # 타이틀
        c.text(q["title"], x=qx + 0.12, y=qy + 0.12, w=cell_w - 0.24, h=0.25,
               size=10, bold=True, color="grey_900", anchor="top", region=r)
        # 아이템
        for ii, item in enumerate(q.get("items", [])):
            iy = qy + 0.40 + ii * 0.22
            if iy + 0.20 > qy + cell_h - 0.05:
                break
            c.text(f"▪  {item}", x=qx + 0.12, y=iy, w=cell_w - 0.24, h=0.20,
                   size=8, color="grey_900", anchor="top", region=r)

    # 축 라벨
    if y_high:
        c.text(y_high, x=0, y=grid_y, w=axis_pad - 0.05, h=0.22,
               size=7, bold=True, color="grey_700", align="right", anchor="top", region=r)
    if y_low:
        c.text(y_low, x=0, y=grid_y + grid_h - 0.22, w=axis_pad - 0.05, h=0.22,
               size=7, bold=True, color="grey_700", align="right", anchor="bottom", region=r)
    if y_axis:
        c.text(y_axis, x=0, y=grid_y + grid_h / 2 - 0.10, w=axis_pad - 0.05, h=0.20,
               size=6, color="grey_700", align="right", anchor="middle", region=r)
    if x_low:
        c.text(x_low, x=grid_x, y=grid_h + 0.06, w=cell_w, h=0.20,
               size=7, bold=True, color="grey_700", align="left", anchor="top", region=r)
    if x_high:
        c.text(x_high, x=grid_x + cell_w + gap, y=grid_h + 0.06, w=cell_w, h=0.20,
               size=7, bold=True, color="grey_700", align="right", anchor="top", region=r)

    return r.h


# ============================================================
# Compound Component 14: Funnel
# ============================================================

def comp_funnel(
    c: Canvas,
    *,
    stages: list[dict],       # [{"label": "TAM", "value": "₩500B", "detail": "..."}, ...]
    style: str = "gradient",
    region: Region,
) -> float:
    """퍼널 다이어그램 — 매출 파이프라인, TAM/SAM/SOM.

    신규 컴포넌트. 위에서 아래로 폭이 줄어드는 사다리꼴.
    """
    r = region
    n = len(stages)
    if n == 0:
        return 0.0

    gap = 0.04
    stage_h = (r.h - gap * (n - 1)) / n
    w_max = r.w * 0.90
    w_min = r.w * 0.30

    fills = ["grey_900", "grey_800", "grey_700", "grey_400", "grey_200"]
    txts = ["white", "white", "white", "white", "grey_900"]

    for i, st in enumerate(stages):
        ratio = i / max(n - 1, 1)
        sw = w_max - (w_max - w_min) * ratio
        sx = (r.w - sw) / 2
        sy = i * (stage_h + gap)
        fi = fills[min(i, len(fills) - 1)]
        tc = txts[min(i, len(txts) - 1)]

        c.box(x=sx, y=sy, w=sw, h=stage_h, fill=fi, border=None, region=r)

        # 라벨 + 값
        label_text = st.get("label", "")
        value_text = st.get("value", "")
        display = f"{label_text}    {value_text}" if value_text else label_text
        c.text(display, x=sx + 0.10, y=sy, w=sw - 0.20, h=stage_h * 0.55,
               size=11, bold=True, color=tc, align="center", anchor="middle", region=r)

        if st.get("detail"):
            c.text(st["detail"], x=sx + 0.10, y=sy + stage_h * 0.52,
                   w=sw - 0.20, h=stage_h * 0.45,
                   size=8, color=tc, align="center", anchor="top", region=r)

    return r.h


# ============================================================
# Compound Component 15: Callout Annotation
# ============================================================

def comp_callout_annotation(
    c: Canvas,
    *,
    annotations: list[dict],  # [{"x": 0.5, "y": 0.3, "text": "핵심 포인트", "anchor": "right"}, ...]
    region: Region,
) -> float:
    """주석 박스 + 리더 라인 — 차트/도형 위에 주석 오버레이.

    신규 컴포넌트. 데이터 포인트에 설명을 붙여 밀도와 전문성 향상.
    anchor: "left" | "right" | "top" | "bottom" — 박스가 포인트의 어느 쪽에 위치할지.
    """
    r = region

    for ann in annotations:
        px = ann["x"] * r.w  # 비율 좌표 (0~1)
        py = ann["y"] * r.h
        text = ann["text"]
        anchor_dir = ann.get("anchor", "right")
        box_w = min(r.w * 0.25, 2.0)
        box_h = 0.35

        if anchor_dir == "right":
            bx, by = px + 0.15, py - box_h / 2
        elif anchor_dir == "left":
            bx, by = px - box_w - 0.15, py - box_h / 2
        elif anchor_dir == "top":
            bx, by = px - box_w / 2, py - box_h - 0.15
        else:  # bottom
            bx, by = px - box_w / 2, py + 0.15

        # 리더 라인
        line_to_x = bx if anchor_dir == "right" else bx + box_w
        if anchor_dir in ("top", "bottom"):
            line_to_x = bx + box_w / 2
            line_to_y = by + box_h if anchor_dir == "top" else by
        else:
            line_to_y = by + box_h / 2

        c.line(x1=px, y1=py, x2=line_to_x, y2=line_to_y,
               color="grey_700", width=0.75, region=r)

        # 주석 박스
        c.box(x=bx, y=by, w=box_w, h=box_h,
              fill="white", border=1.0, border_color="grey_700", region=r)
        c.text(text, x=bx + 0.06, y=by + 0.02, w=box_w - 0.12, h=box_h - 0.04,
               size=8, bold=True, color="grey_900", anchor="middle", region=r)

    return r.h


# ============================================================
# Compound Component 16: Heatmap Grid
# ============================================================

def comp_heatmap_grid(
    c: Canvas,
    *,
    row_labels: list[str],
    col_labels: list[str],
    values: list[list[float]],  # 0.0~1.0 (0=green, 0.5=yellow, 1.0=red)
    cell_texts: list[list[str]] = None,  # 셀 안에 표시할 텍스트 (선택)
    region: Region,
) -> float:
    """2D 히트맵 매트릭스 — 리스크 맵, 역량 평가, 우선순위.

    신규 컴포넌트. 연속 색상 스케일(녹→황→적)로 정량 평가.
    """
    from ppt_builder.primitives import color as resolve_color
    from pptx.dml.color import RGBColor

    r = region
    n_rows = len(row_labels)
    n_cols = len(col_labels)
    if n_rows == 0 or n_cols == 0:
        return 0.0

    cell_texts = cell_texts or [[""]*n_cols for _ in range(n_rows)]
    label_w = min(r.w * 0.18, 1.5)
    header_h = 0.35
    grid_x = label_w + 0.05
    grid_w = r.w - grid_x
    grid_h = r.h - header_h
    col_w = grid_w / n_cols
    row_h = grid_h / n_rows

    def _heat_color(v: float) -> str:
        """0.0=green, 0.5=yellow, 1.0=red."""
        v = max(0.0, min(1.0, v))
        if v <= 0.5:
            # green → yellow
            t = v * 2
            r_val = int(0x4C + (0xF5 - 0x4C) * t)
            g_val = int(0xAF + (0xA6 - 0xAF) * t)
            b_val = int(0x50 + (0x23 - 0x50) * t)
        else:
            # yellow → red
            t = (v - 0.5) * 2
            r_val = int(0xF5 + (0xE5 - 0xF5) * t)
            g_val = int(0xA6 + (0x39 - 0xA6) * t)
            b_val = int(0x23 + (0x3B - 0x23) * t)
        return f"#{r_val:02X}{g_val:02X}{b_val:02X}"

    # 컬럼 헤더
    for j, cl in enumerate(col_labels):
        cx = grid_x + j * col_w
        c.box(x=cx, y=0, w=col_w, h=header_h,
              fill="grey_700", border=0.5, border_color="grey_mid", region=r)
        c.text(cl, x=cx + 0.04, y=0, w=col_w - 0.08, h=header_h,
               size=8, bold=True, color="white", align="center", anchor="middle", region=r)

    # 행
    for i, rl in enumerate(row_labels):
        ry = header_h + i * row_h
        # 라벨
        c.box(x=0, y=ry, w=label_w, h=row_h,
              fill="grey_100", border=0.5, border_color="grey_mid", region=r)
        c.text(rl, x=0.06, y=ry, w=label_w - 0.12, h=row_h,
               size=8, bold=True, color="grey_900", anchor="middle", region=r)
        # 셀
        for j in range(n_cols):
            cx = grid_x + j * col_w
            v = values[i][j] if i < len(values) and j < len(values[i]) else 0.5
            heat = _heat_color(v)
            c.box(x=cx, y=ry, w=col_w, h=row_h,
                  fill=heat, border=0.5, border_color="grey_mid", region=r)
            ct = cell_texts[i][j] if i < len(cell_texts) and j < len(cell_texts[i]) else ""
            if ct:
                txt_color = "white" if v > 0.6 else "grey_900"
                c.text(ct, x=cx + 0.04, y=ry, w=col_w - 0.08, h=row_h,
                       size=8, bold=True, color=txt_color,
                       align="center", anchor="middle", region=r)

    return r.h
