"""Region-aware 컴포넌트 라이브러리.

각 컴포넌트는 (canvas, spec, region) → height_used 시그니처.
Region 안에서 상대좌표로 렌더하고, 사용한 높이를 반환한다.

사용 예:
    from ppt_builder.primitives import Canvas, Region
    from ppt_builder.components import comp_kpi_row

    region = Region(0.3, 1.7, 9.4, 2.0)
    h = comp_kpi_row(canvas, kpis=[...], region=region)
"""

from __future__ import annotations

from ppt_builder.primitives import Canvas, Region
from ppt_builder.assembler.styles import estimate_text_height


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
