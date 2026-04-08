"""레이아웃 엔진 - 컴포넌트를 슬라이드 내에 자동 배치.

MARGIN 0.4", 콘텐츠 영역 9.2" x 5.7" 기준.
"""

from pptx.util import Inches

from .styles import (
    CONTENT_X, CONTENT_Y, CONTENT_W, CONTENT_H,
    COL_GAP, ROW_GAP,
    calc_columns, calc_rows,
)
from ..models.enums import LayoutType


def calculate_positions(
    layout: LayoutType,
    n_elements: int,
    n_cols: int = 1,
    elements: list | None = None,
) -> list[tuple[int, int, int, int]]:
    """(x, y, w, h) EMU 단위 좌표 리스트를 반환."""
    cx, cy, cw, ch = CONTENT_X, CONTENT_Y, CONTENT_W, CONTENT_H

    if layout == LayoutType.FULL:
        return _layout_full(cx, cy, cw, ch, n_elements)
    elif layout == LayoutType.COLUMNS:
        return _layout_columns(cx, cy, cw, ch, n_elements, n_cols, elements)
    elif layout == LayoutType.PROCESS:
        return _layout_process(cx, cy, cw, ch, n_elements)
    elif layout == LayoutType.TABLE:
        return _layout_table(cx, cy, cw, ch, n_elements)
    elif layout == LayoutType.MIXED:
        return _layout_mixed(cx, cy, cw, ch, n_elements, elements)
    return _layout_full(cx, cy, cw, ch, n_elements)


def _layout_full(cx, cy, cw, ch, n):
    if n == 1:
        return [(cx, cy, cw, ch)]
    gap = ROW_GAP
    h = int((ch - gap * (n - 1)) / n)
    return [(cx, cy + i * (h + gap), cw, h) for i in range(n)]


def _layout_columns(cx, cy, cw, ch, n, n_cols, elements=None):
    if n_cols <= 1:
        n_cols = min(n, 4)

    cols = calc_columns(n_cols, gap=0.15, start_x=0.3, total_w=9.4)

    # 분배
    col_elems: dict[int, list[int]] = {i: [] for i in range(n_cols)}
    for idx in range(n):
        assigned = None
        if elements and idx < len(elements):
            assigned = getattr(elements[idx], 'col', None)
        if assigned is not None and 0 <= assigned < n_cols:
            col_elems[assigned].append(idx)
        else:
            min_col = min(range(n_cols), key=lambda c: len(col_elems[c]))
            col_elems[min_col].append(idx)

    positions = [None] * n
    for ci in range(n_cols):
        col_x_inch, col_w_inch = cols[ci]
        col_x = Inches(col_x_inch)
        col_w = Inches(col_w_inch)
        elems = col_elems[ci]
        if not elems:
            continue
        gap = ROW_GAP
        elem_h = int((ch - gap * (len(elems) - 1)) / len(elems))
        for i, ei in enumerate(elems):
            positions[ei] = (col_x, cy + i * (elem_h + gap), col_w, elem_h)

    for i in range(n):
        if positions[i] is None:
            positions[i] = (cx, cy, cw, ch)
    return positions


def _layout_process(cx, cy, cw, ch, n):
    if n == 1:
        return [(cx, cy, cw, ch)]
    proc_h = int(ch * 0.4)
    rem_h = ch - proc_h - ROW_GAP
    positions = [(cx, cy, cw, proc_h)]
    if n > 1:
        eh = int((rem_h - ROW_GAP * (n - 2)) / (n - 1))
        for i in range(1, n):
            positions.append((cx, cy + proc_h + ROW_GAP + (i - 1) * (eh + ROW_GAP), cw, eh))
    return positions


def _layout_table(cx, cy, cw, ch, n):
    return _layout_full(cx, cy, cw, ch, n)


def _layout_mixed(cx, cy, cw, ch, n, elements=None):
    return _layout_full(cx, cy, cw, ch, n)


# ============================================================
# Custom area version (content renderer에서 사용)
# ============================================================

def calculate_positions_custom(
    layout: LayoutType,
    n_elements: int,
    n_cols: int = 1,
    elements: list | None = None,
    content_x=None,
    content_y=None,
    content_w=None,
    content_h=None,
) -> list[tuple[int, int, int, int]]:
    """커스텀 콘텐츠 영역으로 좌표를 계산한다."""
    cx = content_x if content_x is not None else CONTENT_X
    cy = content_y if content_y is not None else CONTENT_Y
    cw = content_w if content_w is not None else CONTENT_W
    ch = content_h if content_h is not None else CONTENT_H

    if layout == LayoutType.FULL:
        return _layout_full(cx, cy, cw, ch, n_elements)
    elif layout == LayoutType.COLUMNS:
        return _layout_columns(cx, cy, cw, ch, n_elements, n_cols, elements)
    elif layout == LayoutType.PROCESS:
        return _layout_process(cx, cy, cw, ch, n_elements)
    elif layout == LayoutType.TABLE:
        return _layout_table(cx, cy, cw, ch, n_elements)
    elif layout == LayoutType.MIXED:
        return _layout_mixed(cx, cy, cw, ch, n_elements, elements)
    return _layout_full(cx, cy, cw, ch, n_elements)
