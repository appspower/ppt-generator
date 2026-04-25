"""Phase A3 Step 2 — paragraph-level catalog 조회 API.

paragraph_labels.json을 메모리에 로드하여 슬라이드별 슬롯 매핑을 제공한다.

주요 사용
--------
>>> store = ParagraphStore.load()
>>> slots = store.fillable_slots(slide_index=44)
>>> # 반환: {"title": [...], "chevron_label": [...], "card_header": [...], ...}

>>> mapping = store.match_content(
...     slide_index=44,
...     content_by_role={"title": "...", "chevron_label": ["a","b","c","d","e"]},
... )
>>> # 반환: [(div_id, paragraph_id, text)] 편집 명령 리스트
"""
from __future__ import annotations

import json
from collections import defaultdict
from dataclasses import dataclass
from functools import cached_property
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parents[2]
DEFAULT_LABELS_PATH = ROOT / "output" / "catalog" / "paragraph_labels.json"


# 사용자-fillable 의미 슬롯 (즉, 실제 컨텐츠를 채울 슬롯)
FILLABLE_ROLES = {
    "title",
    "subtitle",
    "kicker",
    "table_header",
    "table_cell",
    "chevron_label",
    "card_header",
    "card_body",
    "callout_text",
    "kpi_value",
}

# fill 우선순위 (디자인 핵심 슬롯 우선)
ROLE_PRIORITY = [
    "title",
    "subtitle",
    "table_header",
    "chevron_label",
    "card_header",
    "callout_text",
    "kpi_value",
    "card_body",
    "table_cell",
]


@dataclass
class ParagraphSlot:
    slide_index: int
    flat_idx: int
    paragraph_id: int
    role: str
    role_confidence: float
    role_source: str
    text_original: str
    max_chars: int | None
    position_in_group: int | None
    group_size: int | None
    group_signature: str | None
    shape_kind: str
    placeholder_type: str | None
    bbox_left_pct: float
    bbox_top_pct: float
    table_row: int | None
    table_col: int | None

    @classmethod
    def from_dict(cls, d: dict) -> "ParagraphSlot":
        return cls(
            slide_index=d["slide_index"],
            flat_idx=d["flat_idx"],
            paragraph_id=d["paragraph_id"],
            role=d.get("role", "decorative"),
            role_confidence=d.get("role_confidence", 0),
            role_source=d.get("role_source", ""),
            text_original=d.get("text", ""),
            max_chars=d.get("max_chars"),
            position_in_group=d.get("position_in_group"),
            group_size=d.get("group_size"),
            group_signature=d.get("group_signature"),
            shape_kind=d.get("shape_kind", ""),
            placeholder_type=d.get("placeholder_type"),
            bbox_left_pct=d["bbox"]["left_pct"],
            bbox_top_pct=d["bbox"]["top_pct"],
            table_row=d.get("table_row"),
            table_col=d.get("table_col"),
        )


class ParagraphStore:
    """paragraph_labels.json 로드 + 슬라이드별 슬롯 매핑."""

    def __init__(self, paragraphs: list[ParagraphSlot]):
        self._paragraphs = paragraphs
        self._by_slide: dict[int, list[ParagraphSlot]] = defaultdict(list)
        for p in paragraphs:
            self._by_slide[p.slide_index].append(p)

    @classmethod
    def load(cls, path: Path | str = DEFAULT_LABELS_PATH) -> "ParagraphStore":
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        items = [ParagraphSlot.from_dict(p) for p in data["paragraphs"]]
        return cls(items)

    @property
    def n_slides(self) -> int:
        return len(self._by_slide)

    def slots(self, slide_index: int) -> list[ParagraphSlot]:
        return list(self._by_slide.get(slide_index, []))

    def fillable_slots(self, slide_index: int) -> dict[str, list[ParagraphSlot]]:
        """슬라이드의 fillable 슬롯들을 role별로 묶어 반환.
        position_in_group이 있으면 그 순서대로 정렬, 없으면 bbox top→left 순.
        """
        slots = self.slots(slide_index)
        out: dict[str, list[ParagraphSlot]] = defaultdict(list)
        for s in slots:
            if s.role not in FILLABLE_ROLES:
                continue
            out[s.role].append(s)

        # 정렬
        for role, items in out.items():
            if role == "table_header":
                items.sort(key=lambda s: (s.table_row or 0, s.table_col or 0))
            elif role == "table_cell":
                items.sort(key=lambda s: (s.table_row or 0, s.table_col or 0))
            elif role in {"chevron_label", "card_header", "card_body"}:
                # group position 우선, 없으면 bbox 좌→우→상→하
                items.sort(key=lambda s: (
                    s.position_in_group if s.position_in_group is not None else 9999,
                    s.bbox_top_pct,
                    s.bbox_left_pct,
                    s.paragraph_id,
                ))
            else:
                items.sort(key=lambda s: (s.bbox_top_pct, s.bbox_left_pct,
                                          s.paragraph_id))
        return dict(out)

    def slot_capacity(self, slide_index: int) -> dict[str, int]:
        """role별 슬롯 갯수."""
        f = self.fillable_slots(slide_index)
        return {r: len(v) for r, v in f.items()}

    # -----------------------------------------------------------------
    # Content matching
    # -----------------------------------------------------------------

    def match_content(
        self,
        slide_index: int,
        content_by_role: dict[str, list[str] | str],
        truncate_overflow: bool = True,
    ) -> list[dict]:
        """컨텐츠를 슬롯에 매핑.

        Parameters
        ----------
        content_by_role : dict
            role 이름 → 문자열 또는 문자열 리스트
        truncate_overflow : bool
            True면 max_chars 초과 시 잘라냄

        Returns
        -------
        list of dicts: [
            {"flat_idx": int, "paragraph_id": int, "text": str,
             "role": str, "matched_pos": int, "max_chars": int|None,
             "overflow": bool}, ...
        ]
        """
        fillable = self.fillable_slots(slide_index)
        edits: list[dict] = []

        for role, content in content_by_role.items():
            if role not in fillable:
                continue
            slots = fillable[role]
            if isinstance(content, str):
                items = [content]
            else:
                items = list(content)

            for pos, text in enumerate(items):
                if pos >= len(slots):
                    break
                slot = slots[pos]
                overflow = bool(
                    slot.max_chars and len(text) > slot.max_chars
                )
                final_text = text
                if overflow and truncate_overflow:
                    final_text = text[: slot.max_chars - 1] + "…"
                edits.append({
                    "flat_idx": slot.flat_idx,
                    "paragraph_id": slot.paragraph_id,
                    "text": final_text,
                    "text_original": slot.text_original,
                    "role": role,
                    "matched_pos": pos,
                    "max_chars": slot.max_chars,
                    "original_overflow": overflow,
                })
        return edits

    def fill_stats(self, slide_index: int,
                   content_by_role: dict[str, list[str] | str]) -> dict:
        """슬라이드의 fill 통계 (capacity vs content)."""
        capacity = self.slot_capacity(slide_index)
        requested = {
            r: (1 if isinstance(c, str) else len(c))
            for r, c in content_by_role.items()
        }
        n_fillable_total = sum(capacity.values())
        n_filled = 0
        n_short = 0  # capacity > requested
        n_overflow = 0  # requested > capacity
        for r, cap in capacity.items():
            req = requested.get(r, 0)
            n_filled += min(cap, req)
            if req < cap:
                n_short += cap - req
            elif req > cap:
                n_overflow += req - cap
        return {
            "capacity": capacity,
            "requested": requested,
            "n_fillable": n_fillable_total,
            "n_filled": n_filled,
            "n_short": n_short,
            "n_overflow": n_overflow,
            "fill_ratio": (n_filled / n_fillable_total) if n_fillable_total else 0,
        }
