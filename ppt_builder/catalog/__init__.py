"""Phase A2 — 슬라이드 카탈로그 + 덱 스켈레톤 구축 모듈.

Track 1: 슬라이드 카탈로그 (slide_meta → cluster → tag → schema → filter)
Track 2: 덱 스켈레톤 (boundary → role sequence → LCS skeleton)
Track 3: Deck coherence 검증 (ppt_builder/evaluate_deck.py 별도)

설계 근거: memory/project_phase_a2_plan.md
"""
from .schemas import (
    DeckBoundary,
    DeckOutline,
    DeckSkeleton,
    NarrativeRole,
    SlideMeta,
    SlotSchema,
    TagAxes,
)
from .skeletons import SKELETONS, get_skeleton, recommend_skeleton

__all__ = [
    "SlideMeta",
    "TagAxes",
    "NarrativeRole",
    "SlotSchema",
    "DeckBoundary",
    "DeckOutline",
    "DeckSkeleton",
    "SKELETONS",
    "get_skeleton",
    "recommend_skeleton",
]
