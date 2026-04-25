"""Layer 1 — Phase 1B/1C 결과 통합 → multi-label 자동 라벨링.

전략
----
1. 1C `primary_archetype` + `types[]` (multi-label, confidence high/medium/low)
2. 1B `archetype` + `grid_cols/rows` + `largest_text_zone` + `sig`
3. 두 source를 통합 vote → L1 macro / L2 archetype / L3 narrative_role 산출
4. 신뢰도 0~1 환산: high=0.85, medium=0.6, low=0.4 (1C); 1B는 archetype 매핑 일치도 기반
5. overall_confidence < 0.7 → needs_review = True
"""
from __future__ import annotations

import json
from pathlib import Path

from .labels import ArchetypeLabel, Confidence, MacroLabel, SlideLabels
from .schemas import NarrativeRole


# 1C confidence 문자열 → 점수
CONF_SCORE = {"high": 0.85, "medium": 0.6, "low": 0.4}

# 1C 라벨 → L2 ArchetypeLabel 매핑
_1C_TO_L2 = {
    "table_native": ArchetypeLabel.TABLE_NATIVE,
    "matrix_2x2": ArchetypeLabel.MATRIX_2X2,
    "matrix_3x3": ArchetypeLabel.MATRIX_3X3,
    "cards_2col": ArchetypeLabel.CARDS_2COL,
    "cards_3col": ArchetypeLabel.CARDS_3COL,
    "cards_4col": ArchetypeLabel.CARDS_4COL,
    "cards_5col": ArchetypeLabel.CARDS_5PLUS,
    "cards_6col": ArchetypeLabel.CARDS_5PLUS,
    "cards_7col": ArchetypeLabel.CARDS_5PLUS,
    "cards_8col": ArchetypeLabel.CARDS_5PLUS,
    "orgchart": ArchetypeLabel.ORGCHART,
    "hub_spoke": ArchetypeLabel.HUB_SPOKE,
    "flowchart": ArchetypeLabel.FLOWCHART,
    "roadmap": ArchetypeLabel.ROADMAP,
    "timeline_h": ArchetypeLabel.TIMELINE_H,
    "gantt_like": ArchetypeLabel.GANTT,
    "swimlane": ArchetypeLabel.SWIMLANE,
    "venn": ArchetypeLabel.VENN,
    "chart_native": ArchetypeLabel.CHART_NATIVE,
    "cover_or_divider": ArchetypeLabel.COVER_DIVIDER,
    "picture_chart_like": ArchetypeLabel.CHART_NATIVE,  # 보통 차트 PNG
}

# 1B archetype → L2
_1B_TO_L2 = {
    "left_title_right_body": ArchetypeLabel.LEFT_TITLE_RIGHT_BODY,
    "dense_grid": ArchetypeLabel.DENSE_GRID,
    "3x3_matrix": ArchetypeLabel.MATRIX_3X3,
    "2x2_matrix": ArchetypeLabel.MATRIX_2X2,
    "2col_compare": ArchetypeLabel.CARDS_2COL,
    "3col_compare": ArchetypeLabel.CARDS_3COL,
    "4col_compare": ArchetypeLabel.CARDS_4COL,
    "5col_compare": ArchetypeLabel.CARDS_5PLUS,
    "6col_compare": ArchetypeLabel.CARDS_5PLUS,
    "7col_compare": ArchetypeLabel.CARDS_5PLUS,
    "vertical_list": ArchetypeLabel.VERTICAL_LIST,
    "single_block": ArchetypeLabel.SINGLE_BLOCK,
    "cover_or_divider": ArchetypeLabel.COVER_DIVIDER,
    "accent_strip_layout": ArchetypeLabel.LEFT_TITLE_RIGHT_BODY,
    "mixed": None,  # mixed는 L2 라벨 없음
}

# L2 → L1 매크로 매핑
_L2_TO_L1 = {
    ArchetypeLabel.TABLE_NATIVE: MacroLabel.TABLE,
    ArchetypeLabel.DENSE_GRID: MacroLabel.TABLE,
    ArchetypeLabel.MATRIX_2X2: MacroLabel.TABLE,
    ArchetypeLabel.MATRIX_3X3: MacroLabel.TABLE,
    ArchetypeLabel.MATRIX_NXN: MacroLabel.TABLE,
    ArchetypeLabel.CARDS_2COL: MacroLabel.CARD,
    ArchetypeLabel.CARDS_3COL: MacroLabel.CARD,
    ArchetypeLabel.CARDS_4COL: MacroLabel.CARD,
    ArchetypeLabel.CARDS_5PLUS: MacroLabel.CARD,
    ArchetypeLabel.VERTICAL_LIST: MacroLabel.CARD,
    ArchetypeLabel.ORGCHART: MacroLabel.DIAGRAM,
    ArchetypeLabel.HUB_SPOKE: MacroLabel.DIAGRAM,
    ArchetypeLabel.FLOWCHART: MacroLabel.DIAGRAM,
    ArchetypeLabel.ROADMAP: MacroLabel.DIAGRAM,
    ArchetypeLabel.TIMELINE_H: MacroLabel.DIAGRAM,
    ArchetypeLabel.GANTT: MacroLabel.DIAGRAM,
    ArchetypeLabel.SWIMLANE: MacroLabel.DIAGRAM,
    ArchetypeLabel.FUNNEL: MacroLabel.DIAGRAM,
    ArchetypeLabel.VENN: MacroLabel.DIAGRAM,
    ArchetypeLabel.CHART_NATIVE: MacroLabel.CHART,
    ArchetypeLabel.COVER_DIVIDER: MacroLabel.COVER,
    ArchetypeLabel.SINGLE_BLOCK: MacroLabel.COVER,
    ArchetypeLabel.LEFT_TITLE_RIGHT_BODY: MacroLabel.CARD,  # default; 검수에서 보정
}

# L2 → L3 narrative_role candidates (개략 — 검수에서 보정)
_L2_TO_L3 = {
    ArchetypeLabel.COVER_DIVIDER: [NarrativeRole.OPENING, NarrativeRole.DIVIDER, NarrativeRole.CLOSING],
    ArchetypeLabel.SINGLE_BLOCK: [NarrativeRole.DIVIDER, NarrativeRole.OPENING],
    ArchetypeLabel.TIMELINE_H: [NarrativeRole.ROADMAP],
    ArchetypeLabel.ROADMAP: [NarrativeRole.ROADMAP],
    ArchetypeLabel.GANTT: [NarrativeRole.ROADMAP],
    ArchetypeLabel.FLOWCHART: [NarrativeRole.ANALYSIS, NarrativeRole.RECOMMENDATION],
    ArchetypeLabel.ORGCHART: [NarrativeRole.ANALYSIS, NarrativeRole.RECOMMENDATION],
    ArchetypeLabel.HUB_SPOKE: [NarrativeRole.ANALYSIS, NarrativeRole.RECOMMENDATION],
    ArchetypeLabel.SWIMLANE: [NarrativeRole.RECOMMENDATION, NarrativeRole.ROADMAP],
    ArchetypeLabel.FUNNEL: [NarrativeRole.ANALYSIS, NarrativeRole.RECOMMENDATION],
    ArchetypeLabel.VENN: [NarrativeRole.ANALYSIS],
    ArchetypeLabel.MATRIX_2X2: [NarrativeRole.ANALYSIS, NarrativeRole.EVIDENCE],
    ArchetypeLabel.MATRIX_3X3: [NarrativeRole.ANALYSIS, NarrativeRole.EVIDENCE],
    ArchetypeLabel.TABLE_NATIVE: [NarrativeRole.EVIDENCE, NarrativeRole.ANALYSIS],
    ArchetypeLabel.DENSE_GRID: [NarrativeRole.EVIDENCE],
    ArchetypeLabel.CHART_NATIVE: [NarrativeRole.EVIDENCE],
    ArchetypeLabel.CARDS_2COL: [NarrativeRole.RECOMMENDATION, NarrativeRole.ANALYSIS],
    ArchetypeLabel.CARDS_3COL: [NarrativeRole.RECOMMENDATION, NarrativeRole.BENEFIT],
    ArchetypeLabel.CARDS_4COL: [NarrativeRole.RECOMMENDATION, NarrativeRole.BENEFIT],
    ArchetypeLabel.CARDS_5PLUS: [NarrativeRole.RECOMMENDATION],
    ArchetypeLabel.VERTICAL_LIST: [NarrativeRole.AGENDA, NarrativeRole.RECOMMENDATION],
    ArchetypeLabel.LEFT_TITLE_RIGHT_BODY: [NarrativeRole.ANALYSIS, NarrativeRole.RECOMMENDATION],
}


def auto_label_slide(b1: dict, c1: dict) -> SlideLabels:
    """1B + 1C 결과 통합 → SlideLabels."""
    slide_index = c1["slide_index"]

    # ---- L2 archetype 후보 통합 ----
    archetype_set: dict[ArchetypeLabel, Confidence] = {}

    # 1C multi-label
    for t in c1.get("types", []):
        l2 = _1C_TO_L2.get(t["type"])
        if l2 is None:
            continue
        score = CONF_SCORE.get(t["confidence"], 0.4)
        if l2 not in archetype_set or archetype_set[l2].score < score:
            archetype_set[l2] = Confidence(
                label=l2.value, score=score,
                source="1C_detector",
                reason=f"1C: {t['type']} ({t['confidence']})",
            )

    # 1B archetype
    b1_arch = b1.get("archetype", "")
    l2_b = _1B_TO_L2.get(b1_arch)
    if l2_b is not None:
        prev = archetype_set.get(l2_b)
        # 1B + 1C 일치하면 score 부스트
        new_score = 0.75 if prev else 0.6
        if prev and prev.label == l2_b.value:
            new_score = min(0.95, prev.score + 0.15)
        archetype_set[l2_b] = Confidence(
            label=l2_b.value, score=new_score,
            source="1B_grid",
            reason=f"1B: {b1_arch}" + (f" + {prev.reason}" if prev else ""),
        )

    # 상위 3개만 archetype list (multi-label 1~3)
    top3 = sorted(archetype_set.values(), key=lambda c: -c.score)[:3]
    archetype_list = [ArchetypeLabel(c.label) for c in top3]

    # 빈 경우 fallback
    if not archetype_list:
        archetype_list = [ArchetypeLabel.UNKNOWN]
        top3 = [Confidence(label="unknown", score=0.2, source="1C_detector",
                          reason="no detection match")]

    # ---- L1 macro (가장 강한 L2 의 매크로) ----
    primary = archetype_list[0]
    macro = _L2_TO_L1.get(primary, MacroLabel.UNKNOWN)
    macro_conf = top3[0].score

    # ---- L3 narrative_role (L2 기반 후보 1~2개) ----
    role_candidates: list[NarrativeRole] = []
    for arch in archetype_list:
        for r in _L2_TO_L3.get(arch, []):
            if r not in role_candidates:
                role_candidates.append(r)
    role_list = role_candidates[:2] if role_candidates else [NarrativeRole.UNKNOWN]
    role_confs = [
        Confidence(label=r.value, score=0.5, source="1C_detector",
                  reason="L2→L3 mapping (heuristic)")
        for r in role_list
    ]

    # ---- 전체 신뢰도 = top archetype score ----
    overall = top3[0].score
    needs_review = overall < 0.7
    review_reason = None
    if needs_review:
        review_reason = f"top archetype score={overall:.2f} < 0.7"

    return SlideLabels(
        slide_index=slide_index,
        macro=macro,
        archetype=archetype_list,
        narrative_role=role_list,
        macro_confidence=macro_conf,
        archetype_confidences=top3,
        role_confidences=role_confs,
        overall_confidence=overall,
        needs_review=needs_review,
        review_reason=review_reason,
    )


def run_auto_label(
    b1_path: Path,
    c1_path: Path,
    output_path: Path,
) -> dict:
    """전수 자동 라벨링."""
    with open(b1_path, "r", encoding="utf-8") as f:
        b1_data = json.load(f)["per_slide"]
    with open(c1_path, "r", encoding="utf-8") as f:
        c1_data = json.load(f)["per_slide"]

    by_idx_b = {x["slide_index"]: x for x in b1_data}
    by_idx_c = {x["slide_index"]: x for x in c1_data}

    results: list[SlideLabels] = []
    for idx in sorted(by_idx_c.keys()):
        b1 = by_idx_b.get(idx, {})
        c1 = by_idx_c[idx]
        sl = auto_label_slide(b1, c1)
        results.append(sl)

    # 통계
    from collections import Counter
    macro_count = Counter(r.macro.value for r in results)
    arch_count = Counter()
    for r in results:
        for a in r.archetype:
            arch_count[a.value] += 1
    role_count = Counter()
    for r in results:
        for rl in r.narrative_role:
            role_count[rl.value] += 1
    needs_rev = sum(1 for r in results if r.needs_review)

    summary = {
        "total": len(results),
        "macro_distribution": dict(macro_count),
        "archetype_distribution": dict(arch_count.most_common()),
        "narrative_role_distribution": dict(role_count.most_common()),
        "needs_review_count": needs_rev,
        "needs_review_pct": round(needs_rev / len(results) * 100, 1),
        "overall_confidence_p25": round(
            sorted([r.overall_confidence for r in results])[len(results)//4], 3
        ),
        "overall_confidence_median": round(
            sorted([r.overall_confidence for r in results])[len(results)//2], 3
        ),
        "overall_confidence_p75": round(
            sorted([r.overall_confidence for r in results])[3*len(results)//4], 3
        ),
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(
            {"summary": summary, "labels": [r.model_dump(mode="json") for r in results]},
            f, ensure_ascii=False, indent=2, default=str,
        )

    print(f"[auto_label] {len(results)} slides")
    print(f"  macro: {dict(macro_count)}")
    print(f"  needs_review: {needs_rev} ({summary['needs_review_pct']}%)")
    print(f"  conf p25/median/p75: {summary['overall_confidence_p25']}/{summary['overall_confidence_median']}/{summary['overall_confidence_p75']}")
    return summary


if __name__ == "__main__":
    ROOT = Path(__file__).resolve().parent.parent.parent
    b1 = ROOT / "output" / "catalog" / "phase1b_semantic_stats.json"
    c1 = ROOT / "output" / "catalog" / "phase1c_component_types.json"
    out = ROOT / "output" / "catalog" / "auto_labels_v1.json"
    run_auto_label(b1, c1, out)
