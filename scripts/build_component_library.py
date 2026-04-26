"""N1-Lite 컴포넌트 라이브러리 빌드 — Phase 2.

[docs/N1_LITE_IMPLEMENTATION.md] §4 + §7 Phase 2 구현.
헌법: [docs/PROJECT_DIRECTION.md] §1 N1-Lite 백본의 자산 라이브러리 빌드.

입력
----
- `docs/references/_master_templates/PPT 템플릿.pptx` (1,251장 마스터)
- `output/catalog/final_labels_v2.json` (slide-level archetype/role)
- `output/catalog/paragraph_labels.json` (paragraph-level group info — group_signature, role, position_in_group, group_size)

출력
----
- `output/component_library/{family}_family.pptx` × 10 family
- `output/component_library/components_index.json`

알고리즘
--------
1. paragraph_labels를 (slide_idx, group_signature)로 그룹화 → 후보 그룹 풀 생성
2. selection_rules 적용 (사양 §4.3): min/max group_size, min_text_slots,
   exclude_archetypes(dense_grid/dense_table), exclude_decorative_only
3. role 분포로 family 분류 (chevron_label → chevron_family, card_header → card_family, ...)
4. 가족별 score top-N 선정 (기본 5/family)
5. 마스터에서 extract_group → family.pptx 빈 슬라이드에 insert_component
6. components_index.json 생성

실행: python scripts/build_component_library.py
"""

from __future__ import annotations

import json
import sys
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ppt_builder.template import TemplateEditor  # noqa: E402
from ppt_builder.template.component_ops import (  # noqa: E402
    create_blank_slide_with_master_theme,
    extract_group,
    insert_component,
)


MASTER_PATH = ROOT / "docs" / "references" / "_master_templates" / "PPT 템플릿.pptx"
LABELS_PATH = ROOT / "output" / "catalog" / "final_labels_v2.json"
PARAS_PATH = ROOT / "output" / "catalog" / "paragraph_labels.json"
LIBRARY_DIR = ROOT / "output" / "component_library"


# --- selection_rules (사양 §4.3) ----------------------------------------------

SELECTION_RULES = {
    "min_group_size": 2,
    "max_group_size": 12,
    "min_text_slots": 1,
    "exclude_archetypes": {"dense_grid", "dense_table"},
    "exclude_decorative_only": True,
}

MAX_PER_FAMILY = 5  # 사양: 파일당 평균 3~5 variants


# --- family 분류 + applicable_roles (사양 §4.2) -------------------------------

FAMILY_APPLICABLE_ROLES = {
    "chevron": ["roadmap", "recommendation", "complication"],
    "card": ["analysis", "recommendation", "complication", "benefit", "situation"],
    "timeline": ["roadmap"],
    "matrix": ["analysis", "complication", "risk"],
    "callout": ["situation", "complication", "risk"],
    "kpi": ["benefit", "evidence"],
    "icon_grid": ["recommendation", "benefit"],
    "table": ["evidence", "analysis"],
    "flow": ["recommendation", "analysis"],
    "divider": ["divider"],
}

# 모든 family 키 (없는 family도 빈 .pptx로 출력하지 않음)
ALL_FAMILIES = list(FAMILY_APPLICABLE_ROLES.keys())


def classify_family(role_dist: Counter, slide_archetypes: set[str]) -> str | None:
    """그룹의 역할 분포 + 슬라이드 archetype 힌트로 family 결정.

    우선순위 (특이성 높은 순):
      1. kpi_value 존재 → kpi
      2. callout_text 존재 → callout
      3. chevron_label dominant → chevron
      4. archetype roadmap/timeline_h → timeline
      5. archetype matrix/comparison → matrix
      6. archetype flowchart/swimlane/orgchart → flow
      7. archetype divider → divider
      8. card_header dominant + group_size>=3 → card
      9. archetype table_native → table
     10. card_header가 group_size 4~8 + archetype에 cards_*  → icon_grid 가능 (그림 비중 따로 검사 필요. v1: card로)
    매칭 실패 → None (skip).
    """
    if role_dist.get("kpi_value", 0) > 0:
        return "kpi"
    if role_dist.get("callout_text", 0) > 0:
        return "callout"
    if role_dist.get("chevron_label", 0) >= 1:
        return "chevron"
    if "roadmap" in slide_archetypes or "timeline_h" in slide_archetypes or "gantt" in slide_archetypes:
        return "timeline"
    if "matrix_2x2" in slide_archetypes or "comparison_table" in slide_archetypes:
        return "matrix"
    if "flowchart" in slide_archetypes or "swimlane" in slide_archetypes or "orgchart" in slide_archetypes:
        return "flow"
    if "divider" in slide_archetypes or "section_divider" in slide_archetypes:
        return "divider"
    if role_dist.get("card_header", 0) >= 2:
        return "card"
    if "table_native" in slide_archetypes:
        return "table"
    return None


# --- 후보 그룹 풀 생성 ----------------------------------------------------------

def collect_candidate_groups(
    paragraphs: list[dict],
    slide_label_map: dict[int, dict],
) -> list[dict]:
    """paragraph_labels에서 (slide_idx, group_signature)별로 그룹화 → 후보 메타.

    각 후보 dict:
      {slide_idx, group_signature, flat_idxs, role_counter, n_text_slots, n_decorative}
    """
    bucket: dict[tuple[int, str], list[dict]] = defaultdict(list)
    for p in paragraphs:
        sig = p.get("group_signature")
        if not sig:
            continue
        bucket[(p["slide_index"], sig)].append(p)

    candidates: list[dict] = []
    for (sidx, sig), members in bucket.items():
        # 같은 shape에서 N paragraph인 경우 flat_idx 중복 제거
        flat_idxs = sorted({p["flat_idx"] for p in members})
        # role 카운트는 paragraph 단위
        roles = Counter(p.get("role") for p in members if p.get("role"))
        n_text_slots = sum(
            1 for p in members
            if p.get("role") and p["role"] != "decorative"
            and (p.get("text_len", 0) > 0 or "~~" in (p.get("text") or ""))
        )
        n_decorative = roles.get("decorative", 0)
        candidates.append({
            "slide_idx": sidx,
            "group_signature": sig,
            "flat_idxs": flat_idxs,
            "n_members": len(flat_idxs),
            "role_counter": roles,
            "n_text_slots": n_text_slots,
            "n_decorative": n_decorative,
            "n_paragraphs": len(members),
        })
    return candidates


def passes_selection(cand: dict, slide_label: dict) -> tuple[bool, str]:
    """selection_rules 검사. 통과 (True, '') / 실패 (False, reason)."""
    n = cand["n_members"]
    if n < SELECTION_RULES["min_group_size"]:
        return False, f"too small (n={n})"
    if n > SELECTION_RULES["max_group_size"]:
        return False, f"too large (n={n})"
    if cand["n_text_slots"] < SELECTION_RULES["min_text_slots"]:
        return False, "no text slots"
    if SELECTION_RULES["exclude_decorative_only"] and (
        cand["role_counter"]
        and all(r == "decorative" for r in cand["role_counter"].keys())
    ):
        return False, "decorative-only"
    archetypes = set(slide_label.get("archetype", []))
    if archetypes & SELECTION_RULES["exclude_archetypes"]:
        return False, f"slide archetype excluded: {archetypes & SELECTION_RULES['exclude_archetypes']}"
    return True, ""


def score_candidate(cand: dict, slide_label: dict) -> float:
    """순위용 score. 높을수록 좋음.

    점수 요소:
      + text_slots 비율 (decorative 비중 낮을수록)
      + group_size 4~8이 가장 좋음 (단순/복잡 균형)
      + slide overall_confidence + layer2_reviewed 가산
      + role 단일성 (한 가지 role이 dominant면 가산)
    """
    score = 0.0
    n = max(cand["n_members"], 1)
    # text 비율 0~1
    text_ratio = cand["n_text_slots"] / max(cand["n_paragraphs"], 1)
    score += text_ratio * 30

    # group_size 적정성 (gauss-like around 6)
    if 4 <= n <= 8:
        score += 25
    elif 3 <= n <= 10:
        score += 15
    else:
        score += 5

    # slide confidence
    score += float(slide_label.get("overall_confidence", 0.5)) * 15
    if slide_label.get("layer2_reviewed"):
        score += 5
    if slide_label.get("layer2_agreed_with_auto"):
        score += 5

    # role 단일성
    role_counter = cand["role_counter"]
    if role_counter:
        top_count = role_counter.most_common(1)[0][1]
        total = sum(role_counter.values())
        score += (top_count / total) * 10

    return round(score, 3)


# --- 라이브러리 빌드 -----------------------------------------------------------

def build_library() -> int:
    if not MASTER_PATH.exists():
        print(f"[FAIL] master not found: {MASTER_PATH}")
        return 1
    if not LABELS_PATH.exists() or not PARAS_PATH.exists():
        print(f"[FAIL] catalog json not found")
        return 1

    LIBRARY_DIR.mkdir(parents=True, exist_ok=True)

    print("Loading catalogs...")
    labels_doc = json.loads(LABELS_PATH.read_text(encoding="utf-8"))
    paras_doc = json.loads(PARAS_PATH.read_text(encoding="utf-8"))
    slide_labels = {it["slide_index"]: it for it in labels_doc["labels"]}
    paragraphs = paras_doc["paragraphs"]
    print(f"  {len(slide_labels)} slide labels, {len(paragraphs)} paragraphs")

    print("Collecting candidate groups...")
    candidates = collect_candidate_groups(paragraphs, slide_labels)
    print(f"  raw candidates: {len(candidates)}")

    # Apply selection_rules + classify family
    accepted: list[dict] = []
    skip_reasons: Counter = Counter()
    for cand in candidates:
        slide_label = slide_labels.get(cand["slide_idx"], {})
        ok, reason = passes_selection(cand, slide_label)
        if not ok:
            skip_reasons[reason.split(" (")[0]] += 1
            continue
        archetypes = set(slide_label.get("archetype", []))
        family = classify_family(cand["role_counter"], archetypes)
        if family is None:
            skip_reasons["no family match"] += 1
            continue
        cand["family"] = family
        cand["score"] = score_candidate(cand, slide_label)
        cand["slide_archetypes"] = list(archetypes)
        accepted.append(cand)
    print(f"  accepted: {len(accepted)}")
    print(f"  skip reasons: {skip_reasons.most_common()}")

    # Per family, top-N
    by_family: dict[str, list[dict]] = defaultdict(list)
    for c in accepted:
        by_family[c["family"]].append(c)

    chosen: dict[str, list[dict]] = {}
    print()
    print("Per-family selection (top-N by score):")
    for family in ALL_FAMILIES:
        items = sorted(by_family.get(family, []), key=lambda c: -c["score"])
        if not items:
            print(f"  {family:12s}: 0 candidates (skip)")
            continue
        # 다양성: 같은 slide_idx에서 N개 안 뽑기
        seen_slides: set[int] = set()
        picked: list[dict] = []
        for it in items:
            if it["slide_idx"] in seen_slides:
                continue
            picked.append(it)
            seen_slides.add(it["slide_idx"])
            if len(picked) >= MAX_PER_FAMILY:
                break
        chosen[family] = picked
        print(
            f"  {family:12s}: {len(items):4d} cand → picked {len(picked)} "
            f"(scores {[c['score'] for c in picked]})"
        )

    # 마스터 한 번 열어두고 모든 extract 수행
    print()
    print("Opening master for extraction...")
    master_prs = Presentation(str(MASTER_PATH))

    components_index: list[dict] = []

    for family, items in chosen.items():
        family_path = LIBRARY_DIR / f"{family}_family.pptx"
        print(f"\n  Building {family_path.name} ({len(items)} components)...")
        # 같은 master에서 throwaway slide 1장 보존 + blank N장 추가
        editor = TemplateEditor(str(MASTER_PATH))
        editor.keep_slides([0])
        target_prs = editor.prs

        for i, item in enumerate(items):
            try:
                src_slide = master_prs.slides[item["slide_idx"]]
                comp = extract_group(src_slide, item["flat_idxs"])
                blank = create_blank_slide_with_master_theme(target_prs)
                insert_component(blank, comp)
            except Exception as e:
                print(f"    [SKIP] s{item['slide_idx']} sig={item['group_signature']}: {type(e).__name__}: {e}")
                continue

            # library_slide_index: throwaway가 0번이므로 +1
            library_slide_index = i + 1

            # slots metadata
            slots = []
            for p in paragraphs:
                if p["slide_index"] != item["slide_idx"]:
                    continue
                if p["flat_idx"] not in item["flat_idxs"]:
                    continue
                if not p.get("role") or p["role"] == "decorative":
                    continue
                slots.append({
                    "flat_idx": p["flat_idx"],
                    "paragraph_id": p["paragraph_id"],
                    "role": p["role"],
                    "max_chars": p.get("max_chars"),
                    "position_in_group": p.get("position_in_group"),
                })

            # bbox in pct
            sw = labels_doc["summary"].get("slide_w_emu") or 9906000
            sh = labels_doc["summary"].get("slide_h_emu") or 6858000
            # paragraph data has slide w/h
            sw = paras_doc["summary"].get("slide_w_emu", sw)
            sh = paras_doc["summary"].get("slide_h_emu", sh)
            bbox = comp.bbox_emu
            bbox_pct = {
                "left": round(bbox["left"] / sw, 4),
                "top": round(bbox["top"] / sh, 4),
                "width": round(bbox["width"] / sw, 4),
                "height": round(bbox["height"] / sh, 4),
            }

            comp_id = f"{family}_v{i + 1}_s{item['slide_idx']}"
            components_index.append({
                "component_id": comp_id,
                "family": family,
                "library_path": family_path.name,
                "library_slide_index": library_slide_index,
                "source": {
                    "master_slide_index": item["slide_idx"],
                    "group_indices": item["flat_idxs"],
                    "group_signature": item["group_signature"],
                    "slide_archetypes": item["slide_archetypes"],
                },
                "geometry": {
                    "bbox_emu": bbox,
                    "bbox_pct": bbox_pct,
                },
                "slots": slots,
                "applicable_roles": FAMILY_APPLICABLE_ROLES[family],
                "score": item["score"],
            })

        editor.save(family_path)
        editor.cleanup()
        size_mb = family_path.stat().st_size / 1024 / 1024
        print(f"    saved: {family_path.name} ({size_mb:.1f} MB)")

    # components_index.json
    index_doc = {
        "version": "1.0",
        "generated_at": str(date.today()),
        "selection_rules": {**SELECTION_RULES, "exclude_archetypes": sorted(SELECTION_RULES["exclude_archetypes"])},
        "max_per_family": MAX_PER_FAMILY,
        "summary": {
            "total_components": len(components_index),
            "families": dict(Counter(c["family"] for c in components_index)),
        },
        "components": components_index,
    }
    index_path = LIBRARY_DIR / "components_index.json"
    index_path.write_text(
        json.dumps(index_doc, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print()
    print(f"Index saved: {index_path}")
    print(f"Total components: {len(components_index)}")
    print(f"By family: {index_doc['summary']['families']}")
    return 0


if __name__ == "__main__":
    sys.exit(build_library())
