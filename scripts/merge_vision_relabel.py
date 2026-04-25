"""Phase A3 v5 — 8 batch vision 결과 통합 → final_labels_v2.json.

각 batch의 vision_narrative_role을 기존 final_labels.json에 병합.
- 신뢰도 >= 0.6: 기존 narrative_role 교체
- 신뢰도 < 0.6: 기존 + 추가 (병합)
- "decorative" 명시 시: narrative_role 비움
"""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BATCHES_DIR = ROOT / "output" / "catalog" / "vision_relabel_batches"
ORIG_LABELS = ROOT / "output" / "catalog" / "final_labels.json"
OUTPUT = ROOT / "output" / "catalog" / "final_labels_v2.json"


def main():
    print("=" * 70)
    print("Vision Relabel 통합 → final_labels_v2.json")
    print("=" * 70)

    # 원본 라벨
    with open(ORIG_LABELS, encoding="utf-8") as f:
        orig = json.load(f)
    by_slide = {l["slide_index"]: l for l in orig["labels"]}
    print(f"원본 labels: {len(by_slide)}")

    # 8 batch 결과 통합
    vision_results = {}
    for b in range(8):
        rp = BATCHES_DIR / f"batch_{b:02d}_result.json"
        with open(rp, encoding="utf-8") as f:
            data = json.load(f)
        for r in data["results"]:
            vision_results[r["slide_index"]] = r
    print(f"vision 검수: {len(vision_results)}장")

    # 변경 통계
    n_overridden = 0
    n_added = 0
    n_decorative = 0
    n_unchanged = 0
    role_changes = Counter()

    for sidx, vr in vision_results.items():
        if sidx not in by_slide:
            continue
        orig_l = by_slide[sidx]
        orig_roles = set(orig_l.get("narrative_role", []))
        vision_roles = set(vr.get("vision_narrative_role", []))
        conf = vr.get("confidence", 0)

        # decorative만 마크되면 narrative_role 비움
        if vision_roles == {"decorative"} or "decorative" in vision_roles and len(vision_roles) == 1:
            orig_l["narrative_role"] = []
            orig_l["narrative_role_source"] = "vision_decorative"
            orig_l["vision_confidence"] = conf
            n_decorative += 1
            for r in orig_roles:
                role_changes[f"-{r}"] += 1
            continue

        # decorative 제거
        vision_roles.discard("decorative")
        if not vision_roles:
            n_unchanged += 1
            continue

        # 고신뢰도 → 교체. 저신뢰도 → 병합
        if conf >= 0.6:
            new_roles = sorted(vision_roles)
            for r in (orig_roles - vision_roles):
                role_changes[f"-{r}"] += 1
            for r in (vision_roles - orig_roles):
                role_changes[f"+{r}"] += 1
            orig_l["narrative_role"] = new_roles
            orig_l["narrative_role_source"] = "vision_override"
            orig_l["vision_confidence"] = conf
            n_overridden += 1
        else:
            new_roles = sorted(orig_roles | vision_roles)
            for r in (vision_roles - orig_roles):
                role_changes[f"+{r}"] += 1
            orig_l["narrative_role"] = new_roles
            orig_l["narrative_role_source"] = "vision_merge"
            orig_l["vision_confidence"] = conf
            n_added += 1

    # 추가 메타: vision_reason
    for sidx, vr in vision_results.items():
        if sidx in by_slide:
            by_slide[sidx]["vision_reason"] = vr.get("reason", "")

    # 새 narrative_role 분포 출력
    new_role_c = Counter()
    for l in orig["labels"]:
        for r in l.get("narrative_role", []):
            new_role_c[r] += 1

    print(f"\n[통계]")
    print(f"  vision override (conf>=0.6): {n_overridden}")
    print(f"  vision merge (conf<0.6): {n_added}")
    print(f"  decorative 처리: {n_decorative}")
    print(f"  unchanged: {n_unchanged}")
    print()
    print("[role 변경 (+추가/-제거)]")
    for change, n in role_changes.most_common():
        print(f"  {change:>20}: {n}")
    print()
    print("[새 narrative_role 분포]")
    for r, n in new_role_c.most_common():
        print(f"  {r:>15}: {n}")

    # summary 업데이트
    orig["summary"]["narrative_role_distribution_v2"] = dict(new_role_c.most_common())
    orig["summary"]["vision_relabel_stats"] = {
        "checked": len(vision_results),
        "overridden": n_overridden,
        "merged": n_added,
        "decorativized": n_decorative,
        "unchanged": n_unchanged,
    }

    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(orig, f, ensure_ascii=False, indent=2)
    print(f"\n[saved] {OUTPUT}")


if __name__ == "__main__":
    main()
