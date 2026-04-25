"""Layer 2 결과 통합 — 8 batch JSON → auto_labels_v2.json + 검수 큐.

전략
----
- 자동 라벨 (Layer 1) + Agent 검수 결과 (Layer 2) 통합
- review_confidence < 0.6 또는 Agent가 unknown 라벨 → Layer 3 검수 큐
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
auto_v1_path = ROOT / "output" / "catalog" / "auto_labels_v1.json"
batches_dir = ROOT / "output" / "catalog" / "layer2_results"
output_path = ROOT / "output" / "catalog" / "auto_labels_v2.json"
queue_path = ROOT / "output" / "catalog" / "layer3_user_queue.json"


def merge():
    with open(auto_v1_path, "r", encoding="utf-8") as f:
        v1 = json.load(f)

    # v1 → dict by slide_index
    by_idx = {r["slide_index"]: r for r in v1["labels"]}

    # Layer 2 batches 읽기
    batch_files = sorted(batches_dir.glob("batch_*.json"))
    print(f"[merge] {len(batch_files)} batch files")

    layer2_count = 0
    user_queue: list[dict] = []
    for bf in batch_files:
        with open(bf, "r", encoding="utf-8") as f:
            batch = json.load(f)
        for r in batch.get("reviewed_slides", []):
            idx = r["slide_index"]
            v1_rec = by_idx.get(idx)
            if v1_rec is None:
                print(f"  [warn] slide {idx} not in v1")
                continue
            # v1 라벨을 Layer 2 결과로 교체
            v1_rec["macro"] = r.get("final_macro", v1_rec["macro"])
            v1_rec["archetype"] = r.get("final_archetype", v1_rec["archetype"])
            v1_rec["narrative_role"] = r.get("final_narrative_role", v1_rec["narrative_role"])
            v1_rec["overall_confidence"] = r.get("review_confidence", v1_rec["overall_confidence"])
            v1_rec["layer2_reviewed"] = True
            v1_rec["layer2_agreed_with_auto"] = r.get("agreed_with_auto", False)
            v1_rec["layer2_notes"] = r.get("notes", "")
            v1_rec["needs_review"] = v1_rec["overall_confidence"] < 0.7

            # Layer 3 user queue 조건
            if (v1_rec["overall_confidence"] < 0.6
                    or "unknown" in [v1_rec["macro"]] + v1_rec["archetype"]
                    or not v1_rec["layer2_agreed_with_auto"]):
                user_queue.append({
                    "slide_index": idx,
                    "macro": v1_rec["macro"],
                    "archetype": v1_rec["archetype"],
                    "narrative_role": v1_rec["narrative_role"],
                    "confidence": v1_rec["overall_confidence"],
                    "auto_v1_macro": v1_rec.get("macro"),  # 동의 안 한 경우 비교
                    "notes": v1_rec.get("layer2_notes", ""),
                })
            layer2_count += 1

    # 통계
    from collections import Counter
    macro_count = Counter(r["macro"] for r in v1["labels"])
    arch_count = Counter()
    for r in v1["labels"]:
        for a in r["archetype"]:
            arch_count[a] += 1
    needs_rev = sum(1 for r in v1["labels"] if r.get("needs_review", False))

    summary = {
        "total": len(v1["labels"]),
        "layer2_reviewed": layer2_count,
        "layer3_queue": len(user_queue),
        "macro_distribution": dict(macro_count),
        "archetype_top10": dict(arch_count.most_common(10)),
        "needs_review_count": needs_rev,
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump({"summary": summary, "labels": v1["labels"]},
                  f, ensure_ascii=False, indent=2, default=str)
    with open(queue_path, "w", encoding="utf-8") as f:
        json.dump({"summary": summary, "queue": user_queue},
                  f, ensure_ascii=False, indent=2)

    print(f"[merge] summary: {summary}")
    print(f"[merge] auto_labels_v2.json + layer3_user_queue.json saved")


if __name__ == "__main__":
    merge()
