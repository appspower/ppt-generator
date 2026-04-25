"""Layer 2 — needs_review 슬라이드를 30장 batch로 분할 + Agent prompt 생성."""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

ROOT = Path(__file__).resolve().parent.parent
auto_path = ROOT / "output" / "catalog" / "auto_labels_v1.json"
png_dir = ROOT / "output" / "catalog" / "all_pngs"
batches_dir = ROOT / "output" / "catalog" / "layer2_batches"
batches_dir.mkdir(parents=True, exist_ok=True)

with open(auto_path, "r", encoding="utf-8") as f:
    data = json.load(f)

review_slides = [r for r in data["labels"] if r["needs_review"]]
print(f"[layer2] {len(review_slides)} slides need review")

BATCH_SIZE = 30
batches = [review_slides[i:i + BATCH_SIZE] for i in range(0, len(review_slides), BATCH_SIZE)]
print(f"[layer2] {len(batches)} batches × ~{BATCH_SIZE}")

for bi, batch in enumerate(batches, start=1):
    batch_obj = {
        "batch_index": bi,
        "total_batches": len(batches),
        "slides": [
            {
                "slide_index": s["slide_index"],
                "png_path": str(png_dir / f"slide_{s['slide_index']:04d}.png"),
                "auto_macro": s["macro"],
                "auto_archetype": s["archetype"],
                "auto_narrative_role": s["narrative_role"],
                "auto_confidence": round(s["overall_confidence"], 2),
                "review_reason": s["review_reason"],
            }
            for s in batch
        ],
    }
    out = batches_dir / f"batch_{bi:02d}.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(batch_obj, f, ensure_ascii=False, indent=2)

print(f"[layer2] saved {len(batches)} batches -> {batches_dir}")
