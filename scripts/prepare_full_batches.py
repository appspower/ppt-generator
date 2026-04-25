"""1,013장 confidence≥0.7 슬라이드를 30장 batch로 분할 (전수 검증용)."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
v2_path = ROOT / "output" / "catalog" / "auto_labels_v2.json"
png_dir = ROOT / "output" / "catalog" / "all_pngs"
batches_dir = ROOT / "output" / "catalog" / "full_batches"
batches_dir.mkdir(parents=True, exist_ok=True)

with open(v2_path, "r", encoding="utf-8") as f:
    data = json.load(f)

# Layer 2 검수 안 된 슬라이드 (confidence>=0.7로 자동 통과)
unreviewed = [r for r in data["labels"] if not r.get("layer2_reviewed", False)]
print(f"[full] {len(unreviewed)} slides to verify (confidence>=0.7)")

BATCH_SIZE = 30
batches = [unreviewed[i:i+BATCH_SIZE] for i in range(0, len(unreviewed), BATCH_SIZE)]
print(f"[full] {len(batches)} batches × ~{BATCH_SIZE}")

for bi, batch in enumerate(batches, start=1):
    bnum = bi + 8  # 이미 batch_01~08 사용했으므로 09부터
    batch_obj = {
        "batch_index": bnum,
        "total_batches": len(batches) + 8,
        "purpose": "full_verification",
        "slides": [
            {
                "slide_index": s["slide_index"],
                "png_path": str(png_dir / f"slide_{s['slide_index']:04d}.png"),
                "auto_macro": s["macro"],
                "auto_archetype": s["archetype"],
                "auto_narrative_role": s["narrative_role"],
                "auto_confidence": round(s["overall_confidence"], 2),
            }
            for s in batch
        ],
    }
    out = batches_dir / f"batch_{bnum:02d}.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(batch_obj, f, ensure_ascii=False, indent=2)

print(f"[full] batch_09 ~ batch_{8+len(batches):02d} saved -> {batches_dir}")
