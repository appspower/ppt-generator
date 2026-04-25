"""Phase A3 v5 — Vision 재라벨링 후보 슬라이드 추출.

전략: 결손 role (closing/situation/risk/complication/benefit/divider) 후보 발굴
  1. archetype에 cover_divider 있는 슬라이드 (모두)
  2. fillable 슬롯 <= 10 (단순 메시지 = cover/closing/divider 후보)
  3. 8 batch로 분할 → Agent 병렬 vision 검수

출력: output/catalog/vision_relabel_batches/batch_NN.json (8개)
       output/catalog/vision_relabel_manifest.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ppt_builder.catalog.paragraph_query import ParagraphStore


CATALOG_PATH = ROOT / "output" / "catalog" / "final_labels.json"
PNG_DIR = ROOT / "output" / "catalog" / "all_pngs"
OUT_DIR = ROOT / "output" / "catalog" / "vision_relabel_batches"


def main():
    print("=" * 70)
    print("Vision Relabel 후보 추출")
    print("=" * 70)

    with open(CATALOG_PATH, encoding="utf-8") as f:
        labels_data = json.load(f)
    by_slide = {l["slide_index"]: l for l in labels_data["labels"]}

    store = ParagraphStore.load()

    # 후보 추출
    cover_arch_slides = set(
        l["slide_index"] for l in labels_data["labels"]
        if "cover_divider" in l.get("archetype", [])
    )

    sparse_slides = set()
    for sidx in range(1251):
        cap = store.slot_capacity(sidx)
        if sum(cap.values()) <= 10:
            sparse_slides.add(sidx)

    candidates = sorted(cover_arch_slides | sparse_slides)
    print(f"cover_arch: {len(cover_arch_slides)}, sparse(<=10): {len(sparse_slides)}")
    print(f"unique candidates: {len(candidates)}")

    # 8 batch
    n_batches = 8
    batch_size = (len(candidates) + n_batches - 1) // n_batches
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    manifest = {"n_batches": n_batches, "n_total": len(candidates), "batches": []}

    for b in range(n_batches):
        batch_slides = candidates[b * batch_size : (b + 1) * batch_size]
        batch_data = []
        for sidx in batch_slides:
            l = by_slide.get(sidx, {})
            cap = store.slot_capacity(sidx)
            png_path = PNG_DIR / f"slide_{sidx:04d}.png"
            batch_data.append({
                "slide_index": sidx,
                "png_path": str(png_path),
                "current_macro": l.get("macro"),
                "current_archetype": l.get("archetype", []),
                "current_narrative_role": l.get("narrative_role", []),
                "fillable_capacity": dict(cap),
                "fillable_total": sum(cap.values()),
            })

        batch_path = OUT_DIR / f"batch_{b:02d}.json"
        with open(batch_path, "w", encoding="utf-8") as f:
            json.dump({
                "batch_id": b,
                "n_slides": len(batch_data),
                "slides": batch_data,
            }, f, ensure_ascii=False, indent=2)
        manifest["batches"].append({
            "batch_id": b,
            "n_slides": len(batch_data),
            "path": str(batch_path),
        })
        print(f"  batch {b}: {len(batch_data)} slides → {batch_path.name}")

    manifest_path = OUT_DIR / "manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f"\nmanifest: {manifest_path}")


if __name__ == "__main__":
    main()
