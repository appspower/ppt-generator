"""클러스터 다양성 시각 검증 — 샘플 슬라이드 PNG 추출.

목적
----
Stage 2 클러스터링 결과 (12 클러스터, largest=933)가 실제로 유사 슬라이드를
뭉친 것인지 눈으로 확인.

선정 로직
---------
1. 가장 큰 클러스터(id=1, 933장)에서 구조 시그니처 다양한 8장
   - leaf shape 수 percentile 기준: p5, p25, p50, p75, p90, p95, p99, max
2. 다른 11개 클러스터 대표 각 1장
합 ~19장 PNG → output/catalog/samples_png/
"""
from __future__ import annotations

import json
from pathlib import Path

import pythoncom
import win32com.client


def select_samples(meta_path: Path, labels_path: Path, n_biggest: int = 8) -> list[int]:
    """샘플 슬라이드 인덱스 선정 (0-based)."""
    with open(meta_path, "r", encoding="utf-8") as f:
        metas = json.load(f)
    with open(labels_path, "r", encoding="utf-8") as f:
        labels = json.load(f)["labels"]

    from collections import defaultdict
    buckets: dict[int, list[int]] = defaultdict(list)
    for i, lb in enumerate(labels):
        buckets[int(lb)].append(i)

    selected: list[tuple[int, str]] = []  # (slide_idx, reason)

    # 1. 가장 큰 클러스터 내에서 leaf shape count 다양성 기준 샘플
    biggest_id = max(
        (k for k in buckets.keys() if k != -1),
        key=lambda k: len(buckets[k]),
    )
    biggest = buckets[biggest_id]
    leafs = [(i, metas[i].get("shape_count_leaf", 0)) for i in biggest]
    leafs.sort(key=lambda x: x[1])
    n = len(leafs)
    # percentile 포지션
    percentiles = [0.05, 0.25, 0.50, 0.65, 0.80, 0.90, 0.97, 0.999]
    for p in percentiles[:n_biggest]:
        pos = min(int(p * n), n - 1)
        slide_idx, leaf = leafs[pos]
        selected.append((slide_idx, f"cluster{biggest_id}_p{int(p*100)}_leaf{leaf}"))

    # 2. 다른 클러스터 대표
    for lb, members in sorted(buckets.items(), key=lambda kv: -len(kv[1])):
        if lb == biggest_id:
            continue
        if lb == -1:
            # noise에서 2장만
            for idx in members[:2]:
                selected.append((idx, f"noise_leaf{metas[idx].get('shape_count_leaf',0)}"))
            continue
        # 중앙 멤버
        mid = members[len(members) // 2]
        selected.append((mid, f"cluster{lb}_size{len(members)}_leaf{metas[mid].get('shape_count_leaf',0)}"))

    return selected


def export_selected_pngs(
    pptx_path: Path,
    samples: list[tuple[int, str]],
    output_dir: Path,
    width: int = 1568,
    height: int = 1176,
) -> list[Path]:
    """PowerPoint COM으로 선택된 슬라이드만 PNG export."""
    pptx_path = Path(pptx_path).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)

    pythoncom.CoInitialize()
    powerpoint = None
    presentation = None
    out_paths: list[Path] = []

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(
            str(pptx_path),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )

        total = presentation.Slides.Count
        print(f"[open] {total} slides in pptx")

        # manifest
        manifest = []

        for slide_idx_0, reason in samples:
            com_idx = slide_idx_0 + 1  # COM은 1-based
            if com_idx > total:
                print(f"  [skip] idx {slide_idx_0} > total")
                continue
            slide = presentation.Slides(com_idx)
            fname = f"slide_{slide_idx_0:04d}__{reason}.png"
            out_path = output_dir / fname
            try:
                slide.Export(str(out_path), "PNG", width, height)
                out_paths.append(out_path)
                manifest.append({"slide_index": slide_idx_0, "reason": reason, "file": fname})
                print(f"  [ok] #{slide_idx_0:>4} -> {fname}")
            except Exception as e:
                print(f"  [err] #{slide_idx_0}: {e}")

        # manifest 저장
        with open(output_dir / "_manifest.json", "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return out_paths


if __name__ == "__main__":
    ROOT = Path(__file__).resolve().parent.parent
    meta = ROOT / "output" / "catalog" / "slide_meta.json"
    labels = ROOT / "output" / "catalog" / "cluster_labels.json"
    pptx = ROOT / "docs" / "references" / "_master_templates" / "PPT 템플릿.pptx"
    out = ROOT / "output" / "catalog" / "samples_png"

    samples = select_samples(meta, labels)
    print(f"[select] {len(samples)} samples:")
    for s in samples:
        print(f"  #{s[0]:>4}: {s[1]}")

    paths = export_selected_pngs(pptx, samples, out)
    print(f"\n[done] {len(paths)} PNGs -> {out}")
