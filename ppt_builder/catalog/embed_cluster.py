"""Track 1 Stage 2 — BGE-M3 임베딩 + HDBSCAN 클러스터링.

입력: `output/catalog/slide_meta.json` (1,251 entries)
출력: `output/catalog/clusters.json` (60~120 클러스터 예상)

설계 근거
---------
- BGE-M3: 한영 SOTA 임베딩 (2024), structure+text 통합 기반
- HDBSCAN: 밀도 기반, k 지정 불필요, noise 분리 — 잡음 슬라이드 자동 분리
- 특징 벡터: BGE-M3(text_total + layout + structure_sig) ⊕ 수치 feature 정규화

비고
----
- 첫 실행 시 BGE-M3 모델 다운로드 (~2GB, 수 분)
- CPU에서 1,251장 임베딩 ~10-20분, GPU 있으면 훨씬 빠름
- 한글 plaeholder(`~~`)는 의미 없음 → 제거 후 임베딩
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import numpy as np

from .schemas import SlideMeta


BGE_MODEL = "BAAI/bge-m3"


def _clean_text(text: str) -> str:
    """임베딩 전 텍스트 정제: `~~` 제거, 공백 정규화, 길이 cap."""
    if not text:
        return ""
    t = text.replace("~~", " ")
    t = re.sub(r"\s+", " ", t).strip()
    return t[:2000]  # BGE-M3 8192 token까지 가능하나 비용 고려


def _compose_slide_text(m: SlideMeta) -> str:
    """한 슬라이드의 임베딩 입력 문자열 구성.

    layout + title + structure_sig + text 조합. 구조 시그니처를 명시적으로 포함해
    "같은 텍스트/다른 레이아웃"을 구분 가능하게 함.
    """
    parts: list[str] = []
    if m.layout_name:
        parts.append(f"[layout: {m.layout_name}]")
    if m.title_text:
        parts.append(f"[title: {m.title_text}]")
    if m.structure_sig:
        parts.append(f"[structure: {m.structure_sig}]")
    cleaned = _clean_text(m.text_total)
    if cleaned:
        parts.append(cleaned)
    return "\n".join(parts) if parts else f"[empty slide #{m.slide_index}]"


def embed_slides(metas: list[SlideMeta], batch_size: int = 32) -> np.ndarray:
    """BGE-M3로 슬라이드 전체를 임베딩.

    Returns
    -------
    np.ndarray of shape (N, 1024), float32, L2-normalized.
    """
    from sentence_transformers import SentenceTransformer

    print(f"[embed] loading {BGE_MODEL} ...")
    model = SentenceTransformer(BGE_MODEL)

    texts = [_compose_slide_text(m) for m in metas]
    print(f"[embed] encoding {len(texts)} slides, batch={batch_size} ...")
    embs = model.encode(
        texts,
        batch_size=batch_size,
        show_progress_bar=True,
        normalize_embeddings=True,
        convert_to_numpy=True,
    )
    print(f"[embed] done, shape={embs.shape}, dtype={embs.dtype}")
    return embs.astype(np.float32)


def cluster_embeddings(
    embs: np.ndarray,
    min_cluster_size: int = 5,
    min_samples: int | None = None,
) -> tuple[np.ndarray, dict]:
    """HDBSCAN (cosine → euclidean via normalized vectors) 로 클러스터링.

    Returns
    -------
    labels : np.ndarray of shape (N,), -1 = noise
    info   : dict (클러스터 수, noise 수, 평균 크기 등)
    """
    import hdbscan

    # normalized vectors면 euclidean == 2*(1-cos) 이므로 그대로 사용
    clusterer = hdbscan.HDBSCAN(
        min_cluster_size=min_cluster_size,
        min_samples=min_samples,
        metric="euclidean",
        cluster_selection_method="eom",
        prediction_data=False,
    )
    labels = clusterer.fit_predict(embs)

    unique, counts = np.unique(labels, return_counts=True)
    noise_count = int(counts[unique == -1].sum()) if (unique == -1).any() else 0
    cluster_count = int((unique != -1).sum())
    non_noise = counts[unique != -1]
    info = {
        "total": int(len(labels)),
        "cluster_count": cluster_count,
        "noise_count": noise_count,
        "avg_cluster_size": float(non_noise.mean()) if cluster_count > 0 else 0.0,
        "min_cluster_size_config": min_cluster_size,
        "largest_cluster": int(non_noise.max()) if cluster_count > 0 else 0,
    }
    return labels, info


def save_clusters(
    metas: list[SlideMeta],
    labels: np.ndarray,
    embs: np.ndarray,
    output_dir: Path,
) -> None:
    """clusters.json + embeddings.npy 저장."""
    output_dir.mkdir(parents=True, exist_ok=True)

    # clusters.json — 클러스터별 멤버 리스트 + 대표
    from collections import defaultdict
    buckets: dict[int, list[int]] = defaultdict(list)
    for i, lb in enumerate(labels):
        buckets[int(lb)].append(i)

    clusters = []
    for lb, members in sorted(buckets.items(), key=lambda kv: (kv[0] == -1, -len(kv[1]))):
        # 중심에 가장 가까운 멤버를 대표로
        if lb == -1:
            representative = members[0]
        else:
            member_embs = embs[members]
            centroid = member_embs.mean(axis=0)
            sims = member_embs @ centroid
            representative = int(members[int(sims.argmax())])
        clusters.append({
            "cluster_id": int(lb),
            "size": len(members),
            "representative_slide_index": int(representative),
            "member_slide_indices": [int(i) for i in members],
            "representative_layout": metas[representative].layout_name,
            "representative_title": metas[representative].title_text,
            "representative_structure": metas[representative].structure_sig,
        })

    with open(output_dir / "clusters.json", "w", encoding="utf-8") as f:
        json.dump(clusters, f, ensure_ascii=False, indent=2)

    # embeddings.npy — 재활용 용
    np.save(output_dir / "embeddings.npy", embs)

    # 라벨 배열도 별도 저장 (JSON 친화)
    with open(output_dir / "cluster_labels.json", "w", encoding="utf-8") as f:
        json.dump(
            {"labels": [int(l) for l in labels]},
            f, ensure_ascii=False, indent=2,
        )

    print(f"[save] clusters.json + embeddings.npy + cluster_labels.json -> {output_dir}")


def run_embed_cluster(
    meta_path: Path,
    output_dir: Path,
    min_cluster_size: int = 5,
) -> tuple[list[SlideMeta], np.ndarray, np.ndarray, dict]:
    """메타 로드 → 임베딩 → 클러스터 → 저장."""
    with open(meta_path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    metas = [SlideMeta.model_validate(r) for r in raw]
    print(f"[load] {len(metas)} slide metas from {meta_path}")

    embs = embed_slides(metas)
    labels, info = cluster_embeddings(embs, min_cluster_size=min_cluster_size)
    print(f"[cluster] {info}")

    save_clusters(metas, labels, embs, output_dir)
    # summary 저장
    with open(output_dir / "cluster_summary.json", "w", encoding="utf-8") as f:
        json.dump(info, f, ensure_ascii=False, indent=2)
    return metas, embs, labels, info


if __name__ == "__main__":
    import argparse
    ROOT = Path(__file__).resolve().parent.parent.parent
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--meta",
        default=str(ROOT / "output" / "catalog" / "slide_meta.json"),
    )
    parser.add_argument(
        "--out",
        default=str(ROOT / "output" / "catalog"),
    )
    parser.add_argument("--min-cluster-size", type=int, default=5)
    args = parser.parse_args()
    run_embed_cluster(Path(args.meta), Path(args.out), args.min_cluster_size)
