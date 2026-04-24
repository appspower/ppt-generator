"""Track 1 Stage 2a cluster + verify — OOXML shape feature 기반 HDBSCAN.

파이프라인
----------
1. shape_features.npy 로드 (68-dim per slide)
2. HDBSCAN(min_cluster_size=3, min_samples=2) 로 클러스터링
3. noise(-1) 는 k-NN assignment (arXiv 2506.12116) 로 재할당
4. 검증 metrics:
   - silhouette score (전체)
   - structure_sig 교차일치율 (클러스터 내부 동일 sig 비율)
   - cluster size 분포
"""
from __future__ import annotations

import json
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np

from .schemas import SlideMeta


def cluster_hdbscan(
    features: np.ndarray,
    min_cluster_size: int = 3,
    min_samples: int = 2,
    reassign_noise: bool = True,
) -> tuple[np.ndarray, dict]:
    """HDBSCAN + k-NN noise reassignment."""
    import hdbscan

    clusterer = hdbscan.HDBSCAN(
        min_cluster_size=min_cluster_size,
        min_samples=min_samples,
        metric="euclidean",
        cluster_selection_method="eom",
        prediction_data=False,
    )
    labels = clusterer.fit_predict(features)

    noise_count_initial = int((labels == -1).sum())
    if reassign_noise and noise_count_initial > 0:
        # k=3 NN on clustered points
        from sklearn.neighbors import NearestNeighbors

        clustered_mask = labels != -1
        if clustered_mask.sum() >= 3:
            nn = NearestNeighbors(n_neighbors=3).fit(features[clustered_mask])
            noise_idx = np.where(~clustered_mask)[0]
            _, nbrs = nn.kneighbors(features[noise_idx])
            clustered_labels = labels[clustered_mask]
            for i, ni in enumerate(noise_idx):
                # majority vote of 3 nearest
                votes = Counter(clustered_labels[nbrs[i]])
                labels[ni] = votes.most_common(1)[0][0]

    # stats
    unique, counts = np.unique(labels, return_counts=True)
    cluster_ids = unique[unique != -1]
    sizes = counts[unique != -1]
    info = {
        "total": int(len(labels)),
        "cluster_count": int(len(cluster_ids)),
        "noise_count_initial": noise_count_initial,
        "noise_count_final": int((labels == -1).sum()),
        "avg_cluster_size": float(sizes.mean()) if len(sizes) > 0 else 0.0,
        "median_cluster_size": float(np.median(sizes)) if len(sizes) > 0 else 0.0,
        "min_cluster_size": int(sizes.min()) if len(sizes) > 0 else 0,
        "max_cluster_size": int(sizes.max()) if len(sizes) > 0 else 0,
        "min_cluster_size_config": min_cluster_size,
        "min_samples_config": min_samples,
    }
    return labels, info


def evaluate_clusters(
    features: np.ndarray,
    labels: np.ndarray,
    metas: list[SlideMeta],
) -> dict:
    """클러스터 품질 평가 지표 3종."""
    out: dict = {}

    # 1. Silhouette (sklearn)
    try:
        from sklearn.metrics import silhouette_score
        mask = labels != -1
        if mask.sum() >= 10 and len(set(labels[mask])) >= 2:
            out["silhouette"] = float(silhouette_score(features[mask], labels[mask]))
        else:
            out["silhouette"] = None
    except Exception as e:
        out["silhouette_err"] = str(e)

    # 2. structure_sig purity — 클러스터 내 가장 빈도 높은 sig 비율 평균
    buckets: dict[int, list[str]] = defaultdict(list)
    for i, lb in enumerate(labels):
        buckets[int(lb)].append(metas[i].structure_sig or "empty")

    purities = []
    for lb, sigs in buckets.items():
        if lb == -1 or len(sigs) < 2:
            continue
        top = Counter(sigs).most_common(1)[0][1]
        purities.append(top / len(sigs))
    out["structure_sig_purity_mean"] = float(np.mean(purities)) if purities else 0.0
    out["structure_sig_purity_p25"] = float(np.percentile(purities, 25)) if purities else 0.0
    out["structure_sig_purity_p75"] = float(np.percentile(purities, 75)) if purities else 0.0
    out["clusters_evaluated"] = len(purities)

    # 3. Cluster size distribution
    sizes = [len(sigs) for lb, sigs in buckets.items() if lb != -1]
    out["size_distribution"] = {
        "count": len(sizes),
        "min": int(min(sizes)) if sizes else 0,
        "max": int(max(sizes)) if sizes else 0,
        "median": float(np.median(sizes)) if sizes else 0,
        "p90": float(np.percentile(sizes, 90)) if sizes else 0,
    }
    return out


def save_clusters(
    metas: list[SlideMeta],
    features: np.ndarray,
    labels: np.ndarray,
    output_dir: Path,
    tag: str = "v2",
) -> None:
    """cluster_<tag>.json + labels + info 저장."""
    output_dir.mkdir(parents=True, exist_ok=True)

    buckets: dict[int, list[int]] = defaultdict(list)
    for i, lb in enumerate(labels):
        buckets[int(lb)].append(i)

    clusters = []
    for lb, members in sorted(buckets.items(), key=lambda kv: (kv[0] == -1, -len(kv[1]))):
        if lb == -1:
            rep = members[0]
        else:
            sub = features[members]
            centroid = sub.mean(axis=0)
            sims = sub @ centroid
            rep = int(members[int(sims.argmax())])
        clusters.append({
            "cluster_id": int(lb),
            "size": len(members),
            "representative_slide_index": int(rep),
            "member_slide_indices": [int(i) for i in members],
            "representative_layout": metas[rep].layout_name,
            "representative_structure_sig": metas[rep].structure_sig,
            "representative_leaf_count": metas[rep].shape_count_leaf,
        })

    with open(output_dir / f"clusters_{tag}.json", "w", encoding="utf-8") as f:
        json.dump(clusters, f, ensure_ascii=False, indent=2)
    with open(output_dir / f"cluster_labels_{tag}.json", "w", encoding="utf-8") as f:
        json.dump({"labels": [int(l) for l in labels]}, f, ensure_ascii=False, indent=2)


def run_stage_2a(
    meta_path: Path,
    features_path: Path,
    output_dir: Path,
    min_cluster_size: int = 3,
    min_samples: int = 2,
) -> dict:
    """Stage 2a 엔드-투-엔드."""
    with open(meta_path, "r", encoding="utf-8") as f:
        metas = [SlideMeta.model_validate(x) for x in json.load(f)]
    features = np.load(features_path).astype(np.float32)
    print(f"[cluster] feature shape={features.shape}", flush=True)

    labels, info = cluster_hdbscan(
        features, min_cluster_size=min_cluster_size, min_samples=min_samples,
    )
    print(f"[cluster] {info}", flush=True)

    eval_res = evaluate_clusters(features, labels, metas)
    print(f"[evaluate] {eval_res}", flush=True)

    save_clusters(metas, features, labels, output_dir, tag="v2")

    # summary
    summary = {"cluster_info": info, "evaluation": eval_res}
    with open(output_dir / "cluster_summary_v2.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    return summary


if __name__ == "__main__":
    import argparse
    ROOT = Path(__file__).resolve().parent.parent.parent
    p = argparse.ArgumentParser()
    p.add_argument("--meta", default=str(ROOT / "output" / "catalog" / "slide_meta.json"))
    p.add_argument("--features", default=str(ROOT / "output" / "catalog" / "shape_features.npy"))
    p.add_argument("--out", default=str(ROOT / "output" / "catalog"))
    p.add_argument("--min-cluster-size", type=int, default=3)
    p.add_argument("--min-samples", type=int, default=2)
    args = p.parse_args()
    run_stage_2a(
        Path(args.meta), Path(args.features), Path(args.out),
        min_cluster_size=args.min_cluster_size, min_samples=args.min_samples,
    )
