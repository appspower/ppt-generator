"""Stage 2b 최종 — Shape (68-dim) + DiT (768-dim) 앙상블 클러스터링.

Feature 합성
-----------
  S (shape_features.npy) : 68-dim, L2-normed, 가중 w_s
  V (dit_embeddings.npy) : 768-dim, L2-normed, 가중 w_v
  → concat([w_s·S, w_v·V]) → L2-normed final feature (836-dim)

기본 가중치
----------
  w_s = 0.5, w_v = 0.5 (균등) — Stage 2a purity가 0.13으로 약했으므로
  visual을 동등 비중. 검증 후 튜닝.
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np

from .cluster_v2 import cluster_hdbscan, evaluate_clusters, save_clusters
from .schemas import SlideMeta


def build_ensemble(
    shape_feats: np.ndarray,
    vision_feats: np.ndarray,
    w_shape: float = 0.5,
    w_vision: float = 0.5,
) -> np.ndarray:
    """가중 concat + L2 normalize."""
    # 이미 각자 normalized — 가중치만 곱해서 concat 후 재 normalize
    s = shape_feats * w_shape
    v = vision_feats * w_vision
    combined = np.concatenate([s, v], axis=1).astype(np.float32)
    norms = np.linalg.norm(combined, axis=1, keepdims=True)
    norms[norms == 0] = 1
    return combined / norms


def run_stage_2b_cluster(
    meta_path: Path,
    shape_features_path: Path,
    vision_features_path: Path,
    output_dir: Path,
    w_shape: float = 0.5,
    w_vision: float = 0.5,
    min_cluster_size: int = 3,
    min_samples: int = 2,
) -> dict:
    """Stage 2a + DiT 앙상블 → HDBSCAN → 검증 저장."""
    with open(meta_path, "r", encoding="utf-8") as f:
        metas = [SlideMeta.model_validate(x) for x in json.load(f)]
    shape_feats = np.load(shape_features_path).astype(np.float32)
    vision_feats = np.load(vision_features_path).astype(np.float32)
    print(f"[ensemble] shape={shape_feats.shape}  vision={vision_feats.shape}", flush=True)

    feats = build_ensemble(shape_feats, vision_feats, w_shape, w_vision)
    print(f"[ensemble] combined shape={feats.shape}  w_s={w_shape} w_v={w_vision}", flush=True)

    labels, info = cluster_hdbscan(
        feats, min_cluster_size=min_cluster_size, min_samples=min_samples,
    )
    print(f"[cluster] {info}", flush=True)

    eval_res = evaluate_clusters(feats, labels, metas)
    print(f"[evaluate] {eval_res}", flush=True)

    save_clusters(metas, feats, labels, output_dir, tag="v3")

    summary = {
        "cluster_info": info,
        "evaluation": eval_res,
        "weights": {"shape": w_shape, "vision": w_vision},
    }
    with open(output_dir / "cluster_summary_v3.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    return summary


if __name__ == "__main__":
    import argparse
    ROOT = Path(__file__).resolve().parent.parent.parent
    p = argparse.ArgumentParser()
    p.add_argument("--meta", default=str(ROOT / "output" / "catalog" / "slide_meta.json"))
    p.add_argument("--shape", default=str(ROOT / "output" / "catalog" / "shape_features.npy"))
    p.add_argument("--vision", default=str(ROOT / "output" / "catalog" / "dit_embeddings.npy"))
    p.add_argument("--out", default=str(ROOT / "output" / "catalog"))
    p.add_argument("--w-shape", type=float, default=0.5)
    p.add_argument("--w-vision", type=float, default=0.5)
    p.add_argument("--min-cluster-size", type=int, default=3)
    p.add_argument("--min-samples", type=int, default=2)
    args = p.parse_args()
    run_stage_2b_cluster(
        Path(args.meta), Path(args.shape), Path(args.vision),
        Path(args.out),
        w_shape=args.w_shape, w_vision=args.w_vision,
        min_cluster_size=args.min_cluster_size, min_samples=args.min_samples,
    )
