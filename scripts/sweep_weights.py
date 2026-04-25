"""Stage 2b 가중치 sweep — w_shape/w_vision 6 조합 비교."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import json
import numpy as np

from ppt_builder.catalog.schemas import SlideMeta
from ppt_builder.catalog.cluster_v2 import cluster_hdbscan, evaluate_clusters
from ppt_builder.catalog.cluster_v3 import build_ensemble


ROOT = Path(__file__).resolve().parent.parent
meta_path = ROOT / "output" / "catalog" / "slide_meta.json"
shape_path = ROOT / "output" / "catalog" / "shape_features.npy"
vision_path = ROOT / "output" / "catalog" / "dit_embeddings.npy"

with open(meta_path, "r", encoding="utf-8") as f:
    metas = [SlideMeta.model_validate(x) for x in json.load(f)]
shape_feats = np.load(shape_path).astype(np.float32)
vision_feats = np.load(vision_path).astype(np.float32)
print(f"shape={shape_feats.shape}  vision={vision_feats.shape}")

# 가중치 조합 시도
configs = [
    (0.0, 1.0),  # vision only
    (0.2, 0.8),
    (0.3, 0.7),
    (0.5, 0.5),  # 기본
    (0.7, 0.3),
    (1.0, 0.0),  # shape only (참고)
]

results = []
for w_s, w_v in configs:
    feats = build_ensemble(shape_feats, vision_feats, w_s, w_v)
    labels, info = cluster_hdbscan(feats, min_cluster_size=3, min_samples=2)
    eval_res = evaluate_clusters(feats, labels, metas)
    results.append({
        "w_shape": w_s, "w_vision": w_v,
        "cluster_count": info["cluster_count"],
        "noise_initial": info["noise_count_initial"],
        "max_cluster": info["max_cluster_size"],
        "median": info["median_cluster_size"],
        "silhouette": eval_res["silhouette"],
        "purity_mean": eval_res["structure_sig_purity_mean"],
        "purity_p25": eval_res["structure_sig_purity_p25"],
    })
    print(f"  w_s={w_s} w_v={w_v}: clusters={info['cluster_count']:>3}  "
          f"max={info['max_cluster_size']:>3}  noise_initial={info['noise_count_initial']:>4}  "
          f"silhouette={eval_res['silhouette']:.3f}  purity={eval_res['structure_sig_purity_mean']:.3f}")

with open(ROOT / "output" / "catalog" / "weight_sweep.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("\n[best by silhouette]")
best = max(results, key=lambda x: x["silhouette"] or 0)
print(f"  w_s={best['w_shape']} w_v={best['w_vision']} silhouette={best['silhouette']:.3f}")

print("\n[best by purity]")
best = max(results, key=lambda x: x["purity_mean"])
print(f"  w_s={best['w_shape']} w_v={best['w_vision']} purity={best['purity_mean']:.3f}")
