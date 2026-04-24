"""Track 1 Stage 2b — DiT (Document Image Transformer) visual embedding.

근거
----
arXiv 2203.02378 (DiT, Microsoft): document-pretrained ViT + MIM.
arXiv 2506.12116: DiT ARI 0.96+ in unsupervised doc template clustering,
                 SBERT/text는 placeholder 데이터에서 degenerate.

구현
----
- model: `microsoft/dit-base` (~345MB, 768-dim CLS)
- input: 224x224 PIL (원본 1568x1176 PNG 리사이즈)
- output: 768-dim CLS embedding (last_hidden_state[:, 0, :])
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
from PIL import Image


MODEL_NAME = "microsoft/dit-base"


def load_dit():
    """Transformers + DiT 모델 로드 (CPU/GPU 자동 선택)."""
    import torch
    from transformers import AutoImageProcessor, AutoModel

    print(f"[dit] loading {MODEL_NAME} ...", flush=True)
    processor = AutoImageProcessor.from_pretrained(MODEL_NAME)
    model = AutoModel.from_pretrained(MODEL_NAME)
    device = "cuda" if torch.cuda.is_available() else "cpu"
    model = model.to(device).eval()
    print(f"[dit] device={device}", flush=True)
    return model, processor, device


def embed_pngs(
    png_dir: Path,
    n_slides: int,
    batch_size: int = 16,
) -> np.ndarray:
    """각 slide_XXXX.png → 768-dim CLS embedding (L2-normalized)."""
    import torch

    model, processor, device = load_dit()

    embeddings = np.zeros((n_slides, 768), dtype=np.float32)
    missing = 0

    for batch_start in range(0, n_slides, batch_size):
        batch_end = min(batch_start + batch_size, n_slides)
        imgs = []
        batch_indices = []
        for i in range(batch_start, batch_end):
            p = png_dir / f"slide_{i:04d}.png"
            if not p.exists():
                missing += 1
                imgs.append(Image.new("RGB", (224, 224), "white"))
            else:
                imgs.append(Image.open(p).convert("RGB"))
            batch_indices.append(i)

        inputs = processor(images=imgs, return_tensors="pt")
        inputs = {k: v.to(device) for k, v in inputs.items()}
        with torch.no_grad():
            out = model(**inputs)
        # last_hidden_state : (B, seq, 768), CLS = idx 0
        cls = out.last_hidden_state[:, 0, :].cpu().numpy()
        # L2 normalize
        norms = np.linalg.norm(cls, axis=1, keepdims=True)
        norms[norms == 0] = 1
        cls = cls / norms

        for k, i in enumerate(batch_indices):
            embeddings[i] = cls[k]

        if batch_start % (batch_size * 10) == 0:
            print(f"  [{batch_end}/{n_slides}] embedded (missing={missing})", flush=True)

    print(f"[dit] done. missing={missing}", flush=True)
    return embeddings


def run_stage_2b_embed(
    png_dir: Path,
    meta_path: Path,
    output_path: Path,
    batch_size: int = 16,
) -> np.ndarray:
    with open(meta_path, "r", encoding="utf-8") as f:
        n = len(json.load(f))
    embs = embed_pngs(png_dir, n, batch_size=batch_size)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    np.save(output_path, embs)
    print(f"[save] shape={embs.shape} -> {output_path}", flush=True)
    return embs


if __name__ == "__main__":
    ROOT = Path(__file__).resolve().parent.parent.parent
    pngs = ROOT / "output" / "catalog" / "all_pngs"
    meta = ROOT / "output" / "catalog" / "slide_meta.json"
    out = ROOT / "output" / "catalog" / "dit_embeddings.npy"
    run_stage_2b_embed(pngs, meta, out)
