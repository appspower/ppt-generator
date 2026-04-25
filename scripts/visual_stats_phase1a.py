"""Phase 1A: Visual diversity statistics for 1,251 master template slides.

Measures (per slide):
  A1. Color profile: dominant 3 colors, accent ratio, grayscale flag
  A2. Area occupancy: text-like / shape-like / blank, split patterns
  A3. Alignment entropy on 8x8 activity grid
  A4. Visual element area ratios (heuristic)

Outputs:
  output/catalog/phase1a_visual_stats.json (per-slide + global summary)
"""
from __future__ import annotations

import json
import math
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Any

import numpy as np
from PIL import Image
from sklearn.cluster import MiniBatchKMeans

ROOT = Path(r"c:/Users/y2kbo/Apps/PPT")
PNG_DIR = ROOT / "output" / "catalog" / "all_pngs"
OUT_PATH = ROOT / "output" / "catalog" / "phase1a_visual_stats.json"

# Downsample for k-means speed: 1568x1176 -> 392x294
DOWNSAMPLE_W = 392
DOWNSAMPLE_H = 294

# Activity grid for entropy
GRID = 8

# Thresholds
WHITE_THR = 240          # pixel >= this on all RGB -> blank/white
BLACK_THR = 60           # pixel <= this on all RGB -> text-like (dark)
SAT_THR = 30 / 255.0     # HSV saturation threshold for accent (0..1)
GRAYSCALE_ACCENT_RATIO = 0.005  # if accent < 0.5%, slide is grayscale


def rgb_to_hsv_array(rgb: np.ndarray) -> np.ndarray:
    """rgb shape (..., 3) uint8 -> hsv float (..., 3) in 0..1."""
    arr = rgb.astype(np.float32) / 255.0
    r, g, b = arr[..., 0], arr[..., 1], arr[..., 2]
    mx = np.max(arr, axis=-1)
    mn = np.min(arr, axis=-1)
    diff = mx - mn
    # value
    v = mx
    # saturation
    s = np.where(mx > 0, diff / np.maximum(mx, 1e-9), 0.0)
    # hue (degrees / 360)
    h = np.zeros_like(mx)
    mask = diff > 1e-9
    rc = np.where(mask, (mx - r) / np.maximum(diff, 1e-9), 0.0)
    gc = np.where(mask, (mx - g) / np.maximum(diff, 1e-9), 0.0)
    bc = np.where(mask, (mx - b) / np.maximum(diff, 1e-9), 0.0)
    h = np.where((mx == r) & mask, (bc - gc), h)
    h = np.where((mx == g) & mask, 2.0 + (rc - bc), h)
    h = np.where((mx == b) & mask, 4.0 + (gc - rc), h)
    h = (h / 6.0) % 1.0
    out = np.stack([h, s, v], axis=-1)
    return out


def dominant_colors(pixels: np.ndarray, k: int = 3) -> list[dict]:
    """k-means on a sample of pixels. pixels shape (N, 3) uint8."""
    n = pixels.shape[0]
    sample_size = min(n, 4000)
    idx = np.random.choice(n, sample_size, replace=False)
    sample = pixels[idx].astype(np.float32)
    km = MiniBatchKMeans(n_clusters=k, n_init=3, max_iter=50, batch_size=512, random_state=0)
    km.fit(sample)
    labels = km.predict(sample)
    out = []
    for i in range(k):
        cnt = int((labels == i).sum())
        c = km.cluster_centers_[i]
        out.append({
            "rgb": [int(round(c[0])), int(round(c[1])), int(round(c[2]))],
            "ratio": cnt / sample_size,
        })
    out.sort(key=lambda d: -d["ratio"])
    return out


def compute_stats_for_image(path: Path) -> dict[str, Any]:
    """All A1..A4 stats."""
    with Image.open(path) as img:
        img = img.convert("RGB").resize((DOWNSAMPLE_W, DOWNSAMPLE_H), Image.BILINEAR)
        arr = np.asarray(img)  # (H, W, 3) uint8

    H, W = arr.shape[:2]
    flat = arr.reshape(-1, 3)
    hsv = rgb_to_hsv_array(arr)
    sat = hsv[..., 1]

    # ---- A1. color profile ----
    dom = dominant_colors(flat, k=3)
    accent_mask = sat >= SAT_THR
    accent_ratio = float(accent_mask.mean())
    grayscale = accent_ratio < GRAYSCALE_ACCENT_RATIO

    # accent color top hue (binned to 12 hues)
    if accent_mask.sum() > 0:
        accent_hues = hsv[..., 0][accent_mask]
        # representative hue: median in the largest bin
        hue_bin = np.floor(accent_hues * 12).astype(np.int32)
        bins = np.bincount(hue_bin, minlength=12)
        top_bin = int(bins.argmax())
        # use mean rgb of pixels in that hue bin as accent rep
        bin_mask = accent_mask & (np.floor(hsv[..., 0] * 12).astype(np.int32) == top_bin)
        if bin_mask.sum() > 0:
            accent_rgb = arr[bin_mask].mean(axis=0).astype(int).tolist()
        else:
            accent_rgb = [0, 0, 0]
        accent_hue_deg = (top_bin + 0.5) / 12 * 360
    else:
        accent_rgb = [0, 0, 0]
        accent_hue_deg = -1.0

    # ---- A2. area occupancy ----
    # white: all rgb >= WHITE_THR
    white_mask = (arr >= WHITE_THR).all(axis=-1)
    # text-like dark: all rgb <= BLACK_THR
    dark_mask = (arr <= BLACK_THR).all(axis=-1)
    # shape-like: not white, not dark, low-mid saturation OR borders (mid-gray range)
    mid_mask = ~white_mask & ~dark_mask
    blank_ratio = float(white_mask.mean())
    dark_ratio = float(dark_mask.mean())
    mid_ratio = float(mid_mask.mean())

    # Splits: occupancy in halves
    left = ~white_mask[:, : W // 2]
    right = ~white_mask[:, W // 2 :]
    top = ~white_mask[: H // 2, :]
    bot = ~white_mask[H // 2 :, :]
    left_occ = float(left.mean())
    right_occ = float(right.mean())
    top_occ = float(top.mean())
    bot_occ = float(bot.mean())
    lr_diff = abs(left_occ - right_occ)
    tb_diff = abs(top_occ - bot_occ)

    # 3-split horizontal (left/center/right)
    third = W // 3
    occ_l = float((~white_mask[:, :third]).mean())
    occ_c = float((~white_mask[:, third : 2 * third]).mean())
    occ_r = float((~white_mask[:, 2 * third :]).mean())
    # 4-split: 2x2 grid
    quad = [
        float((~white_mask[: H // 2, : W // 2]).mean()),
        float((~white_mask[: H // 2, W // 2 :]).mean()),
        float((~white_mask[H // 2 :, : W // 2]).mean()),
        float((~white_mask[H // 2 :, W // 2 :]).mean()),
    ]

    # ---- A3. alignment entropy on GRID x GRID ----
    cell_h = H // GRID
    cell_w = W // GRID
    grid_act = np.zeros((GRID, GRID), dtype=np.float32)
    for r in range(GRID):
        for c in range(GRID):
            sub = ~white_mask[r * cell_h : (r + 1) * cell_h, c * cell_w : (c + 1) * cell_w]
            grid_act[r, c] = float(sub.mean())
    # entropy of normalized activity
    p = grid_act.flatten()
    if p.sum() > 0:
        p = p / p.sum()
        nz = p[p > 0]
        ent = float(-(nz * np.log2(nz)).sum())
        max_ent = math.log2(GRID * GRID)
        ent_norm = ent / max_ent
    else:
        ent = 0.0
        ent_norm = 0.0

    # ---- A4. element-type area heuristic ----
    # text-like: dark pixels in small/horizontal strokes -> dark_ratio approx
    # chart/line-like: mid_mask with high local variance
    # large box: large continuous mid_mask blocks
    # icon-like: small accent regions
    text_area = dark_ratio
    # variance proxy: compute std over 16x16 tiles
    tile = 16
    tH, tW = H // tile, W // tile
    block = arr[: tH * tile, : tW * tile].reshape(tH, tile, tW, tile, 3).mean(axis=(1, 3))
    block_std = arr[: tH * tile, : tW * tile].astype(np.float32).reshape(tH, tile, tW, tile, 3).std(axis=(1, 3)).mean(axis=-1)
    # high variance tiles (lines/charts)
    chart_like = float((block_std > 25).mean())
    # large mid-tone areas (filled boxes): low variance + non-white tile
    tile_white = (block.mean(axis=-1) >= WHITE_THR)
    tile_dark = (block.mean(axis=-1) <= BLACK_THR)
    big_box_like = float(((~tile_white) & (~tile_dark) & (block_std < 12)).mean())
    icon_like = float((accent_mask & ~white_mask).mean()) - chart_like
    icon_like = max(0.0, icon_like)

    return {
        # A1
        "dominant_colors": dom,
        "accent_ratio": accent_ratio,
        "grayscale": bool(grayscale),
        "accent_rgb": accent_rgb,
        "accent_hue_deg": accent_hue_deg,
        # A2
        "blank_ratio": blank_ratio,
        "dark_ratio": dark_ratio,
        "mid_ratio": mid_ratio,
        "split_lr_diff": lr_diff,
        "split_tb_diff": tb_diff,
        "occ_left": left_occ, "occ_right": right_occ,
        "occ_top": top_occ, "occ_bot": bot_occ,
        "thirds": [occ_l, occ_c, occ_r],
        "quads": quad,
        # A3
        "grid_entropy": ent,
        "grid_entropy_norm": ent_norm,
        # A4
        "text_area": text_area,
        "chart_like_area": chart_like,
        "big_box_area": big_box_like,
        "icon_like_area": icon_like,
    }


def hue_label(hue_deg: float) -> str:
    if hue_deg < 0:
        return "none"
    bins = [
        (15, "red"), (45, "orange"), (75, "yellow"),
        (165, "green"), (195, "cyan"), (255, "blue"),
        (285, "purple"), (345, "magenta"), (361, "red"),
    ]
    for thr, name in bins:
        if hue_deg < thr:
            return name
    return "red"


def _worker(path_str: str) -> dict[str, Any]:
    p = Path(path_str)
    try:
        stats = compute_stats_for_image(p)
    except Exception as e:
        stats = {"error": str(e)}
    stats["slide_index"] = int(p.stem.split("_")[1])
    return stats


def main() -> None:
    pngs = sorted(PNG_DIR.glob("slide_*.png"))
    print(f"Found {len(pngs)} PNGs")
    per_slide: list[dict[str, Any]] = []
    t0 = time.time()
    import os
    workers = max(2, (os.cpu_count() or 4) - 1)
    print(f"using {workers} workers")
    with ProcessPoolExecutor(max_workers=workers) as ex:
        futures = {ex.submit(_worker, str(p)): p for p in pngs}
        done = 0
        for fut in as_completed(futures):
            per_slide.append(fut.result())
            done += 1
            if done % 100 == 0:
                elapsed = time.time() - t0
                rate = done / elapsed
                eta = (len(pngs) - done) / rate
                print(f"  {done}/{len(pngs)}  rate={rate:.1f}/s  eta={eta:.0f}s")
    per_slide.sort(key=lambda s: s["slide_index"])
    print(f"All done in {time.time()-t0:.1f}s")

    # ---- global summary ----
    valid = [s for s in per_slide if "error" not in s]
    total = len(valid)
    grayscale_n = sum(1 for s in valid if s["grayscale"])
    high_accent_n = sum(1 for s in valid if s["accent_ratio"] >= 0.10)
    mid_accent_n = sum(1 for s in valid if 0.02 <= s["accent_ratio"] < 0.10)

    # accent color top 10 (by binned hue label of accent_rgb among colored slides)
    hue_count: dict[str, int] = {}
    for s in valid:
        if s["grayscale"]:
            continue
        lab = hue_label(s["accent_hue_deg"])
        hue_count[lab] = hue_count.get(lab, 0) + 1
    top_hues = sorted(hue_count.items(), key=lambda x: -x[1])[:10]

    # entropy distribution
    ent_vals = [s["grid_entropy_norm"] for s in valid]
    ent_low = sum(1 for v in ent_vals if v < 0.4)
    ent_mid = sum(1 for v in ent_vals if 0.4 <= v < 0.7)
    ent_high = sum(1 for v in ent_vals if v >= 0.7)

    # split patterns: classify
    def classify_split(s: dict) -> str:
        lr = s["split_lr_diff"]
        tb = s["split_tb_diff"]
        thirds = s["thirds"]
        quads = s["quads"]
        thirds_var = float(np.std(thirds))
        quads_var = float(np.std(quads))
        # heuristic categories
        if max(s["occ_left"], s["occ_right"], s["occ_top"], s["occ_bot"]) < 0.05:
            return "near_blank"
        if lr > 0.15 and lr > tb:
            return "lr_split"
        if tb > 0.15 and tb > lr:
            return "tb_split"
        if thirds_var > 0.10:
            return "three_split"
        if quads_var > 0.08:
            return "four_split"
        return "balanced_full"

    split_dist: dict[str, int] = {}
    for s in valid:
        c = classify_split(s)
        split_dist[c] = split_dist.get(c, 0) + 1

    # density: blank ratio buckets
    blank_low = sum(1 for s in valid if s["blank_ratio"] < 0.3)
    blank_mid = sum(1 for s in valid if 0.3 <= s["blank_ratio"] < 0.6)
    blank_high = sum(1 for s in valid if s["blank_ratio"] >= 0.6)

    # outliers: extreme grid entropy and extreme accent ratio
    valid_sorted_ent = sorted(valid, key=lambda s: -s["grid_entropy_norm"])
    valid_sorted_acc = sorted(valid, key=lambda s: -s["accent_ratio"])
    valid_sorted_blank = sorted(valid, key=lambda s: -s["blank_ratio"])
    outliers = {
        "highest_entropy": [s["slide_index"] for s in valid_sorted_ent[:5]],
        "highest_accent_ratio": [s["slide_index"] for s in valid_sorted_acc[:5]],
        "highest_blank_ratio": [s["slide_index"] for s in valid_sorted_blank[:5]],
        "lowest_entropy_nonblank": [
            s["slide_index"]
            for s in sorted(
                [s for s in valid if s["blank_ratio"] < 0.85],
                key=lambda s: s["grid_entropy_norm"],
            )[:5]
        ],
    }

    summary = {
        "total_slides": total,
        "grayscale_slides": grayscale_n,
        "color_slides": total - grayscale_n,
        "high_accent_slides": high_accent_n,  # >=10% saturated
        "mid_accent_slides": mid_accent_n,
        "top_accent_hues": top_hues,
        "entropy_distribution": {
            "low_<0.4": ent_low,
            "mid_0.4-0.7": ent_mid,
            "high_>=0.7": ent_high,
        },
        "blank_ratio_distribution": {
            "dense_<0.3": blank_low,
            "mid_0.3-0.6": blank_mid,
            "sparse_>=0.6": blank_high,
        },
        "split_pattern_distribution": split_dist,
        "mean_blank_ratio": float(np.mean([s["blank_ratio"] for s in valid])),
        "mean_dark_ratio": float(np.mean([s["dark_ratio"] for s in valid])),
        "mean_accent_ratio": float(np.mean([s["accent_ratio"] for s in valid])),
        "mean_chart_like_area": float(np.mean([s["chart_like_area"] for s in valid])),
        "mean_big_box_area": float(np.mean([s["big_box_area"] for s in valid])),
        "outliers": outliers,
    }

    out = {"per_slide": per_slide, "summary": summary}
    OUT_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nWrote {OUT_PATH} ({OUT_PATH.stat().st_size/1024:.1f} KB)")
    print("\n--- SUMMARY ---")
    for k, v in summary.items():
        if k == "outliers":
            continue
        print(f"  {k}: {v}")
    print(f"  outliers: {outliers}")


if __name__ == "__main__":
    main()
