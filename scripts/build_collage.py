"""1,251 PNG → 16x16 grid 5 collages with index labels.

cell 96×54 (16:9 비율 유지)
collage 1568×864 (Claude Vision 입력 가능 크기)
인덱스 라벨: 각 셀 좌상단에 작은 흰색 박스 + 검은 텍스트
"""
from __future__ import annotations

import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


CELL_W = 156
CELL_H = 117  # PPT 4:3 비율 (원본 1568x1176)
GRID_COLS = 10
GRID_ROWS = 10
COLLAGE_W = CELL_W * GRID_COLS  # 1560
COLLAGE_H = CELL_H * GRID_ROWS  # 1170 (Claude Vision 1568 안)


def _label_font():
    try:
        return ImageFont.truetype("arial.ttf", 9)
    except Exception:
        return ImageFont.load_default()


def build_collages(png_dir: Path, out_dir: Path) -> list[Path]:
    pngs = sorted(png_dir.glob("slide_*.png"))
    print(f"[collage] {len(pngs)} PNGs", flush=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    cells_per = GRID_COLS * GRID_ROWS  # 256
    n_collages = (len(pngs) + cells_per - 1) // cells_per
    print(f"[collage] {n_collages} collages × {cells_per} cells each", flush=True)

    font = _label_font()
    out_paths: list[Path] = []
    manifest: list[dict] = []

    for ci in range(n_collages):
        canvas = Image.new("RGB", (COLLAGE_W, COLLAGE_H), "white")
        draw = ImageDraw.Draw(canvas)

        start = ci * cells_per
        end = min(start + cells_per, len(pngs))
        for k, png in enumerate(pngs[start:end]):
            idx = start + k
            row = k // GRID_COLS
            col = k % GRID_COLS
            x = col * CELL_W
            y = row * CELL_H

            try:
                im = Image.open(png).convert("RGB")
                im = im.resize((CELL_W, CELL_H), Image.LANCZOS)
                canvas.paste(im, (x, y))
            except Exception as e:
                print(f"  [err] {png.name}: {e}")
                continue

            # 인덱스 라벨 (좌상단, 흰색 배경 + 검은 텍스트)
            label = f"{idx}"
            tw = len(label) * 6 + 4
            draw.rectangle([(x, y), (x + tw, y + 11)], fill="white")
            draw.text((x + 1, y), label, fill="black", font=font)

            # 셀 경계선 (옅은 회색)
            draw.rectangle(
                [(x, y), (x + CELL_W - 1, y + CELL_H - 1)],
                outline="#cccccc", width=1,
            )

        out = out_dir / f"collage_{ci+1:02d}_of_{n_collages}.png"
        canvas.save(out, "PNG", optimize=True)
        out_paths.append(out)
        manifest.append({
            "collage_index": ci + 1,
            "file": out.name,
            "slide_index_range": [start, end - 1],
            "cell_count": end - start,
            "grid": f"{GRID_COLS}x{GRID_ROWS}",
        })
        print(f"  [ok] {out.name}: slides {start}..{end-1}", flush=True)

    with open(out_dir / "_collage_manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return out_paths


if __name__ == "__main__":
    ROOT = Path(__file__).resolve().parent.parent
    png_dir = ROOT / "output" / "catalog" / "all_pngs"
    out_dir = ROOT / "output" / "catalog" / "collages"
    build_collages(png_dir, out_dir)
