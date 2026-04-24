"""1,251장 전체 PNG 렌더 — Stage 2b (DiT embedding) 용.

PowerPoint COM Application을 통해 슬라이드별 PNG 추출.
시간: 1,251장 기준 20-40분 (단일 pptx라 Open 1회).
"""
from __future__ import annotations

import json
from pathlib import Path
import sys
import time

import pythoncom
import win32com.client


def render_all(
    pptx_path: Path,
    output_dir: Path,
    width: int = 1568,
    height: int = 1176,
    progress_every: int = 50,
) -> list[Path]:
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
        print(f"[open] {total} slides", flush=True)
        t0 = time.time()

        for i in range(1, total + 1):
            slide_idx_0 = i - 1
            out_path = output_dir / f"slide_{slide_idx_0:04d}.png"
            if out_path.exists():
                out_paths.append(out_path)
                continue
            try:
                slide = presentation.Slides(i)
                slide.Export(str(out_path), "PNG", width, height)
                out_paths.append(out_path)
            except Exception as e:
                print(f"  [err] #{slide_idx_0}: {e}", flush=True)

            if i % progress_every == 0:
                elapsed = time.time() - t0
                eta = elapsed / i * (total - i)
                print(f"  [{i}/{total}] elapsed={elapsed:.0f}s  ETA={eta:.0f}s", flush=True)

        print(f"[done] {len(out_paths)} PNGs in {time.time()-t0:.0f}s", flush=True)
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
    pptx = ROOT / "docs" / "references" / "_master_templates" / "PPT 템플릿.pptx"
    out = ROOT / "output" / "catalog" / "all_pngs"
    render_all(pptx, out)
