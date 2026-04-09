"""PowerPoint COM 기반 PNG 익스포터.

Windows + PowerPoint 설치 환경 전제. python-pptx로는 PNG 렌더가 불가능하므로
PowerPoint Application COM 객체를 직접 호출해 슬라이드별 PNG를 추출한다.
"""

from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client


def pptx_to_pngs(
    pptx_path: Path,
    output_dir: Path,
    width: int = 1920,
    height: int = 1440,
) -> list[Path]:
    """pptx 파일을 슬라이드별 PNG로 변환.

    Args:
        pptx_path: 입력 .pptx 파일
        output_dir: PNG 출력 디렉토리 (자동 생성)
        width: PNG 가로 픽셀 (4:3 기준 1920×1440 권장)
        height: PNG 세로 픽셀

    Returns:
        생성된 PNG 파일 경로 리스트 (슬라이드 순)
    """
    pptx_path = Path(pptx_path).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")

    pythoncom.CoInitialize()
    powerpoint = None
    presentation = None
    png_paths: list[Path] = []

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # WithWindow=False는 일부 환경에서 실패하므로 기본값으로 연다
        presentation = powerpoint.Presentations.Open(
            str(pptx_path),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )

        for idx, slide in enumerate(presentation.Slides, start=1):
            png_path = output_dir / f"slide_{idx:02d}.png"
            # Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight)
            slide.Export(str(png_path), "PNG", width, height)
            png_paths.append(png_path)

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

    return png_paths
