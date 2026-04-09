"""HTML/Tailwind → PNG 렌더러 (Playwright Chromium 기반).

복잡한 도식(2×2 매트릭스, BCG, 워터폴 등)을 python-pptx로 그리기 어려운 경우
HTML+CSS로 작성해 PNG로 렌더한 뒤 슬라이드에 이미지로 삽입한다.

Track A의 content.py 렌더러는 현재 ImageComponent를 처리하지 않으므로
Track C가 직접 python-pptx로 image-slide pptx를 빌드한다 (build_image_slide_pptx).

회사 컬러:
    Orange: #FD5108 / #FE7C39 / #FFAA72
    Grey:   #A1A8B3 / #B5BCC4 / #CBD1D6
    Body:   #000000
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from playwright.sync_api import sync_playwright
from pptx import Presentation
from pptx.util import Emu, Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


# 회사 컬러 팔레트 (project_ppt_generator.md 메모리 기반)
COLORS = {
    "accent": "#FD5108",
    "accent_med": "#FE7C39",
    "accent_light": "#FFAA72",
    "grey_dark": "#A1A8B3",
    "grey_med": "#B5BCC4",
    "grey_light": "#CBD1D6",
    "body": "#000000",
}


def html_to_png(
    html: str,
    output_path: Path,
    width: int = 1600,
    height: int = 1000,
    device_scale_factor: int = 2,
) -> Path:
    """HTML 문자열을 PNG로 렌더한다.

    Args:
        html: 완전한 HTML 문서 (Tailwind CDN 포함 권장)
        output_path: PNG 출력 경로
        width, height: 뷰포트 픽셀 (기본 1600×1000, 4:3 적합)
        device_scale_factor: 2x retina 렌더링 (인쇄 품질)

    Returns:
        생성된 PNG 경로
    """
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        try:
            context = browser.new_context(
                viewport={"width": width, "height": height},
                device_scale_factor=device_scale_factor,
            )
            page = context.new_page()
            page.set_content(html, wait_until="networkidle")
            page.screenshot(path=str(output_path), full_page=False, omit_background=False)
        finally:
            browser.close()

    return output_path


# ============================================================
# 템플릿: 2×2 매트릭스 (BCG 스타일)
# ============================================================

_BCG_TEMPLATE = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>2x2 Matrix</title>
<script src="https://cdn.tailwindcss.com"></script>
<style>
  body {{ font-family: 'Pretendard', 'Malgun Gothic', sans-serif; }}
  .quad {{ position: relative; }}
  .quad-label {{ position: absolute; top: 0.75rem; left: 0.75rem; font-weight: 700; font-size: 0.875rem; opacity: 0.7; }}
</style>
</head>
<body class="bg-white p-12">
  <h1 class="text-3xl font-bold text-black mb-2">{title}</h1>
  <p class="text-base text-gray-600 mb-8">{subtitle}</p>

  <div class="relative" style="height: 720px;">
    <!-- Y axis label -->
    <div class="absolute left-0 top-1/2 -translate-y-1/2 -translate-x-12 -rotate-90 origin-center text-sm font-semibold text-gray-700 whitespace-nowrap">{y_axis_label}</div>
    <!-- Y high/low -->
    <div class="absolute left-2 top-2 text-xs font-bold text-gray-500">↑ {y_high}</div>
    <div class="absolute left-2 bottom-2 text-xs font-bold text-gray-500">↓ {y_low}</div>

    <!-- X axis label -->
    <div class="absolute left-1/2 -bottom-10 -translate-x-1/2 text-sm font-semibold text-gray-700">{x_axis_label}</div>
    <!-- X low/high -->
    <div class="absolute -bottom-6 left-2 text-xs font-bold text-gray-500">← {x_low}</div>
    <div class="absolute -bottom-6 right-2 text-xs font-bold text-gray-500">{x_high} →</div>

    <!-- 2x2 grid -->
    <div class="grid grid-cols-2 grid-rows-2 gap-1 ml-12" style="height: 100%;">
      <!-- Top-left (low x, high y) -->
      <div class="quad p-6 border-2" style="background-color: #FFAA72; border-color: #FD5108;">
        <div class="quad-label">{q1_label}</div>
        <div class="mt-8">
          <h3 class="text-xl font-bold text-black mb-3">{q1_title}</h3>
          <ul class="text-sm text-black space-y-2">{q1_items}</ul>
        </div>
      </div>
      <!-- Top-right (high x, high y) -->
      <div class="quad p-6 border-2" style="background-color: #FD5108; color: white; border-color: #FD5108;">
        <div class="quad-label" style="color: white;">{q2_label}</div>
        <div class="mt-8">
          <h3 class="text-xl font-bold mb-3">{q2_title}</h3>
          <ul class="text-sm space-y-2">{q2_items}</ul>
        </div>
      </div>
      <!-- Bottom-left (low x, low y) -->
      <div class="quad p-6 border-2" style="background-color: #CBD1D6; border-color: #A1A8B3;">
        <div class="quad-label">{q3_label}</div>
        <div class="mt-8">
          <h3 class="text-xl font-bold text-black mb-3">{q3_title}</h3>
          <ul class="text-sm text-black space-y-2">{q3_items}</ul>
        </div>
      </div>
      <!-- Bottom-right (high x, low y) -->
      <div class="quad p-6 border-2" style="background-color: #B5BCC4; border-color: #A1A8B3;">
        <div class="quad-label">{q4_label}</div>
        <div class="mt-8">
          <h3 class="text-xl font-bold text-black mb-3">{q4_title}</h3>
          <ul class="text-sm text-black space-y-2">{q4_items}</ul>
        </div>
      </div>
    </div>
  </div>
</body>
</html>
"""


def render_2x2_matrix(
    output_path: Path,
    title: str,
    quadrants: dict[str, dict[str, Any]],
    subtitle: str = "",
    x_axis_label: str = "",
    x_low: str = "Low",
    x_high: str = "High",
    y_axis_label: str = "",
    y_low: str = "Low",
    y_high: str = "High",
    width: int = 1600,
    height: int = 1000,
) -> Path:
    """2×2 매트릭스(BCG 스타일)를 렌더링.

    Args:
        output_path: PNG 출력 경로
        title: 매트릭스 상단 제목 (액션 타이틀)
        quadrants: {
            "top_left":     {"label": "01", "title": "...", "items": ["..."]},
            "top_right":    {"label": "02", "title": "...", "items": ["..."]},
            "bottom_left":  {"label": "03", "title": "...", "items": ["..."]},
            "bottom_right": {"label": "04", "title": "...", "items": ["..."]},
        }
        subtitle: 부제목
        x/y_axis_label: 축 라벨
        x/y_low/high: 축 끝 라벨
    """

    def _items_html(items: list[str]) -> str:
        return "".join(f"<li>• {item}</li>" for item in items)

    def _quad(key: str) -> dict[str, Any]:
        q = quadrants.get(key, {})
        return {
            "label": q.get("label", ""),
            "title": q.get("title", ""),
            "items_html": _items_html(q.get("items", [])),
        }

    tl = _quad("top_left")
    tr = _quad("top_right")
    bl = _quad("bottom_left")
    br = _quad("bottom_right")

    html = _BCG_TEMPLATE.format(
        title=title,
        subtitle=subtitle,
        x_axis_label=x_axis_label,
        x_low=x_low,
        x_high=x_high,
        y_axis_label=y_axis_label,
        y_low=y_low,
        y_high=y_high,
        q1_label=tl["label"], q1_title=tl["title"], q1_items=tl["items_html"],
        q2_label=tr["label"], q2_title=tr["title"], q2_items=tr["items_html"],
        q3_label=bl["label"], q3_title=bl["title"], q3_items=bl["items_html"],
        q4_label=br["label"], q4_title=br["title"], q4_items=br["items_html"],
    )

    return html_to_png(html, output_path, width=width, height=height)


# ============================================================
# Image Slide PPTX 빌더 (Track A ImageComponent 미구현 우회)
# ============================================================

def build_image_slide_pptx(
    output_pptx: Path,
    image_path: Path,
    title: str = "",
    breadcrumb: str = "",
    footnote: str = "",
) -> Path:
    """이미지 1장을 메인 콘텐츠로 하는 1-슬라이드 pptx를 빌드한다.

    4:3 (10×7.5인치) 슬라이드:
        - 상단 0.5인치: 액션 타이틀 (회사 컬러 검정 18pt 볼드)
        - 우상단: breadcrumb (작은 회색)
        - 본문 영역 0.7~6.7 인치: 이미지 (1:1.6 비율로 폭 자동 맞춤)
        - 하단 0.05인치: 회사 컬러 강조선
        - 최하단: footnote (작은 회색)

    Track A content 렌더러가 ImageComponent를 처리하지 않아 Track C가 직접 빌드.
    """
    output_pptx = Path(output_pptx).resolve()
    image_path = Path(image_path).resolve()
    output_pptx.parent.mkdir(parents=True, exist_ok=True)

    if not image_path.exists():
        raise FileNotFoundError(f"image not found: {image_path}")

    prs = Presentation()
    prs.slide_width = Inches(10)   # 4:3
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]  # 빈 레이아웃
    slide = prs.slides.add_slide(blank_layout)

    accent = RGBColor(0xFD, 0x51, 0x08)
    grey_dark = RGBColor(0xA1, 0xA8, 0xB3)
    black = RGBColor(0x00, 0x00, 0x00)

    # 1. Title
    if title:
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.3), Inches(8.5), Inches(0.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = black

    # 2. Breadcrumb (우상단)
    if breadcrumb:
        bb = slide.shapes.add_textbox(Inches(7.5), Inches(0.2), Inches(2.4), Inches(0.3))
        bf = bb.text_frame
        bp = bf.paragraphs[0]
        bp.text = breadcrumb
        bp.font.size = Pt(9)
        bp.font.color.rgb = grey_dark
        bp.alignment = 2  # right

    # 3. Image (본문)
    img_left = Inches(0.5)
    img_top = Inches(0.95)
    img_width = Inches(9.0)  # 좌우 0.5인치 마진
    slide.shapes.add_picture(
        str(image_path),
        img_left,
        img_top,
        width=img_width,
    )

    # 4. Accent line (하단 구분)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.4),
        Inches(7.05),
        Inches(9.2),
        Inches(0.04),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = accent
    line.line.fill.background()

    # 5. Footnote
    if footnote:
        fb = slide.shapes.add_textbox(Inches(0.4), Inches(7.15), Inches(9.2), Inches(0.3))
        ff = fb.text_frame
        fp = ff.paragraphs[0]
        fp.text = footnote
        fp.font.size = Pt(8)
        fp.font.color.rgb = grey_dark

    prs.save(str(output_pptx))
    return output_pptx
