"""렌더러 기본 클래스 — 완성본 품질 기준 헤더/푸터/브레드크럼."""

from abc import ABC, abstractmethod
from pptx.presentation import Presentation
from pptx.slide import Slide
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from ..styles import (
    SLIDE_WIDTH, SLIDE_HEIGHT,
    TITLE_X, TITLE_Y, TITLE_W, TITLE_H,
    HEADER_X, HEADER_Y, HEADER_W, HEADER_H,
    FOOTER_X, FOOTER_Y, FOOTER_W, FOOTER_H,
    CL_WHITE, CL_BLACK, CL_BODY_TEXT, CL_ACCENT, CL_GREY, CL_GREY_LIGHT,
    CL_DARK, CL_BORDER,
    FONT_TITLE, FONT_BODY,
    FONT_SIZE_TITLE, FONT_SIZE_HEADER, FONT_SIZE_FOOTNOTE,
)


class BaseRenderer(ABC):

    @abstractmethod
    def render(self, prs: Presentation, slide_def) -> Slide:
        pass

    def add_blank_slide(self, prs: Presentation) -> Slide:
        layout = prs.slide_layouts[6]
        return prs.slides.add_slide(layout)

    def add_header_bar(self, slide: Slide) -> None:
        """완성본 스타일: 얇은 다크 헤더 바 (0.55") + 좌측 오렌지 악센트."""
        # 다크 바
        bar = slide.shapes.add_shape(1, 0, 0, SLIDE_WIDTH, Inches(0.55))
        bar.fill.solid()
        bar.fill.fore_color.rgb = CL_DARK
        bar.line.fill.background()
        # 좌측 오렌지 악센트 (3px)
        accent = slide.shapes.add_shape(1, 0, 0, Inches(0.04), Inches(0.55))
        accent.fill.solid()
        accent.fill.fore_color.rgb = CL_ACCENT
        accent.line.fill.background()

    def add_title(self, slide: Slide, text: str) -> None:
        """헤더 바 안에 흰색 제목."""
        txBox = slide.shapes.add_textbox(
            Inches(0.15), Inches(0.08), Inches(6.5), Inches(0.4),
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = CL_WHITE
        p.font.name = FONT_TITLE

    def add_breadcrumb(self, slide: Slide, text: str) -> None:
        """우상단 브레드크럼."""
        if not text:
            return
        bc = slide.shapes.add_textbox(
            Inches(6.8), Inches(0.12), Inches(3.0), Inches(0.3),
        )
        tf = bc.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(7)
        p.font.color.rgb = CL_GREY_LIGHT
        p.font.name = FONT_BODY
        p.alignment = PP_ALIGN.RIGHT

    def add_header_message(self, slide: Slide, text: str) -> None:
        if not text:
            return
        txBox = slide.shapes.add_textbox(
            Inches(0.3), Inches(0.6), Inches(9.4), Inches(0.45),
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(8)
        p.font.color.rgb = CL_BODY_TEXT
        p.font.name = FONT_BODY

    def add_footnote(self, slide: Slide, text: str) -> None:
        if not text:
            return
        # 구분선
        line = slide.shapes.add_shape(
            1, Inches(0.3), Inches(7.15), Inches(9.4), Emu(4572),
        )
        line.fill.solid()
        line.fill.fore_color.rgb = CL_BORDER
        line.line.fill.background()
        # 텍스트
        txBox = slide.shapes.add_textbox(
            Inches(0.3), Inches(7.2), Inches(9.4), Inches(0.2),
        )
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(6)
        p.font.color.rgb = CL_GREY
        p.font.name = FONT_BODY
