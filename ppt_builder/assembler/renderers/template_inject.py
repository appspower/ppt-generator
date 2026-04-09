"""템플릿 인젝션 렌더러 - 두 가지 모드 지원:

Track A (기본): template_library.pptx에서 슬라이드를 복제+치환 (SlideCloner)
Track B (직접 편집): 외부 .pptx를 직접 편집하여 관계 100% 보존 (TemplateEditor)

Track B 사용법:
  slide_def.template_name = "ext:smartsheet_general:4"
  → templates/external/smartsheet_general.pptx의 4번 슬라이드를 직접 편집
"""

from pathlib import Path
from pptx.presentation import Presentation
from pptx.slide import Slide

from .base import BaseRenderer
from ...template.cloner import SlideCloner
from ...template.editor import TemplateEditor
from ...template.substitutor import TextSubstitutor
from ...models.schema import TemplateSlide

# Track A: 템플릿 이름 → 슬라이드 인덱스 매핑
TEMPLATE_INDEX = {
    "4col_reference": 0,
    "framework_sidebar": 1,
    "decision_flow": 2,
    "comparison": 3,
    "timeline": 4,
    "before_after": 5,
    "process_grid": 6,
    "kpi_dashboard": 7,
    "swot": 8,
    "hub_spoke": 9,
    "task_image_activity": 10,
    "waterfall": 11,
    "center_focus": 12,
    "dense_table": 13,
    "two_panel": 14,
    "raci": 15,
    "swimlane": 16,
    "pestel": 17,
    "scr": 18,
    "left_right_split": 19,
    "porter_five_forces": 20,
    "value_chain": 21,
    "bcg_matrix": 22,
    "org_chart": 23,
    "gantt_roadmap": 24,
    "prioritization_2x2": 25,
    "tornado": 26,
    "decision_tree": 27,
    "revenue_tree": 28,
    "three_option": 29,
    "circular_loop": 30,
    "mckinsey_7s": 31,
    "three_horizons": 32,
    "mekko": 33,
    "table_with_bars": 34,
}

# Track B: 외부 템플릿 디렉토리
_EXTERNAL_DIR = Path(__file__).parent.parent.parent.parent / "templates" / "external"


def _parse_ext_name(template_name: str) -> tuple[Path, int] | None:
    """ext:파일명:인덱스 형식을 파싱한다.

    Returns:
        (pptx_path, slide_index) 또는 None (ext: 형식이 아닌 경우)
    """
    if not template_name.startswith("ext:"):
        return None

    parts = template_name.split(":")
    if len(parts) != 3:
        raise ValueError(
            f"Invalid ext template format: '{template_name}'. "
            "Expected: 'ext:<filename>:<slide_index>' (e.g., 'ext:smartsheet_general:4')"
        )

    filename = parts[1]
    slide_index = int(parts[2])

    # .pptx 확장자 자동 추가
    if not filename.endswith(".pptx"):
        filename += ".pptx"

    pptx_path = _EXTERNAL_DIR / filename
    if not pptx_path.exists():
        raise FileNotFoundError(
            f"External template not found: {pptx_path}\n"
            f"Available: {[f.name for f in _EXTERNAL_DIR.glob('*.pptx')]}"
        )

    return pptx_path, slide_index


class TemplateInjectRenderer(BaseRenderer):
    """사전 제작 슬라이드를 복제하고 텍스트를 치환한다.

    두 가지 모드:
    - Track A: template_name이 TEMPLATE_INDEX에 있으면 SlideCloner 사용
    - Track B: template_name이 "ext:파일명:인덱스"이면 TemplateEditor 사용
    """

    def __init__(self):
        self._cloner = None
        self._lib_path = Path(__file__).parent.parent.parent.parent / "templates" / "template_library.pptx"

    def _get_cloner(self) -> SlideCloner:
        if self._cloner is None:
            if not self._lib_path.exists():
                raise FileNotFoundError(
                    f"template_library.pptx not found: {self._lib_path}\n"
                    "Run: python -m ppt_builder.template.build_library"
                )
            self._cloner = SlideCloner(self._lib_path)
        return self._cloner

    def render(self, prs: Presentation, slide_def: TemplateSlide) -> Slide:
        template_name = slide_def.template_name

        # Track B: ext:파일명:인덱스
        ext_info = _parse_ext_name(template_name)
        if ext_info is not None:
            return self._render_external(prs, ext_info, slide_def.replacements)

        # Track A: template_library.pptx 기반
        if template_name not in TEMPLATE_INDEX:
            print(f"[WARN] 알 수 없는 템플릿: {template_name}")
            print(f"  사용 가능: {list(TEMPLATE_INDEX.keys())}")
            print(f"  외부 템플릿: ext:<filename>:<slide_index>")
            return self.add_blank_slide(prs)

        idx = TEMPLATE_INDEX[template_name]
        cloner = self._get_cloner()

        new_slide = cloner.clone_slide(prs, idx)

        if slide_def.replacements:
            sub = TextSubstitutor(new_slide)
            sub.replace_all(slide_def.replacements)

        return new_slide

    def _render_external(
        self,
        prs: Presentation,
        ext_info: tuple[Path, int],
        replacements: dict[str, str],
    ) -> Slide:
        """Track B: 외부 .pptx에서 슬라이드를 직접 편집하여 추출한다.

        원본 직접 편집 → 관계 100% 보존 → merge_into로 대상 prs에 추가.
        """
        pptx_path, slide_index = ext_info

        with TemplateEditor(pptx_path) as editor:
            editor.keep_slides([slide_index])

            if replacements:
                editor.substitute(0, replacements)

            # 편집된 슬라이드를 대상 프레젠테이션에 병합
            slides = editor.merge_into(prs)
            return slides[0] if slides else self.add_blank_slide(prs)
