"""템플릿 인젝션 렌더러 - template_library.pptx에서 슬라이드를 복제+치환."""

from pathlib import Path
from pptx.presentation import Presentation
from pptx.slide import Slide

from .base import BaseRenderer
from ...template.cloner import SlideCloner
from ...template.substitutor import TextSubstitutor
from ...models.schema import TemplateSlide

# 템플릿 이름 → 슬라이드 인덱스 매핑
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


class TemplateInjectRenderer(BaseRenderer):
    """사전 제작 슬라이드를 복제하고 텍스트를 치환한다."""

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

        if template_name not in TEMPLATE_INDEX:
            print(f"[WARN] 알 수 없는 템플릿: {template_name}")
            print(f"  사용 가능: {list(TEMPLATE_INDEX.keys())}")
            # Fallback: 빈 슬라이드
            return self.add_blank_slide(prs)

        idx = TEMPLATE_INDEX[template_name]
        cloner = self._get_cloner()

        # 1. 슬라이드 복제
        new_slide = cloner.clone_slide(prs, idx)

        # 2. 텍스트 치환
        if slide_def.replacements:
            sub = TextSubstitutor(new_slide)
            count = sub.replace_all(slide_def.replacements)

        return new_slide
