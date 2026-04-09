"""PPT Builder - 컨설팅 품질 프레젠테이션 렌더링 라이브러리."""

from pathlib import Path
from .models.schema import PresentationSchema
from .assembler.engine import PresentationEngine


def render_presentation(
    schema: PresentationSchema,
    template: Path | None = None,
    output: Path = Path("output/presentation.pptx"),
) -> Path:
    """
    슬라이드 스키마를 받아서 .pptx 파일을 생성하는 순수 함수.

    Args:
        schema: Pydantic 슬라이드 스키마
        template: .pptx 마스터 템플릿 경로 (None이면 빈 프레젠테이션)
        output: 출력 파일 경로

    Returns:
        생성된 .pptx 파일 경로
    """
    engine = PresentationEngine(template=template)
    return engine.render(schema, output)


# 워크플로 강제 진입점 (Layer 1) — 검증 단계를 빠뜨리지 못하게 함
from .workflow import (  # noqa: E402
    ValidationResult,
    refine_loop,
    render_validated,
    workflow_phases,
)

__all__ = [
    "render_presentation",
    "render_validated",
    "ValidationResult",
    "refine_loop",
    "workflow_phases",
    "PresentationSchema",
]
